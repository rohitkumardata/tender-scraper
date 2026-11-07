"""PLACE-VDA scraper for the new /Portale/index.asp portal.

Run in PyCharm, or: python place_vda_updated_scraper.py
Default: current/non-expired Forniture, every page, with tender-detail CPVs.
Excel, CSV and TXT are written to Downloads. No Selenium/browser required.
Only the Excel export needs pandas/openpyxl (installed if missing).
"""

import argparse
import csv
import html
import importlib.util
import json
import re
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from html.parser import HTMLParser
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin
from urllib.request import HTTPCookieProcessor, Request, build_opener

HOME = "https://place-vda.aflink.it/Portale/index.asp"
KEY = "BandiPubblici"
FIELDS = ["S.No", "Notice ID", "Tender Description", "Tipo", "Contract Type",
          "Ente Proponente", "Ente Appaltante", "Importo", "Scadenza",
          "Data Pubblicazione", "CPV Codes", "CPV Details", "CIG", "Tender URL",
          "Listing Source", "Detail Data Source", "Extraction Status", "Error"]


class TextOnly(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)


def clean(value):
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    text = str(value)
    if "<" in text:
        parser = TextOnly()
        parser.feed(text)
        text = " ".join(parser.parts)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def cpv_codes(value):
    # Search ONLY the CPV field, never the whole page (dates/IDs are numbers too).
    return "; ".join(dict.fromkeys(re.findall(r"(?<!\d)\d{8}(?:-\d)?(?!\d)", clean(value))))


def format_deadline(value):
    value = clean(value)
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value


class PortalClient:
    def __init__(self, api, timeout=45):
        self.api = api.rstrip("/") + "/"
        self.timeout = timeout
        self.opener = build_opener(HTTPCookieProcessor(CookieJar()))

    def fetch(self, url, payload=None, is_json=True):
        body = json.dumps(payload).encode("utf-8") if payload is not None else None
        for attempt in range(3):
            try:
                headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json" if is_json else "text/html"}
                if body is not None:
                    headers["Content-Type"] = "application/json"
                with self.opener.open(Request(url, data=body, headers=headers), timeout=self.timeout) as response:
                    text = response.read().decode("utf-8-sig")
                result = json.loads(text) if is_json else text
                if is_json and isinstance(result, dict) and (result.get("error") or result.get("status") is False):
                    raise ValueError("Portal error: " + clean(result.get("message") or result.get("error")))
                return result
            except HTTPError as exc:
                # Do not keep retrying access denials or missing endpoints.
                if exc.code not in (429, 500, 502, 503, 504) or attempt == 2:
                    raise
                time.sleep(min(int(exc.headers.get("Retry-After", "2")) if exc.headers.get("Retry-After", "").isdigit() else 2, 30))
            except (URLError, TimeoutError, ConnectionError):
                if attempt == 2:
                    raise
                time.sleep(2 ** attempt)
        raise RuntimeError("Request retries exhausted")

    def call(self, endpoint, payload=None, **query):
        url = self.api + endpoint
        if query:
            url += "?" + urlencode(query)
        return self.fetch(url, payload)


def discover_api():
    client = PortalClient("")
    page = client.fetch(HOME, is_json=False)
    version = re.search(r"(?:const|var|let)\s+apiVersion\s*=\s*[`'\"]([^`'\"]+)", page)
    server = re.search(r"(?:const|var|let)\s+WebApiServer\s*=\s*[`'\"]([^`'\"]+)", page)
    if not version or not server:
        raise RuntimeError("Cannot identify the portal API configuration. The homepage may have changed.")
    return urljoin(HOME, server.group(1).rstrip("/") + "/api/" + version.group(1) + "/Rss/")


def get_listing(client, category, expired):
    columns = client.call("getRssColumns", type=category, key=KEY)
    if not isinstance(columns, list) or not all(isinstance(c, dict) and "prop" in c for c in columns):
        raise RuntimeError("Unexpected listing column format")
    props = [c["prop"] for c in columns]
    if not {"IdMsg", "Oggetto", "Tipo"}.issubset(props):
        raise RuntimeError("Required listing columns are missing")
    records, seen, total, fetched = [], set(), None, 0
    page_size = 20
    for page in range(1, 501):
        result = client.call("GetRss", {"key": KEY, "type": category, "scaduti": expired}, page=page, pageSize=page_size)
        if not isinstance(result, dict) or not isinstance(result.get("data"), list) or "recordCount" not in result:
            raise RuntimeError("Unexpected listing response; stopping instead of saving an empty result")
        current_total = int(result["recordCount"])
        if total is None:
            total = current_total
        elif total != current_total:
            raise RuntimeError("Listing changed during extraction. Run again for a consistent export.")
        rows = result["data"]
        if not rows and fetched < total:
            raise RuntimeError("Listing stopped before its reported total")
        added = 0
        for values in rows:
            if not isinstance(values, list) or len(values) != len(props):
                raise RuntimeError("Listing row does not match its column definitions")
            item = dict(zip(props, values))
            identity = clean(item["IdMsg"])
            if not identity:
                raise RuntimeError("Listing row has no notice ID")
            if identity not in seen:
                seen.add(identity)
                records.append(item)
                added += 1
        fetched += len(rows)
        print(f"Listing page {page}: {len(rows)} rows; {len(records)}/{total} unique notices", flush=True)
        if fetched >= total:
            break
        if not added:
            raise RuntimeError("Pagination repeated a page; extraction is incomplete")
        time.sleep(0.3)
    else:
        raise RuntimeError("Pagination safety limit reached")
    if len(records) != total:
        raise RuntimeError(f"Expected {total} unique notices, received {len(records)}. Run again.")
    return records


def extract_detail(api, item, category):
    identity = clean(item["IdMsg"])
    detail_url = HOME + "?" + urlencode({"ID": identity})
    data_url = api + "getRssDpcm"
    record = {field: "" for field in FIELDS}
    record.update({"Notice ID": identity, "Tender Description": clean(item.get("Oggetto")),
                   "Tipo": clean(item.get("Tipo")), "Contract Type": category,
                   "Ente Proponente": clean(item.get("EnteProponente")),
                   "Ente Appaltante": clean(item.get("DenominazioneEnte")),
                   "Importo": clean(item.get("a_base_asta")),
                   "Scadenza": format_deadline(item.get("DtScadenzaBandoTecnical")),
                   "Tender URL": detail_url, "Listing Source": HOME,
                   "Detail Data Source": data_url})
    try:
        # This is the exact detail/indexing-table request made when opening a notice.
        # It returns CPV, publication date and CIG without unrelated attachment dates.
        result = PortalClient(api).call("getRssDpcm", {"key": KEY, "idBando": identity, "type": item["Tipo"]})
        details = result.get("data") if isinstance(result, dict) else None
        if not isinstance(details, dict) or not details:
            raise ValueError("Tender indexing details unavailable")
        normalized = {str(k).casefold(): v for k, v in details.items()}
        cpv = clean(normalized.get("cpv"))
        codes = cpv_codes(cpv)
        record.update({"Contract Type": clean(details.get("Contratto")) or category,
                       "Data Pubblicazione": clean(details.get("DtPubblicazione")),
                       "CPV Codes": codes, "CPV Details": cpv,
                       "CIG": clean(normalized.get("cig"))})
        issues = []
        if not cpv:
            issues.append("CPV not provided in tender details")
        elif not codes:
            issues.append("CPV field present but code format not recognized")
        if not record["Data Pubblicazione"]:
            issues.append("Publication date not provided in tender details")
        record["Extraction Status"] = "Review" if issues else "OK"
        record["Error"] = "; ".join(issues)
    except Exception as exc:
        record["Extraction Status"] = "Detail fetch failed"
        record["Error"] = f"{type(exc).__name__}: {exc}"
    return record


def safe_cell(value):
    # Preserve CPV identifiers as text and prevent scraped text becoming formulas.
    value = clean(value)
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def write_csv(records, path):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows({k: safe_cell(row.get(k, "")) for k in FIELDS} for row in records)


def save_excel(records, path):
    missing = [name for name in ("pandas", "openpyxl") if importlib.util.find_spec(name) is None]
    if missing:
        print("Installing Excel packages in this Python interpreter: " + ", ".join(missing), flush=True)
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
    import pandas as pd
    df = pd.DataFrame([{k: safe_cell(r.get(k, "")) for k in FIELDS} for r in records], columns=FIELDS)
    dates = pd.to_datetime(df["Data Pubblicazione"], format="%d/%m/%Y", errors="coerce")
    ordered = df.assign(_date=dates).sort_values("_date", ascending=False, na_position="last").drop(columns="_date")
    summary = dates.dropna().value_counts().sort_index(ascending=False).rename_axis("Data Pubblicazione").reset_index(name="Count")
    summary["Data Pubblicazione"] = summary["Data Pubblicazione"].dt.strftime("%d/%m/%Y")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Tenders", index=False)
        ordered.to_excel(writer, sheet_name="Sorted by Date", index=False)
        summary.to_excel(writer, sheet_name="Summary by Date", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col) + 2, 65)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--category", choices=["Forniture", "Tutti", "Servizi", "Lavori", "Esito"], default="Forniture")
    parser.add_argument("--expired", action="store_true", help="Scrape the portal's expired view instead of its current view")
    parser.add_argument("--output-dir", type=Path, default=Path.home() / "Downloads")
    parser.add_argument("--workers", type=int, choices=range(1, 5), default=3)
    parser.add_argument("--csv-only", action="store_true", help="Skip Excel export and package installation")
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    prefix = args.output_dir / ("PLACE_VDA_" + args.category + "_" + ("expired_" if args.expired else "active_") + datetime.now().strftime("%Y%m%d_%H%M%S"))
    csv_path = prefix.with_suffix(".csv")
    api = discover_api()
    print("Source: " + HOME)
    print(f"Scope: {args.category}; {'expired' if args.expired else 'current/non-expired'}", flush=True)
    items = get_listing(PortalClient(api), args.category, args.expired)
    if not items:
        print("The portal returned zero notices for this filter.")
        return 0
    results = {}
    interrupted = False
    try:
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            pending = {pool.submit(extract_detail, api, item, args.category): i for i, item in enumerate(items)}
            for future in as_completed(pending):
                index = pending[future]
                row = future.result()
                row["S.No"] = index + 1
                results[index] = row
                print(f"Details {len(results)}/{len(items)}: {row['Notice ID']} | CPV {row['CPV Codes'] or '(missing)'} | {row['Extraction Status']}", flush=True)
                # A checkpoint after each tender protects completed work.
                write_csv([results[i] for i in sorted(results)], csv_path)
    except KeyboardInterrupt:
        interrupted = True
        print("Interrupted. Saving completed notices.")
    records = [results[i] for i in sorted(results)]
    if not records:
        print("No detail records completed.")
        return 1
    write_csv(records, csv_path)
    prefix.with_suffix(".txt").write_text("\n\n".join(f"{r['Notice ID']} | {r['Tender Description']}\n{r['Tender URL']}\nCPV: {r['CPV Details']}\nPublication: {r['Data Pubblicazione']}\nStatus: {r['Extraction Status']} {r['Error']}" for r in records), encoding="utf-8")
    print("Saved CSV: " + str(csv_path))
    if not args.csv_only:
        try:
            save_excel(records, prefix.with_suffix(".xlsx"))
            print("Saved Excel: " + str(prefix.with_suffix(".xlsx")))
        except Exception as exc:
            print(f"Excel export failed: {exc}. Your CSV and TXT were saved.")
    failed = sum(r["Extraction Status"] == "Detail fetch failed" for r in records)
    review = sum(r["Extraction Status"] == "Review" for r in records)
    print(f"Completed {len(records)}/{len(items)} notices; detail failures: {failed}; review: {review}.")
    return 1 if interrupted or failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:
        print(f"ERROR: {type(error).__name__}: {error}", file=sys.stderr)
        sys.exit(1)
