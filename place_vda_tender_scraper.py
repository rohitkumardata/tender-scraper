import importlib.util
import os
import re
import sys
import time
import subprocess
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin


def install_missing_packages():
    """Install packages in the SAME Python interpreter used to run this script."""
    required_packages = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "requests": "requests",
        "bs4": "beautifulsoup4",
    }

    missing_packages = [
        pip_name
        for import_name, pip_name in required_packages.items()
        if importlib.util.find_spec(import_name) is None
    ]

    if missing_packages:
        print(f"Installing missing packages: {', '.join(missing_packages)}")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", *missing_packages]
        )
        print("Packages installed successfully!\n")


# Install first, then import third-party packages.
install_missing_packages()

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class TenderScraper:
    def __init__(self):
        # Valle d'Aosta / PLACE-VDA portal
        self.site_root = "https://place-vda.aflink.it"
        self.base_url = f"{self.site_root}/portalegare"
        self.listing_url = f"{self.base_url}/index.php/bandi"

        # Windows-friendly Downloads folder
        self.download_path = Path.home() / "Downloads"

        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/152.0.0.0 Safari/537.36"
                ),
                "Accept": (
                    "text/html,application/xhtml+xml,application/xml;q=0.9,"
                    "image/avif,image/webp,*/*;q=0.8"
                ),
                "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
                "Connection": "keep-alive",
            }
        )

        # Retry temporary HTTP failures automatically.
        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=1,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    @staticmethod
    def _clean_text(value):
        return re.sub(r"\s+", " ", value or "").strip()

    def _absolute_detail_url(self, href, current_page_url):
        """Convert the portal's possible relative detail links to absolute URLs."""
        if not href:
            return None

        href = href.strip()
        if not href or href.lower().startswith("javascript:"):
            return None

        if href.startswith("http://") or href.startswith("https://"):
            return href

        # Common AFLink pattern: href="?RicQ=YES&..."
        if href.startswith("?"):
            return self.listing_url + href

        # Root-relative links, e.g. /portalegare/index.php/bandi?...
        if href.startswith("/"):
            return urljoin(self.site_root, href)

        # Portal-relative links, e.g. index.php/bandi?...
        if href.startswith("index.php"):
            return urljoin(self.base_url + "/", href)

        return urljoin(current_page_url, href)

    def extract_publication_date_from_tender(self, tender_url):
        """
        Extract the TENDER publication date.

        Important for PLACE-VDA:
        the page can contain other 'Data Pubblicazione' values for documents.
        We therefore target the indexing table, identified by fields such as
        'Codice CPV' and 'Data Scadenza Bando'.
        """
        try:
            print(f"      Fetching: {tender_url[:100]}...")
            response = self.session.get(tender_url, timeout=30)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, "html.parser")

            # Method 1: Locate the indexing table and read its Data Pubblicazione column.
            for table in soup.find_all("table"):
                table_text = self._clean_text(table.get_text(" ", strip=True))
                table_lower = table_text.lower()

                if "data pubblicazione" not in table_lower:
                    continue

                # These labels distinguish the tender indexing table from the
                # document table, which can also contain 'Data Pubblicazione'.
                looks_like_index_table = (
                    "codice cpv" in table_lower
                    or "data scadenza bando" in table_lower
                    or "valore importo a base asta" in table_lower
                    or "denominazione dell'amministrazione aggiudicatrice" in table_lower
                )

                if not looks_like_index_table:
                    continue

                rows = table.find_all("tr")
                for row_index, row in enumerate(rows):
                    cells = row.find_all(["th", "td"])
                    texts = [self._clean_text(c.get_text(" ", strip=True)) for c in cells]

                    for col_index, text in enumerate(texts):
                        if "data pubblicazione" not in text.lower():
                            continue

                        # Key/value layout: date may be in a cell after the label.
                        for candidate in texts[col_index + 1 :]:
                            match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", candidate)
                            if match:
                                return match.group(1)

                        # Header/value layout: date is normally in the same column
                        # of the next row.
                        for next_row in rows[row_index + 1 : row_index + 4]:
                            next_cells = next_row.find_all(["th", "td"])
                            next_texts = [
                                self._clean_text(c.get_text(" ", strip=True))
                                for c in next_cells
                            ]
                            if col_index < len(next_texts):
                                match = re.search(
                                    r"\b(\d{2}/\d{2}/\d{4})\b",
                                    next_texts[col_index],
                                )
                                if match:
                                    return match.group(1)

            # Method 2: HTML fallback, but only AFTER the indexing-table marker.
            html = response.text
            lower_html = html.lower()
            marker_positions = [
                lower_html.find("tabella informativa di indicizzazione"),
                lower_html.find("tabella informativa d'indicizzazione"),
                lower_html.find("tabella informativa"),
            ]
            marker_positions = [p for p in marker_positions if p != -1]

            if marker_positions:
                start = min(marker_positions)
                fragment = html[start:]
                fragment_lower = fragment.lower()
                pub_pos = fragment_lower.find("data pubblicazione")

                if pub_pos != -1:
                    snippet = fragment[pub_pos : pub_pos + 5000]
                    match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", snippet)
                    if match:
                        return match.group(1)

            return "Date not found"

        except requests.exceptions.Timeout:
            return "Error: Timeout"
        except requests.exceptions.ConnectionError:
            return "Error: Connection failed"
        except requests.exceptions.HTTPError as e:
            return f"Error: HTTP {e.response.status_code if e.response else 'unknown'}"
        except Exception as e:
            return f"Error: {str(e)[:80]}"

    def get_all_tenders_from_listing(self):
        """Get all current tenders from all available listing pages."""
        all_tenders = []
        seen_urls = set()

        print("=" * 80)
        print("FETCHING PLACE-VDA TENDER LISTINGS")
        print("=" * 80)

        # High safety ceiling. The loop normally stops much earlier when a page
        # returns no new tenders.
        max_pages = 500

        for page_num in range(1, max_pages + 1):
            print(f"\nProcessing page {page_num}...")

            if page_num == 1:
                url = self.listing_url
            else:
                url = f"{self.listing_url}?page={page_num}"

            print(f"  URL: {url}")

            try:
                response = self.session.get(url, timeout=30)
                if response.status_code != 200:
                    print(f"  Page {page_num} returned HTTP {response.status_code}. Stopping.")
                    break

                soup = BeautifulSoup(response.content, "html.parser")

                tenders_on_page = 0
                new_tenders_on_page = 0

                for table in soup.find_all("table"):
                    for row in table.find_all("tr"):
                        cells = row.find_all("td")

                        # Expected PLACE-VDA columns:
                        # Descrizione, Tipo, Ente Proponente, Ente Appaltante,
                        # Importo, Scadenza, Dettaglio
                        if len(cells) < 7:
                            continue

                        description = self._clean_text(cells[0].get_text(" ", strip=True))
                        tipo = self._clean_text(cells[1].get_text(" ", strip=True))
                        ente_proponente = self._clean_text(cells[2].get_text(" ", strip=True))
                        ente_appaltante = self._clean_text(cells[3].get_text(" ", strip=True))
                        importo = self._clean_text(cells[4].get_text(" ", strip=True))
                        scadenza = self._clean_text(cells[5].get_text(" ", strip=True))

                        link = cells[6].find("a", href=True)
                        if not link:
                            # Small fallback in case the detail anchor is elsewhere in the row.
                            possible_links = row.find_all("a", href=True)
                            for possible_link in possible_links:
                                href = possible_link.get("href", "")
                                if "bando=" in href or "getdettaglio" in href.lower():
                                    link = possible_link
                                    break

                        detail_link = None
                        if link:
                            detail_link = self._absolute_detail_url(
                                link.get("href"), response.url
                            )

                        if not description or not detail_link:
                            continue

                        tenders_on_page += 1

                        if detail_link in seen_urls:
                            continue

                        seen_urls.add(detail_link)
                        new_tenders_on_page += 1

                        all_tenders.append(
                            {
                                "description": description,
                                "tipo": tipo,
                                "ente_proponente": ente_proponente,
                                "ente_appaltante": ente_appaltante,
                                "importo": importo,
                                "scadenza": scadenza,
                                "detail_url": detail_link,
                            }
                        )

                print(
                    f"  Parsed {tenders_on_page} tender row(s); "
                    f"{new_tenders_on_page} new tender(s)."
                )

                # Stop safely if the next page is empty or repeats the last page.
                if tenders_on_page == 0:
                    print("\nNo tender rows found. Last page reached.")
                    break

                if page_num > 1 and new_tenders_on_page == 0:
                    print("\nNo new tender URLs found. Last page reached.")
                    break

                time.sleep(0.7)

            except requests.exceptions.RequestException as e:
                print(f"  Request error on page {page_num}: {e}")
                break
            except Exception as e:
                print(f"  Error on page {page_num}: {e}")
                break

        print(f"\nTOTAL UNIQUE TENDERS FOUND: {len(all_tenders)}")
        return all_tenders

    def extract_all_dates(self, tenders):
        """Open each tender and extract its publication date."""
        print("\n" + "=" * 80)
        print("EXTRACTING TENDER PUBLICATION DATES")
        print("=" * 80)
        print(f"This will process {len(tenders)} tenders.\n")

        successful = 0
        failed = 0

        for i, tender in enumerate(tenders, 1):
            description = tender["description"]
            desc_short = description[:70] + "..." if len(description) > 70 else description
            print(f"\n[{i}/{len(tenders)}] {desc_short}")

            pub_date = self.extract_publication_date_from_tender(tender["detail_url"])
            tender["data_pubblicazione"] = pub_date

            if pub_date != "Date not found" and not pub_date.startswith("Error"):
                successful += 1
                print(f"  Data Pubblicazione: {pub_date}")
            else:
                failed += 1
                print(f"  {pub_date}")

            time.sleep(0.5)

        print("\n" + "=" * 80)
        print("DATE EXTRACTION SUMMARY")
        print("=" * 80)
        print(f"Successfully extracted: {successful}")
        print(f"Failed to extract: {failed}")
        if tenders:
            print(f"Success rate: {successful / len(tenders) * 100:.1f}%")

        return tenders

    @staticmethod
    def _parse_date_for_sorting(date_str):
        if date_str and isinstance(date_str, str):
            match = re.search(r"(\d{2}/\d{2}/\d{4})", date_str)
            if match:
                try:
                    return datetime.strptime(match.group(1), "%d/%m/%Y")
                except ValueError:
                    pass
        return datetime(1900, 1, 1)

    def save_to_excel(self, tenders):
        """Save the same output structure as the user's current SoReSa scraper."""
        self.download_path.mkdir(parents=True, exist_ok=True)

        if not tenders:
            print("No tenders to save.")
            return None

        data = []
        for i, tender in enumerate(tenders, 1):
            data.append(
                {
                    "S.No": i,
                    "Tender Description": tender["description"],
                    "Tipo": tender["tipo"],
                    "Ente Proponente": tender["ente_proponente"],
                    "Ente Appaltante": tender["ente_appaltante"],
                    "Importo": tender["importo"],
                    "Scadenza": tender["scadenza"],
                    "Data Pubblicazione": tender.get("data_pubblicazione", ""),
                    "Tender URL": tender["detail_url"],
                }
            )

        df = pd.DataFrame(data)

        df_sorted = df.copy()
        df_sorted["_sort_date"] = df_sorted["Data Pubblicazione"].apply(
            self._parse_date_for_sorting
        )
        df_sorted = (
            df_sorted.sort_values("_sort_date", ascending=False)
            .drop("_sort_date", axis=1)
            .reset_index(drop=True)
        )

        valid_dates = (
            df["Data Pubblicazione"]
            .fillna("")
            .astype(str)
            .str.match(r"^\d{2}/\d{2}/\d{4}$")
        )
        date_counts = (
            df.loc[valid_dates, "Data Pubblicazione"]
            .value_counts()
            .rename_axis("Data Pubblicazione")
            .reset_index(name="Count")
        )

        if not date_counts.empty:
            date_counts["_sort_date"] = date_counts["Data Pubblicazione"].apply(
                self._parse_date_for_sorting
            )
            date_counts = (
                date_counts.sort_values("_sort_date", ascending=False)
                .drop("_sort_date", axis=1)
                .reset_index(drop=True)
            )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"place_vda_all_tenders_with_dates_{timestamp}.xlsx"
        filepath = self.download_path / filename

        try:
            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="All Tenders", index=False)
                df_sorted.to_excel(writer, sheet_name="Sorted by Date", index=False)
                date_counts.to_excel(writer, sheet_name="Summary by Date", index=False)

                for worksheet in writer.sheets.values():
                    worksheet.freeze_panes = "A2"
                    worksheet.auto_filter.ref = worksheet.dimensions

                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter

                        for cell in column:
                            try:
                                if cell.value is not None:
                                    max_length = max(max_length, len(str(cell.value)))
                            except Exception:
                                pass

                        worksheet.column_dimensions[column_letter].width = min(
                            max_length + 2, 80
                        )

            print("\nExcel file saved successfully!")
            print(f"Location: {filepath}")

            txt_filename = f"place_vda_all_tenders_urls_{timestamp}.txt"
            txt_filepath = self.download_path / txt_filename

            with open(txt_filepath, "w", encoding="utf-8") as f:
                f.write("PLACE-VDA - ALL TENDERS WITH PUBLICATION DATES\n")
                f.write("=" * 100 + "\n\n")

                for tender in tenders:
                    f.write(f"Date: {tender.get('data_pubblicazione', '')}\n")
                    f.write(f"URL: {tender['detail_url']}\n")
                    f.write(f"Description: {tender['description']}\n")
                    f.write("-" * 100 + "\n")

            print(f"Text file saved: {txt_filepath}")
            return str(filepath)

        except Exception as e:
            print(f"Error saving Excel: {e}")
            csv_path = str(filepath).replace(".xlsx", ".csv")
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"Saved as CSV instead: {csv_path}")
            return csv_path

    def display_summary(self, tenders):
        """Display publication-date summary in the PyCharm console."""
        print("\n" + "=" * 80)
        print("FINAL SUMMARY")
        print("=" * 80)

        date_counts = {}
        for tender in tenders:
            date_value = tender.get("data_pubblicazione", "")
            if (
                date_value
                and date_value != "Date not found"
                and not str(date_value).startswith("Error")
            ):
                date_counts[date_value] = date_counts.get(date_value, 0) + 1

        sorted_dates = sorted(
            date_counts.items(),
            key=lambda item: self._parse_date_for_sorting(item[0]),
            reverse=True,
        )

        if sorted_dates:
            print("\nTenders by Publication Date:")
            print("-" * 50)
            for date_value, count in sorted_dates:
                print(f"  {date_value}: {count} tender(s)")

            most_recent = sorted_dates[0][0]
            print(f"\nSample tenders from {most_recent}:")
            print("-" * 50)

            examples = [
                t for t in tenders if t.get("data_pubblicazione") == most_recent
            ][:5]
            for i, tender in enumerate(examples, 1):
                print(f"  {i}. {tender['description'][:90]}")
                print(f"     {tender['detail_url']}")


def main():
    print("=" * 80)
    print("PLACE-VDA TENDER SCRAPER - ALL TENDERS WITH PUBLICATION DATES")
    print("=" * 80)
    print("\nSource:")
    print("  https://place-vda.aflink.it/portalegare/index.php/bandi")
    print("\nThis script will:")
    print("  1. Fetch tender rows from every listing page")
    print("  2. Open each tender detail page")
    print("  3. Extract the tender's Data Pubblicazione from the indexing table")
    print("  4. Save Excel + TXT output in your Windows Downloads folder")
    print(f"\nDownload path: {Path.home() / 'Downloads'}")

    input("\nPress Enter to start scraping...")

    scraper = TenderScraper()

    all_tenders = scraper.get_all_tenders_from_listing()
    if not all_tenders:
        print("\nNo tenders found. See troubleshooting notes below/ask me with the console output.")
        return

    tenders_with_dates = scraper.extract_all_dates(all_tenders)
    scraper.save_to_excel(tenders_with_dates)
    scraper.display_summary(tenders_with_dates)

    print("\n" + "=" * 80)
    print("PROCESS COMPLETED!")
    print("=" * 80)
    print("\nOpen the Excel file from Downloads and use:")
    print("  - All Tenders")
    print("  - Sorted by Date")
    print("  - Summary by Date")


if __name__ == "__main__":
    main()
