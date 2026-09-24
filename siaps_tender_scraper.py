import os
import re
import sys
import time
import subprocess
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup


class TenderScraper:
    def __init__(self):
        self.base_url = "https://siaps.soresa.it/portalegare"
        self.listing_url = f"{self.base_url}/index.php/bandi"

        # Windows-friendly Downloads path
        self.download_path = Path.home() / "Downloads"

        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
        })

    def extract_publication_date_from_tender(self, tender_url):
        """Extract Data Pubblicazione from individual tender page."""
        try:
            print(f"      Fetching: {tender_url[:80]}...")
            response = self.session.get(tender_url, timeout=30)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            # Method 1: Look for "Data Pubblicazione" in tables
            all_tables = soup.find_all("table")
            for table in all_tables:
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all(["th", "td"])
                    row_text = " ".join([c.get_text(strip=True) for c in cells])

                    if "Data Pubblicazione" in row_text or "Data di Pubblicazione" in row_text:
                        for cell in cells:
                            cell_text = cell.get_text(strip=True)
                            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", cell_text)
                            if date_match:
                                return date_match.group(1)

                        next_row = row.find_next_sibling("tr")
                        if next_row:
                            next_cells = next_row.find_all("td")
                            for cell in next_cells:
                                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", cell.get_text())
                                if date_match:
                                    return date_match.group(1)

            # Method 2: Look for specific patterns
            patterns = [
                r"Data\s+Pubblicazione\s*:\s*(\d{2}/\d{2}/\d{4})",
                r"Data\s+Pubblicazione\s*</th>\s*<td[^>]*>\s*(\d{2}/\d{2}/\d{4})",
                r"<th[^>]*>Data\s+Pubblicazione</th>\s*<td[^>]*>(\d{2}/\d{2}/\d{4})",
                r"Data\s+Pubblicazione\s*</td>\s*<td[^>]*>\s*(\d{2}/\d{2}/\d{4})",
                r"Pubblicato\s+il\s*:\s*(\d{2}/\d{2}/\d{4})",
            ]

            for pattern in patterns:
                match = re.search(pattern, response.text, re.IGNORECASE | re.DOTALL)
                if match:
                    date_value = match.group(1)
                    if re.match(r"\d{2}/\d{2}/\d{4}", date_value):
                        return date_value

            # Method 3: Look for any date that appears near "Data Pubblicazione"
            text = response.text
            pub_index = text.lower().find("data pubblicazione")
            if pub_index != -1:
                snippet = text[pub_index:pub_index + 300]
                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", snippet)
                if date_match:
                    return date_match.group(1)

            # Method 4: Look for the specific table structure
            if "TABELLA INFORMATIVA DI INDICIZZAZIONE" in text or "TABELLA INFORMATIVA" in text:
                all_dates = re.findall(r"(\d{2}/\d{2}/\d{4})", text)
                if all_dates:
                    return all_dates[0]

            return "Date not found"

        except requests.exceptions.Timeout:
            return "Error: Timeout"
        except requests.exceptions.ConnectionError:
            return "Error: Connection failed"
        except Exception as e:
            return f"Error: {str(e)[:50]}"

    def get_all_tenders_from_listing(self):
        """Get all tenders from all pages of the listing."""
        all_tenders = []
        page_num = 1
        seen_urls = set()

        print("=" * 80)
        print("FETCHING TENDER LISTINGS")
        print("=" * 80)

        while page_num <= 20:  # Safety limit
            print(f"\nProcessing page {page_num}...")
            try:
                if page_num == 1:
                    url = self.listing_url
                else:
                    url = f"{self.listing_url}?page={page_num}"

                print(f"  URL: {url}")
                response = self.session.get(url, timeout=20)

                if response.status_code != 200:
                    print(f"  Page {page_num} not accessible")
                    break

                soup = BeautifulSoup(response.text, "html.parser")

                tables = soup.find_all("table")
                tenders_on_page = 0

                for table in tables:
                    rows = table.find_all("tr")
                    for row in rows:
                        cells = row.find_all("td")
                        if len(cells) >= 6:
                            description = cells[0].get_text(strip=True) if len(cells) > 0 else ""
                            tipo = cells[1].get_text(strip=True) if len(cells) > 1 else ""
                            ente_proponente = cells[2].get_text(strip=True) if len(cells) > 2 else ""
                            ente_appaltante = cells[3].get_text(strip=True) if len(cells) > 3 else ""
                            importo = cells[4].get_text(strip=True) if len(cells) > 4 else ""
                            scadenza = cells[5].get_text(strip=True) if len(cells) > 5 else ""

                            detail_link = None
                            if len(cells) > 6:
                                link = cells[6].find("a")
                                if link and link.get("href"):
                                    detail_link = urljoin(self.base_url, link.get("href"))

                            if detail_link and description and detail_link not in seen_urls:
                                seen_urls.add(detail_link)
                                all_tenders.append({
                                    "description": description,
                                    "tipo": tipo,
                                    "ente_proponente": ente_proponente,
                                    "ente_appaltante": ente_appaltante,
                                    "importo": importo,
                                    "scadenza": scadenza,
                                    "detail_url": detail_link,
                                })
                                tenders_on_page += 1

                print(f"  Found {tenders_on_page} tenders on page {page_num}")

                pagination_div = soup.find("div", id=re.compile(r"paginazione", re.I))
                if pagination_div:
                    current_page = pagination_div.find("span", class_=re.compile(r"active|current", re.I))
                    if current_page:
                        next_page_link = current_page.find_next_sibling("a")
                        if not next_page_link:
                            print(f"\nLast page reached (page {page_num})")
                            break
                else:
                    if tenders_on_page < 20 and page_num > 1:
                        print(f"\nLast page reached (only {tenders_on_page} tenders on page {page_num})")
                        break

                page_num += 1
                time.sleep(1)

            except Exception as e:
                print(f"  Error on page {page_num}: {e}")
                break

        print(f"\nTOTAL TENDERS FOUND: {len(all_tenders)}")
        return all_tenders

    def extract_all_dates(self, tenders):
        """Extract publication dates for all tenders."""
        print("\n" + "=" * 80)
        print("EXTRACTING PUBLICATION DATES")
        print("=" * 80)
        print(f"This will process {len(tenders)} tenders. This may take several minutes...\n")

        successful = 0
        failed = 0

        for i, tender in enumerate(tenders, 1):
            desc_short = (
                tender["description"][:60] + "..."
                if len(tender["description"]) > 60
                else tender["description"]
            )
            print(f"\n[{i}/{len(tenders)}] {desc_short}")

            pub_date = self.extract_publication_date_from_tender(tender["detail_url"])
            tender["data_pubblicazione"] = pub_date

            if pub_date != "Date not found" and not pub_date.startswith("Error"):
                successful += 1
                print(f"  Data Pubblicazione: {pub_date}")
            else:
                failed += 1
                print(f"  {pub_date}")

            time.sleep(1)

        print(f"\n{'=' * 80}")
        print("DATE EXTRACTION SUMMARY")
        print(f"{'=' * 80}")
        print(f"Successfully extracted: {successful}")
        print(f"Failed to extract: {failed}")
        if len(tenders) > 0:
            print(f"Success rate: {successful / len(tenders) * 100:.1f}%")

        return tenders

    def save_to_excel(self, tenders):
        """Save all tenders with their publication dates to Excel."""
        self.download_path.mkdir(parents=True, exist_ok=True)

        if not tenders:
            print("No tenders to save")
            return None

        data = []
        for i, tender in enumerate(tenders, 1):
            data.append({
                "S.No": i,
                "Tender Description": tender["description"],
                "Tipo": tender["tipo"],
                "Ente Proponente": tender["ente_proponente"],
                "Ente Appaltante": tender["ente_appaltante"],
                "Importo": tender["importo"],
                "Scadenza": tender["scadenza"],
                "Data Pubblicazione": tender.get("data_pubblicazione", ""),
                "Tender URL": tender["detail_url"],
            })

        df = pd.DataFrame(data)

        def parse_date(date_str):
            if date_str and isinstance(date_str, str):
                match = re.search(r"(\d{2}/\d{2}/\d{4})", date_str)
                if match:
                    try:
                        return datetime.strptime(match.group(1), "%d/%m/%Y")
                    except Exception:
                        pass
            return datetime(1900, 1, 1)

        df_sorted = df.copy()
        df_sorted["_sort_date"] = df_sorted["Data Pubblicazione"].apply(parse_date)
        df_sorted = df_sorted.sort_values("_sort_date", ascending=False).drop("_sort_date", axis=1)

        valid_dates = df["Data Pubblicazione"].fillna("").astype(str).str.match(r"\d{2}/\d{2}/\d{4}")
        date_counts = df.loc[valid_dates, "Data Pubblicazione"].value_counts().reset_index()

        if not date_counts.empty:
            date_counts.columns = ["Data Pubblicazione", "Count"]
            date_counts = date_counts.sort_values("Data Pubblicazione")
        else:
            date_counts = pd.DataFrame(columns=["Data Pubblicazione", "Count"])

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"all_tenders_with_dates_{timestamp}.xlsx"
        filepath = self.download_path / filename

        try:
            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="All Tenders", index=False)
                df_sorted.to_excel(writer, sheet_name="Sorted by Date", index=False)
                date_counts.to_excel(writer, sheet_name="Summary by Date", index=False)

                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if cell.value and len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except Exception:
                                pass
                        adjusted_width = min(max_length + 2, 80)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

            print("\nExcel file saved successfully!")
            print(f"Location: {filepath}")

            txt_filename = f"all_tenders_urls_{timestamp}.txt"
            txt_filepath = self.download_path / txt_filename

            with open(txt_filepath, "w", encoding="utf-8") as f:
                f.write("ALL TENDERS WITH PUBLICATION DATES\n")
                f.write("=" * 100 + "\n\n")
                for tender in tenders:
                    f.write(f"Date: {tender.get('data_pubblicazione', '')}\n")
                    f.write(f"URL: {tender['detail_url']}\n")
                    f.write(f"Description: {tender['description'][:100]}\n")
                    f.write("-" * 100 + "\n")

            print(f"Text file saved: {txt_filepath}")
            return str(filepath)

        except Exception as e:
            print(f"Error saving Excel: {e}")
            csv_path = str(filepath).replace(".xlsx", ".csv")
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"Saved as CSV: {csv_path}")
            return csv_path

    def display_summary(self, tenders):
        """Display summary of results."""
        print("\n" + "=" * 80)
        print("FINAL SUMMARY")
        print("=" * 80)

        date_counts = {}
        for tender in tenders:
            date_value = tender.get("data_pubblicazione", "")
            if date_value and date_value != "Date not found" and not date_value.startswith("Error"):
                date_counts[date_value] = date_counts.get(date_value, 0) + 1

        sorted_dates = []
        if date_counts:
            print("\nTenders by Publication Date:")
            print("-" * 50)

            sorted_dates = sorted(
                date_counts.items(),
                key=lambda x: datetime.strptime(x[0], "%d/%m/%Y")
                if re.match(r"\d{2}/\d{2}/\d{4}", x[0])
                else datetime(1900, 1, 1),
                reverse=True,
            )

            for date_value, count in sorted_dates:
                print(f"  {date_value}: {count} tender(s)")

        if sorted_dates:
            most_recent = sorted_dates[0][0]
            print(f"\nSample tenders from {most_recent}:")
            print("-" * 50)
            tenders_for_date = [t for t in tenders if t.get("data_pubblicazione") == most_recent]
            for i, tender in enumerate(tenders_for_date[:5], 1):
                url_short = (
                    tender["detail_url"][:80] + "..."
                    if len(tender["detail_url"]) > 80
                    else tender["detail_url"]
                )
                print(f"  {i}. {url_short}")


def install_missing_packages():
    required_packages = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "requests": "requests",
        "bs4": "beautifulsoup4",
    }

    missing_packages = []

    for import_name, pip_name in required_packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing_packages.append(pip_name)

    if missing_packages:
        print(f"Installing missing packages: {', '.join(missing_packages)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])
        print("Packages installed successfully!\n")


def main():
    print("=" * 80)
    print("TENDER SCRAPER - EXTRACT ALL TENDERS WITH PUBLICATION DATES")
    print("=" * 80)
    print("\nThis script will:")
    print("  1. Fetch ALL tenders from ALL pages of the portal")
    print("  2. Open EACH tender page to extract 'Data Pubblicazione'")
    print("  3. Save all results to Excel with publication dates")
    print(f"\nDownload path: {Path.home() / 'Downloads'}")
    print("\nThis may take several minutes for many tenders...")

    input("\nPress Enter to start scraping...")

    scraper = TenderScraper()

    all_tenders = scraper.get_all_tenders_from_listing()
    if not all_tenders:
        print("\nNo tenders found.")
        return

    tenders_with_dates = scraper.extract_all_dates(all_tenders)
    scraper.save_to_excel(tenders_with_dates)
    scraper.display_summary(tenders_with_dates)

    print("\n" + "=" * 80)
    print("PROCESS COMPLETED!")
    print("=" * 80)
    print("\nYou can now open the Excel file and:")
    print("  - Filter by 'Data Pubblicazione' column")
    print("  - Sort to see newest tenders first")
    print("  - Check the 'Summary by Date' sheet")


if __name__ == "__main__":
    install_missing_packages()
    main()
