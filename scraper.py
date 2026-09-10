"""
Nium Playbook — Scraper & Data Layer (no Streamlit dependency)

Provides:
  scrape_dataset()        — Selenium + BeautifulSoup scrape for one dataset type
  transform_raw_to_wide() — Pivot long-format rows into wide DataFrame
  save_scraped_data()     — Write to data/ and scraped_data/ archive
  log_scrape_failures()   — Append failed-country details to logs/scrape_failures.log
  get_last_updated()      — Timestamp of most recent data file
"""

import json
import os
import re
import time
import unicodedata
from io import BytesIO
from pathlib import Path
from datetime import datetime

import pandas as pd

from config import SCRAPE_PAGE_TIMEOUT_SECONDS, SCRAPE_MAX_RETRIES, SCRAPE_RETRY_DELAY_SECONDS
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ─── Paths ───────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
DATA_DIR    = BASE_DIR / "data"
ARCHIVE_DIR = BASE_DIR / "scraped_data"
DATA_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR.mkdir(exist_ok=True)

FI_PATH     = DATA_DIR / "FI_data.xlsx"
NON_FI_PATH = DATA_DIR / "Non_FI_data.xlsx"

# ─── Country list (loaded from countries.txt — edit that file to add/remove) ─
_countries_file = BASE_DIR / "countries.txt"
COUNTRIES = [
    line.strip()
    for line in _countries_file.read_text(encoding="utf-8").splitlines()
    if line.strip()
]

# ─── URL slug overrides (countries whose Nium slug ≠ auto-derived slug) ──────
# Edit country_url_overrides.json to fix a country that shows up in "failed".
_overrides_file = BASE_DIR / "country_url_overrides.json"
URL_OVERRIDES: dict = json.loads(_overrides_file.read_text(encoding="utf-8")) if _overrides_file.exists() else {}

# ─── Internal helpers ─────────────────────────────────────────────────────────

def _format_country_url(name: str) -> str:
    """Convert a country display name to its Nium Playbook URL slug.

    Checks URL_OVERRIDES first for any country whose slug doesn't match the
    auto-derived form, then falls back to: strip accents → lowercase →
    remove punctuation → collapse spaces/hyphens.
    """
    if name in URL_OVERRIDES:
        return URL_OVERRIDES[name]
    # Decompose accented chars: ç→c, é→e, ü→u, ñ→n, etc.
    nfkd = unicodedata.normalize("NFKD", name.strip())
    ascii_name = nfkd.encode("ascii", "ignore").decode("ascii")
    slug = ascii_name.lower()
    # Strip everything except letters, digits, spaces, hyphens
    slug = re.sub(r"[^a-z0-9\s-]", "", slug)
    # Collapse runs of whitespace/hyphens into a single hyphen
    slug = re.sub(r"[\s-]+", "-", slug)
    return slug.strip("-")


# ─── Scraper ──────────────────────────────────────────────────────────────────

def _parse_method_page(soup, country_name, method_label):
    """Parse a single atlas.nium.com method page and return raw_rows."""
    rows = []

    # ── Currencies ───────────────────────────────────────────────────────────
    currencies = [
        span.get("data-atlas-currency", "")
        for span in soup.find_all("span", attrs={"data-atlas-currency": True})
    ]
    currency_str = ", ".join(c for c in currencies if c)

    # ── TAT (speed badge + description) ──────────────────────────────────────
    tat = ""
    tat_details = ""
    timing_row = soup.find("div", class_=lambda c: c and "figma-route-summary__row--timing" in c)
    if timing_row:
        timing_div = timing_row.find("div", class_=lambda c: c and "figma-route-summary__timing" in c)
        if timing_div:
            speed_span = timing_div.find("span")
            if speed_span:
                tat = speed_span.get_text(strip=True)
            details_div = timing_div.find("div", class_=lambda c: c and "atlas-timing-details" in c)
            if details_div:
                tat_details = " | ".join(
                    p.get_text(strip=True)
                    for p in details_div.find_all("p")
                    if p.get_text(strip=True)
                )

    # ── Supported modes (from the mode tablist) ───────────────────────────────
    supported_modes = []
    for tablist in soup.find_all(attrs={"role": "tablist"}):
        label = (tablist.get("aria-label") or "").lower()
        if "supported modes" in label:
            for tab in tablist.find_all(attrs={"role": "tab"}):
                text = tab.get_text(strip=True)
                # Tab text is like "B2BBusiness to business" — take first 3 chars
                mode = text[:3] if len(text) >= 3 else text
                supported_modes.append(mode)
    modes_str = ", ".join(supported_modes)

    base = {
        "Country":        country_name,
        "Payment Mode":   method_label,
        "Currency":       currency_str,
        "TAT":            tat,
        "Supported Modes": modes_str,
    }

    # TAT description as a key-value row
    if tat_details:
        rows.append({**base, "Key": "Cutoff & delivery timing", "Value": tat_details})

    # ── Transaction limits table ──────────────────────────────────────────────
    table = soup.find("table", class_=lambda c: c and "atlas-route-limit" in c)
    if table:
        mode_headers = []  # ["B2B", "B2P", "P2P", "P2B"]
        thead = table.find("thead")
        if thead:
            all_col_ths = thead.find_all("th", attrs={"scope": "col"})
            # First th is the table title ("Transaction limit per end-user"); skip it
            mode_headers = [
                th.get_text(separator=" ", strip=True)
                for th in all_col_ths[1:]
            ]
        tbody = table.find("tbody")
        if tbody:
            for tr in tbody.find_all("tr"):
                row_th = tr.find("th", attrs={"scope": "row"})
                if not row_th:
                    continue
                row_label = re.sub(r'^[+\-−]\s*', '', row_th.get_text(strip=True)).strip()
                for ci, td in enumerate(tr.find_all("td")):
                    if ci < len(mode_headers):
                        rows.append({
                            **base,
                            "Key":   f"Transaction limit per end-user - {mode_headers[ci]} - {row_label}",
                            "Value": td.get_text(strip=True),
                        })

    # ── Ledger rows (simple key → paragraph value) ────────────────────────────
    simple_fields = {
        "Beneficiary statement narrative",
        "Network participant",
        "Channels",
        "Routing code",
        "Proof of payment",
        "Notes",
        "Beneficiary account type",
    }

    for ledger_div in soup.find_all("div", class_=lambda c: c and "figma-ledger-row" in c):
        heading_tag = ledger_div.find(["h2", "h3"])
        if not heading_tag:
            continue
        field_name = heading_tag.get_text(strip=True)
        copy_div = ledger_div.find("div", class_=lambda c: c and "atlas-ledger-copy" in c)
        if not copy_div:
            continue

        if field_name in simple_fields:
            value = " ".join(
                p.get_text(strip=True)
                for p in copy_div.find_all("p")
                if p.get_text(strip=True)
            )
            if value:
                rows.append({**base, "Key": field_name, "Value": value})

        elif field_name == "Mandatory data requirements":
            groups = []
            for group_div in copy_div.find_all("div", class_=lambda c: c and "atlas-requirement-group" in c):
                group_h = group_div.find(["h3", "h4"])
                group_name = group_h.get_text(strip=True) if group_h else ""
                items = [li.get_text(strip=True) for li in group_div.find_all("li")]
                if items:
                    groups.append(f"{group_name}: {', '.join(items)}" if group_name else ", ".join(items))
            if groups:
                rows.append({**base, "Key": "Mandatory data requirements", "Value": "; ".join(groups)})

        elif field_name == "Supporting documents":
            items = [li.get_text(strip=True) for li in copy_div.find_all("li")]
            if items:
                rows.append({**base, "Key": "Supporting documents", "Value": ", ".join(items)})

    return rows


def scrape_dataset(dataset_type, progress_bar, status_text):
    """
    Scrape atlas.nium.com for all countries.
    dataset_type: 'FI' or 'Non-FI'
    Returns: (raw_rows: list[dict], failed: list[str])
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from webdriver_manager.chrome import ChromeDriverManager
    from bs4 import BeautifulSoup

    fi_param = "&audience=fi" if dataset_type == "FI" else ""

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--log-level=3")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    raw_rows = []
    total    = len(COUNTRIES)
    failed   = []

    for idx, country_name in enumerate(COUNTRIES):
        pct = (idx + 1) / total
        progress_bar.progress(pct, text=f"Scraping {dataset_type}: {country_name} ({idx+1}/{total})")
        status_text.caption(f"⏳ {country_name}...")

        slug        = _format_country_url(country_name)
        country_url = f"https://atlas.nium.com/country/{slug}"
        if dataset_type == "FI":
            country_url += "?audience=fi"

        # ── Step 1: load country page to discover available method IDs ────────
        page_loaded = False
        for attempt in range(SCRAPE_MAX_RETRIES):
            try:
                driver.get(country_url)
                WebDriverWait(driver, SCRAPE_PAGE_TIMEOUT_SECONDS).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select"))
                )
                page_loaded = True
                break
            except Exception:
                if attempt < SCRAPE_MAX_RETRIES - 1:
                    time.sleep(SCRAPE_RETRY_DELAY_SECONDS)

        if not page_loaded:
            failed.append(country_name)
            continue

        soup = BeautifulSoup(driver.page_source, "html.parser")
        select_el = soup.find("select")
        if not select_el:
            failed.append(country_name)
            continue

        method_options = [
            (opt.get("value", "").strip(), opt.get_text(strip=True).split(" · ")[0].strip())
            for opt in select_el.find_all("option")
            if opt.get("value", "").strip()
        ]
        if not method_options:
            failed.append(country_name)
            continue

        # ── Step 2: scrape each method by navigating to its direct URL ────────
        for method_id, method_label in method_options:
            method_url = f"https://atlas.nium.com/country/{slug}?method={method_id}{fi_param}"

            page_loaded = False
            for attempt in range(SCRAPE_MAX_RETRIES):
                try:
                    driver.get(method_url)
                    WebDriverWait(driver, SCRAPE_PAGE_TIMEOUT_SECONDS).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, ".figma-route-summary__row")
                        )
                    )
                    page_loaded = True
                    break
                except Exception:
                    if attempt < SCRAPE_MAX_RETRIES - 1:
                        time.sleep(SCRAPE_RETRY_DELAY_SECONDS)

            if not page_loaded:
                continue

            try:
                method_soup = BeautifulSoup(driver.page_source, "html.parser")
                method_rows = _parse_method_page(method_soup, country_name, method_label)
                raw_rows.extend(method_rows)
            except Exception:
                continue

    driver.quit()
    return raw_rows, failed


def transform_raw_to_wide(raw_rows):
    """Convert long-format key-value rows into wide-format DataFrame."""
    if not raw_rows:
        return pd.DataFrame()

    df_long = pd.DataFrame(raw_rows)
    id_cols = ["Country", "Payment Mode", "Currency", "TAT", "Supported Modes"]

    # Rename scraped Keys that collide with id_cols — pivot + reset_index
    # raises ValueError: "cannot insert X, already exists" otherwise
    conflicts = set(id_cols) & set(df_long["Key"].unique())
    if conflicts:
        df_long["Key"] = df_long["Key"].replace({k: f"{k} (Detail)" for k in conflicts})

    df_wide = df_long.pivot_table(
        index=id_cols, columns="Key", values="Value", aggfunc="first"
    ).reset_index()
    df_wide.columns.name = None

    other_cols = [c for c in df_wide.columns if c not in id_cols]
    preferred_order = [
        "Network participant", "Channels", "Routing code",
        "Cutoff & delivery timing", "Mandatory data requirements",
        "Supporting documents", "Beneficiary statement narrative",
        "Proof of payment", "Beneficiary account type", "Notes",
    ]
    ordered = [c for c in preferred_order if c in other_cols]
    remaining = sorted(c for c in other_cols if c not in ordered)
    return df_wide[id_cols + ordered + remaining]


def save_scraped_data(df, dataset_type):
    """Save to data/ (live) and scraped_data/ (date-stamped archive)."""
    today    = datetime.now().strftime("%Y-%m-%d")
    app_path = FI_PATH if dataset_type == "FI" else NON_FI_PATH
    df.to_excel(str(app_path), index=False)
    archive_path = ARCHIVE_DIR / f"{dataset_type}_{today}.xlsx"
    df.to_excel(str(archive_path), index=False)
    return str(app_path), str(archive_path)


def log_scrape_failures(failures_by_dataset: dict) -> str | None:
    """Append failed-country details to logs/scrape_failures.log.

    failures_by_dataset: {"FI": [...], "Non-FI": [...]}
    Returns the log file path, or None if there were no failures.
    """
    all_failed = {ds: countries for ds, countries in failures_by_dataset.items() if countries}
    if not all_failed:
        return None
    log_dir  = BASE_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / "scrape_failures.log"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"\n{'='*60}", f"Scrape run: {timestamp}"]
    for ds, countries in all_failed.items():
        lines.append(f"\n[{ds}] {len(countries)} countries skipped:")
        for c in sorted(countries):
            lines.append(f"  - {c}")
    lines.append("")
    with open(log_path, "a", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    return str(log_path)


def create_formatted_excel(data, selected_cols, dataset_type):
    """Build a styled Excel workbook from a filtered DataFrame and return a BytesIO buffer."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Nium Capabilities'

    blue_header   = '4472C4'
    thin_border   = Border(
        left=Side(style='thin', color='B4B4B4'), right=Side(style='thin', color='B4B4B4'),
        top=Side(style='thin', color='B4B4B4'),  bottom=Side(style='thin', color='B4B4B4'),
    )
    header_border = Border(
        left=Side(style='thin', color='FFFFFF'), right=Side(style='thin', color='FFFFFF'),
        top=Side(style='thin', color='FFFFFF'),  bottom=Side(style='medium', color='2F5496'),
    )

    ws.row_dimensions[1].height = 8
    ws.merge_cells('A2:E2')
    title_cell           = ws['A2']
    title_cell.value     = "Nium Payout Capability Matrix"
    title_cell.font      = Font(name='Segoe UI Semibold', size=14, bold=True, color='1A1A1A')
    title_cell.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 6

    headers = ['#'] + selected_cols
    for col_idx, header in enumerate(headers, 1):
        cell           = ws.cell(row=4, column=col_idx)
        cell.value     = header
        cell.font      = Font(name='Segoe UI Semibold', size=11, bold=True, color='FFFFFF')
        cell.fill      = PatternFill(start_color=blue_header, end_color=blue_header, fill_type='solid')
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        cell.border    = header_border
    ws.row_dimensions[4].height = 24
    ws.auto_filter.ref = f'A4:{get_column_letter(len(headers))}4'

    export_df = data[selected_cols].reset_index(drop=True)
    for row_idx, (_, row) in enumerate(export_df.iterrows()):
        excel_row         = row_idx + 5
        sn_cell           = ws.cell(row=excel_row, column=1)
        sn_cell.value     = row_idx + 1
        sn_cell.font      = Font(name='Segoe UI Semilight', size=11, color='333333')
        sn_cell.alignment = Alignment(horizontal='right', vertical='top')
        sn_cell.border    = thin_border
        for col_idx, col_name in enumerate(selected_cols):
            cell           = ws.cell(row=excel_row, column=col_idx + 2)
            val            = row[col_name]
            cell.value     = None if (pd.isna(val) or str(val) in ('nan', 'None', '')) else str(val)
            cell.font      = Font(name='Segoe UI Semilight', size=11, color='333333')
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            cell.border    = thin_border

    ws.column_dimensions['A'].width = 5
    for col_idx, col_name in enumerate(selected_cols):
        col_letter = get_column_letter(col_idx + 2)
        max_len    = len(col_name)
        for ri in range(5, ws.max_row + 1):
            cv = ws.cell(row=ri, column=col_idx + 2).value
            if cv:
                max_len = max(max_len, min(len(str(cv)), 45))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

    ws.freeze_panes = 'B5'
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def get_last_updated():
    """Return human-readable timestamp of the FI data file, or None."""
    if FI_PATH.exists():
        ts = os.path.getmtime(str(FI_PATH))
        return datetime.fromtimestamp(ts).strftime("%d %b %Y, %I:%M %p")
    return None
