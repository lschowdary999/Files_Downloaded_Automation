"""
download_renewals.py
====================
Downloads SDS PDF files directly from an Excel sheet:
1. Validates PDF integrity (skips non-PDF / HTML error pages and logs them).
2. Extracts Revision Date strictly:
     - Multi-language Revision Date detection (English, German, French, Spanish, Italian, Portuguese, Dutch)
     - Validates and standardizes date format to MM-DD-YYYY
     - Files WITH Revision Date    -> saved to main folder as "<DocumentID>_<MM-DD-YYYY>.pdf"
     - Files WITHOUT Revision Date -> saved to no_revisions folder as "<DocumentID>.pdf"
3. Generates Excel reports:
     - download_report.xlsx (full details)
     - no_revisions_report.xlsx (files saved without revision date in no_revisions)
     - skipped_downloads_report.xlsx (skipped non-PDFs / failed downloads for re-downloading)

Requirements (auto-installed on first run):
  pip install openpyxl pdfplumber requests
"""

import os
import re
import sys
import shutil
import tempfile
import argparse
import time
from datetime import datetime

# ---------------------------------------------------------------------------
# Auto-install missing dependencies
# ---------------------------------------------------------------------------
def ensure_packages(*packages):
    import importlib, subprocess
    for pkg, import_name in packages:
        try:
            importlib.import_module(import_name)
        except ImportError:
            print(f"  Installing {pkg}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])

print("\n[0/4] Checking dependencies...")
ensure_packages(
    ("openpyxl",   "openpyxl"),
    ("pdfplumber", "pdfplumber"),
    ("requests",   "requests"),
    ("selenium",   "selenium"),
)

import requests
import logging
import warnings
import urllib3
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Suppress noisy PDF parser warnings & insecure request warnings
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)
warnings.filterwarnings("ignore", category=UserWarning)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# CONFIGURATION & PATHS
# ---------------------------------------------------------------------------
BASE_DIR        = r"c:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF"
EXCEL_FILE      = r"C:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF\SHEET_P.xlsx"
OUTPUT_DIR      = r"C:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF\Sheet_P"
NO_REV_DIR      = r"C:\Users\I01884\OneDrive - 3E Company Environmental\Desktop\AP1_DRF\Sheet_P\no_revisions"
TEMP_DIR        = os.path.join(tempfile.gettempdir(), "sds_temp_downloads")
REPORT_PATH     = os.path.join(BASE_DIR, "download_report.xlsx")
NO_REV_REPORT   = os.path.join(NO_REV_DIR, "no_revisions_report.xlsx")
SKIPPED_REPORT  = os.path.join(BASE_DIR, "skipped_downloads_report.xlsx")

# Excel Column Mappings (Used as fallback if auto-header detection fails)
COL_DOC_ID      = "A"
COL_PRODUCT     = "B"
COL_URL_1       = "L"
COL_URL_2       = "M"
COL_URL_3       = "W"

for d in [OUTPUT_DIR, NO_REV_DIR, TEMP_DIR]:
    os.makedirs(d, exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,application/pdf,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# ---------------------------------------------------------------------------
# MULTI-LANGUAGE MONTH LOOKUP
# Covers: English, French, Spanish, German, Italian, Portuguese, Dutch
# ---------------------------------------------------------------------------
def get_month_number(month_str: str) -> str | None:
    """Return 2-digit month number string from a month name in any supported language."""
    s = re.sub(r"[^a-zA-Z]", "", month_str).lower()

    # January
    if re.match(r"^jan", s): return "01"
    if re.match(r"^ene", s): return "01"
    # February
    if re.match(r"^feb", s): return "02"
    if re.match(r"^fev", s): return "02"
    # March
    if re.match(r"^mar", s): return "03"
    # April
    if re.match(r"^apr", s): return "04"
    if re.match(r"^avr", s): return "04"
    if re.match(r"^abr", s): return "04"
    # May
    if s in ("may", "mai", "mayo", "mei"): return "05"
    if re.match(r"^mag", s): return "05"
    # June
    if re.match(r"^jun", s): return "06"
    if re.match(r"^giu", s): return "06"
    # July
    if re.match(r"^jul", s): return "07"
    if re.match(r"^lug", s): return "07"
    # August
    if re.match(r"^aug", s): return "08"
    if re.match(r"^ago", s): return "08"
    if re.match(r"^aou", s) or s == "aot": return "08"
    # September
    if re.match(r"^sep", s): return "09"
    if re.match(r"^set", s): return "09"
    # October
    if re.match(r"^oct", s): return "10"
    if re.match(r"^okt", s): return "10"
    if re.match(r"^out", s): return "10"
    if re.match(r"^ott", s): return "10"
    # November
    if re.match(r"^nov", s): return "11"
    # December
    if re.match(r"^dec", s): return "12"
    if re.match(r"^dez", s): return "12"
    if re.match(r"^dic", s): return "12"
    if re.match(r"^dce", s): return "12"

    return None

# ---------------------------------------------------------------------------
# DATE EXTRACTION KEYWORDS & PATTERNS (STRICT REVISION DATE ONLY)
# ---------------------------------------------------------------------------
REVISION_KEYWORDS = [
    # English (and combined English terms)
    r"Date\s*of\s*issue\s*[\/\\]\s*Date\s*of\s*revision",
    r"Issue\s*Date\s*[\/\\]\s*Revision\s*Date",
    r"Date\s*of\s*(?:Last\s*)?Revision",
    r"Revision\s*Date",
    r"Revised\s*Date",
    r"Date\s*Revised",
    r"Revised\s*on",
    r"Revised",
    r"Last\s*Revision(?:\s*Date)?",
    r"Last\s*Revised",
    r"Date\s*of\s*(?:Last\s*)?Update",
    r"Update\s*Date",
    r"Updated\s*on",
    r"Updated",
    r"Last\s*Updated",
    r"Rev\.\s*Date",
    r"Rev\s*Date",
    r"Rev\s*[:\.\-]",
    r"Rev\.\s*[:\.\-]",
    r"Revision\s*[:\.\-]",
    r"Revisiondate",
    r"Revision",

    # German
    r"Ausgabedatum\s*[\/\\]\s*[ÜüUeue]+berarbeitungsdatum",
    r"Datum\s*der\s*(?:letzten\s*)?[ÜüUeue]+berarbeitung",
    r"[ÜüUeue]+berarbeitungsdatum",
    r"[ÜüUeue]+berarbeitet\s*am",
    r"[ÜüUeue]+berarbeitet",
    r"Revisionsdatum",
    r"Datum\s*der\s*Revision",
    r"Revision\s*vom",
    r"Revidiert\s*am",
    r"Stand\s*der\s*Information",
    r"Stand\s*vom",

    # French
    r"Date\s*d['’]?[eE]mission\s*[\/\\]\s*Date\s*de\s*r[eéèEÉÈ]vision",
    r"Date\s*de\s*(?:la\s*)?(?:derni[eèéEÈÉ]re\s*)?r[eéèEÉÈ]vision",
    r"Date\s*de\s*(?:derni[eèéEÈÉ]re\s*)?mise\s*(?:[aAàÀ]\s*)?jour",
    r"Mise\s*(?:[aAàÀ]\s*)?jour\s*(?:le)?",
    r"R[eéèEÉÈ]vis[eéèEÉÈ]\s*le",
    r"Derni[eèéEÈÉ]re\s*r[eéèEÉÈ]vision",

    # Spanish
    r"Fecha\s*de\s*emisi[oóOÓ]n\s*[\/\\]\s*Fecha\s*de\s*revisi[oóOÓ]n",
    r"Fecha\s*de\s*(?:la\s*)?(?:[uúUÚ]ltima\s*)?revisi[oóOÓ]n",
    r"Fecha\s*de\s*(?:[uúUÚ]ltima\s*)?actualizaci[oóOÓ]n",
    r"Actualizado\s*(?:el)?",
    r"Revisado\s*(?:el)?",
    r"[UúUÚ]ltima\s*revisi[oóOÓ]n",

    # Italian
    r"Data\s*di\s*emissione\s*[\/\\]\s*Data\s*di\s*revisione",
    r"Data\s*di\s*(?:dell['’]ultima\s*)?revisione",
    r"Data\s*di\s*(?:dell['’]ultimo\s*)?aggiornamento",
    r"Aggiornato\s*il",
    r"Revisionato\s*il",

    # Portuguese
    r"Data\s*de\s*emiss[aãAÃ]o\s*[\/\\]\s*Data\s*de\s*revis[aãAÃ]o",
    r"Data\s*de\s*(?:da\s*)?(?:[uúUÚ]ltima\s*)?revis[aãAÃ]o",
    r"Data\s*de\s*(?:da\s*)?(?:[uúUÚ]ltima\s*)?atualiza[cçCÇ][aãAÃ]o",
    r"Revisado\s*em",
    r"Atualizado\s*em",

    # Dutch
    r"Herzieningsdatum",
    r"Datum\s*van\s*herziening",
    r"Revisiedatum",
    r"Datum\s*van\s*bijwerking",
    r"Bijgewerkt\s*op",
    r"Herzien\s*op",
]

def build_keyword_patterns(keyword_list: list[str]) -> list[re.Pattern]:
    combined_kw = "(?:" + "|".join(keyword_list) + ")"
    return [
        re.compile(rf"{combined_kw}\s*[:\-\/]?\s*([A-Za-z0-9\s,\.\/\-]+)", re.IGNORECASE),
        re.compile(rf"{combined_kw}[\s\r\n]+([A-Za-z0-9\s,\.\/\-]+)", re.IGNORECASE),
    ]

REVISION_PATTERNS = build_keyword_patterns(REVISION_KEYWORDS)

# ---------------------------------------------------------------------------
# DATE PARSING HELPER
# ---------------------------------------------------------------------------
def parse_date_string(snippet: str) -> str | None:
    """Parse a valid MM-DD-YYYY date from a text snippet."""
    if not snippet:
        return None

    # 1. ISO YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
    iso = re.search(r"\b(19\d{2}|20\d{2})[\/\-\.](0?[1-9]|1[0-2])[\/\-\.](0?[1-9]|[12]\d|3[01])\b", snippet)
    if iso:
        yyyy = iso.group(1)
        mm   = iso.group(2).zfill(2)
        dd   = iso.group(3).zfill(2)
        return f"{mm}-{dd}-{yyyy}"

    # 2. Numeric DD.MM.YYYY (European standard format e.g. 26.05.2025 / 10.10.2024)
    n_eu = re.search(r"\b(0?[1-9]|[12]\d|3[01])\.(0?[1-9]|1[0-2])\.(19\d{2}|20\d{2})\b", snippet)
    if n_eu:
        dd   = n_eu.group(1).zfill(2)
        mm   = n_eu.group(2).zfill(2)
        yyyy = n_eu.group(3)
        return f"{mm}-{dd}-{yyyy}"

    # 3. DD MonthName YYYY (e.g. 29 August 2022 / 21-Dec-2025 / 15. Januar 2024)
    t1 = re.search(r"\b(0?[1-9]|[12]\d|3[01])[\s\-\.\/]+([A-Za-z]{3,12})\.?[\s\-\.\/,]*(19\d{2}|20\d{2})\b", snippet)
    if t1:
        mn = get_month_number(t1.group(2))
        if mn:
            return f"{mn}-{t1.group(1).zfill(2)}-{t1.group(3)}"

    # 4. MonthName DD, YYYY (e.g. August 29, 2022 / Aug-29-2022)
    t2 = re.search(r"\b([A-Za-z]{3,12})\.?[\s\-\.\/]+(0?[1-9]|[12]\d|3[01])[\s\-\.\/,]*(19\d{2}|20\d{2})\b", snippet)
    if t2:
        mn = get_month_number(t2.group(1))
        if mn:
            return f"{mn}-{t2.group(2).zfill(2)}-{t2.group(3)}"

    # 5. Numeric DD/MM/YYYY or DD-MM-YYYY (where day is clearly > 12)
    d_first = re.search(r"\b(1[3-9]|[23]\d)[\/\-](0?[1-9]|1[0-2])[\/\-](19\d{2}|20\d{2}|\d{2})\b", snippet)
    if d_first:
        dd   = d_first.group(1).zfill(2)
        mm   = d_first.group(2).zfill(2)
        yyyy = d_first.group(3)
        if len(yyyy) == 2:
            yyyy = "20" + yyyy
        return f"{mm}-{dd}-{yyyy}"

    # 6. Numeric MM/DD/YYYY or MM-DD-YYYY
    n = re.search(r"\b(0?[1-9]|1[0-2])[\/\-](0?[1-9]|[12]\d|3[01])[\/\-](19\d{2}|20\d{2}|\d{2})\b", snippet)
    if n:
        mm   = n.group(1).zfill(2)
        dd   = n.group(2).zfill(2)
        yyyy = n.group(3)
        if len(yyyy) == 2:
            yyyy = "20" + yyyy
        return f"{mm}-{dd}-{yyyy}"

    # 7. MonthName YYYY (e.g. August 2022 / May 2015 -> 08-01-2022)
    t3 = re.search(r"\b([A-Za-z]{3,12})\.?[\s\-\.\/,]+(19\d{2}|20\d{2})\b", snippet)
    if t3:
        mn = get_month_number(t3.group(1))
        if mn:
            return f"{mn}-01-{t3.group(2)}"

    return None

# ---------------------------------------------------------------------------
# PDF VALIDATION & TEXT EXTRACTION
# ---------------------------------------------------------------------------
def is_valid_pdf(file_path: str) -> bool:
    """Validate that the file exists, is non-empty, contains PDF magic bytes, and can be read."""
    if not os.path.exists(file_path):
        return False
    if os.path.getsize(file_path) < 300:
        return False
    try:
        with open(file_path, "rb") as f:
            header = f.read(1024)
            if b"%PDF-" not in header:
                return False
        with pdfplumber.open(file_path) as pdf:
            if len(pdf.pages) == 0:
                return False
        return True
    except Exception:
        return False

def extract_pdf_text(pdf_path: str) -> str:
    """Extract full text from a PDF file."""
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text
    except Exception:
        return ""

# ---------------------------------------------------------------------------
# STRICT REVISION DATE EXTRACTION ONLY
# - ONLY extracts genuine Revision Date
# - Skips all Issue / Print / Creation / Effective Dates
# - If no revision date is found -> Returns (None, "Not Found", False)
# ---------------------------------------------------------------------------
def extract_sds_date_info(pdf_text: str) -> tuple[str | None, str, bool]:
    """
    STRICT REVISION DATE EXTRACTION ONLY.
    Returns: (date_string, date_type_label, is_found_in_pdf)
    """
    if not pdf_text:
        return (None, "Not Found", False)

    for pat in REVISION_PATTERNS:
        for m in pat.finditer(pdf_text):
            snippet = m.group(1).strip()
            # Isolate revision date portion before any subsequent issue/previous date labels
            snippet = re.split(
                r"(?:Date\s*of\s*previous|Datum\s*der\s*letzten|Previous\s*issue|Ausgabe|Druckdatum|Version\s*:)",
                snippet,
                flags=re.IGNORECASE
            )[0]
            snippet = re.sub(r"^[\s:\-\/\.]+", "", snippet).strip()[:60]
            d = parse_date_string(snippet)
            if d:
                return (d, "Revision Date", True)

    # If NO revision date was found:
    return (None, "Not Found", False)

# ---------------------------------------------------------------------------
# EXCEL READ / WRITE UTILITIES
# ---------------------------------------------------------------------------
def read_excel_items(excel_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    ws = wb.active

    # Detect header columns dynamically
    doc_id_col = None
    product_col = None
    url_cols = []

    for col_idx, cell in enumerate(ws[1], start=1):
        val = str(cell.value or "").strip().lower()
        if not doc_id_col and re.search(r"document\s*id|doc\s*id|docid", val):
            doc_id_col = col_idx
        elif not product_col and re.search(r"product\s*name|productname|product", val):
            product_col = col_idx
        elif re.search(r"link\s*to\s*sds|link|url|sds\s*link|sds\s*url|web\s*link", val):
            url_cols.append(col_idx)

    # Fallbacks if headers not detected
    if not doc_id_col:
        doc_id_col = 1
    if not product_col:
        product_col = 2
    if not url_cols:
        url_cols = [12, 13, 23]  # Columns L, M, W

    items = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        doc_cell = row[doc_id_col - 1] if len(row) >= doc_id_col else None
        prod_cell = row[product_col - 1] if len(row) >= product_col else None

        doc_id = str(doc_cell.value).strip() if doc_cell and doc_cell.value is not None else None
        if not doc_id or doc_id.lower() in ("none", "nan", ""):
            continue

        product = str(prod_cell.value).strip() if prod_cell and prod_cell.value is not None else ""

        # Extract URL or local file path from priority URL columns
        url = None
        for col_i in url_cols:
            if len(row) >= col_i:
                c = row[col_i - 1]
                if c.hyperlink and c.hyperlink.target:
                    url = str(c.hyperlink.target).strip()
                    break
                elif c.value and str(c.value).strip():
                    val = str(c.value).strip()
                    if val.startswith(("http://", "https://", "file:///")) or val.lower().endswith((".pdf", ".html")):
                        url = val
                        break

        # Fallback: scan all cells in this row for any hyperlink or URL / file path
        if not url:
            for c in row:
                if c.hyperlink and c.hyperlink.target:
                    url = str(c.hyperlink.target).strip()
                    break
                elif c.value and str(c.value).strip().startswith(("http://", "https://", "file:///")):
                    url = str(c.value).strip()
                    break

        if doc_id and url:
            items.append({
                "row": row_idx,
                "doc_id": doc_id,
                "product": product,
                "url": url,
            })

    return items

HEADER_FILL = PatternFill("solid", fgColor="00467F")
HEADER_FONT = Font(bold=True, color="FFFFFF")

class LiveExcelReport:
    """Manages an Excel report that saves incrementally on every row addition."""
    def __init__(self, file_path: str, headers: list[str], sheet_name: str):
        self.file_path = file_path
        self.headers = headers
        self.sheet_name = sheet_name
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = sheet_name

        # Header row
        for col_idx, h in enumerate(headers, start=1):
            cell = self.ws.cell(row=1, column=col_idx, value=h)
            cell.font      = HEADER_FONT
            cell.fill      = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        self.save()

    def append_row(self, record: dict):
        row_values = [str(record.get(h, "")) for h in self.headers]
        self.ws.append(row_values)
        self.save()

    def save(self):
        try:
            self.wb.save(self.file_path)
        except PermissionError:
            print(f"   [WARNING] Could not save {os.path.basename(self.file_path)} (file is currently open in Excel).", flush=True)
        except Exception as e:
            print(f"   [WARNING] Failed to save {os.path.basename(self.file_path)}: {e}", flush=True)

    def finalize(self):
        try:
            for col in self.ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=10)
                self.ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
            self.wb.save(self.file_path)
        except Exception:
            pass

# ---------------------------------------------------------------------------
# SILENT HEADLESS SESSION RESOLVER (ZERO GUI / POPUPS)
# Resolves session cookies for protected portals (e.g. thewercs.com)
# ---------------------------------------------------------------------------
_HEADLESS_DRIVER = None

def get_headless_driver():
    global _HEADLESS_DRIVER
    if _HEADLESS_DRIVER is None:
        try:
            from selenium import webdriver
            from selenium.webdriver.edge.options import Options as EdgeOptions
            options = EdgeOptions()
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--log-level=3")
            _HEADLESS_DRIVER = webdriver.Edge(options=options)
            _HEADLESS_DRIVER.set_page_load_timeout(20)
            _HEADLESS_DRIVER.set_script_timeout(20)
        except Exception:
            try:
                from selenium import webdriver
                from selenium.webdriver.chrome.options import Options as ChromeOptions
                options = ChromeOptions()
                options.add_argument("--headless=new")
                options.add_argument("--disable-gpu")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--log-level=3")
                _HEADLESS_DRIVER = webdriver.Chrome(options=options)
                _HEADLESS_DRIVER.set_page_load_timeout(20)
                _HEADLESS_DRIVER.set_script_timeout(20)
            except Exception:
                _HEADLESS_DRIVER = None
    return _HEADLESS_DRIVER

def close_headless_driver():
    global _HEADLESS_DRIVER
    if _HEADLESS_DRIVER is not None:
        try:
            _HEADLESS_DRIVER.quit()
        except Exception:
            pass
        _HEADLESS_DRIVER = None

def download_with_session_cookies(url: str, dest_temp_file: str) -> None:
    """Uses a background headless browser to resolve portal session cookies, then downloads the PDF."""
    if os.path.exists(dest_temp_file):
        os.remove(dest_temp_file)

    driver = get_headless_driver()
    if driver is None:
        raise ValueError("Headless browser driver could not be initialized.")

    try:
        driver.set_page_load_timeout(20)
        driver.get(url)
        time.sleep(2.0)
    except Exception as e:
        close_headless_driver()
        raise ValueError(f"Browser navigation timed out or failed: {e}")

    cookies = driver.get_cookies()
    if not cookies:
        raise ValueError("No session cookies obtained from portal.")

    session = requests.Session()
    session.headers.update(HEADERS)
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer": driver.current_url
    })
    for c in cookies:
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))

    resp = session.get(url, timeout=20, stream=True)
    resp.raise_for_status()

    ctype = resp.headers.get("Content-Type", "").lower()
    if "text/html" in ctype or "text/plain" in ctype:
        raise ValueError(f"Server returned HTML/text even with session cookies (Content-Type: {ctype})")

    with open(dest_temp_file, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

    if not is_valid_pdf(dest_temp_file):
        if os.path.exists(dest_temp_file):
            os.remove(dest_temp_file)
        raise ValueError("Downloaded file is not a valid PDF.")

# ---------------------------------------------------------------------------
# DOWNLOAD HANDLER
# ---------------------------------------------------------------------------
def download_regular(url: str, dest_temp_file: str) -> None:
    """Attempts direct HTTP download or local file copy with SSL bypass and headless fallback."""
    import urllib.parse
    if os.path.exists(dest_temp_file):
        os.remove(dest_temp_file)

    # Handle local file:/// links or local file paths
    if url.startswith("file://") or (os.path.isabs(url) and not url.startswith("http")):
        clean_path = url.replace("file:///", "").replace("file://", "").replace("/", "\\")
        clean_path = urllib.parse.unquote(clean_path)
        if os.path.exists(clean_path):
            shutil.copyfile(clean_path, dest_temp_file)
            return
        else:
            raise FileNotFoundError(f"Local file link does not exist on disk: {clean_path}")

    try:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=20, stream=True)
        except requests.exceptions.SSLError:
            # Fallback: ignore SSL certificate verification for expired/misconfigured servers
            resp = requests.get(url, headers=HEADERS, timeout=20, stream=True, verify=False)

        resp.raise_for_status()

        ctype = resp.headers.get("Content-Type", "").lower()
        if "text/html" in ctype or "text/plain" in ctype:
            download_with_session_cookies(url, dest_temp_file)
            return

        with open(dest_temp_file, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)

        if not is_valid_pdf(dest_temp_file):
            download_with_session_cookies(url, dest_temp_file)

    except Exception as e:
        err_str = str(e).lower()
        if "thewercs.com" in url or "html" in err_str or "403" in err_str or "timeout" in err_str or "forbidden" in err_str:
            download_with_session_cookies(url, dest_temp_file)
        else:
            raise e

# ---------------------------------------------------------------------------
# MAIN WORKFLOW
# ---------------------------------------------------------------------------
def main(limit: int = 0):
    print("\n[1/4] Reading Excel file...")
    items = read_excel_items(EXCEL_FILE)
    print(f"   Found {len(items)} records in Excel.")

    if limit > 0:
        items = items[:limit]

    print(f"\n[2/4] Initializing Live Excel Reports...")
    report_headers = ["Document ID", "Product Name", "URL", "Status", "Date Type",
                      "Revision Date", "Saved As", "Folder", "Timestamp", "Error"]
    no_rev_headers = ["Document ID", "Product Name", "URL", "Saved As", "Folder", "Timestamp"]
    skipped_headers = ["Document ID", "Product Name", "URL", "Status", "Reason", "Timestamp"]

    main_report_writer    = LiveExcelReport(REPORT_PATH, report_headers, "Download Report")
    no_rev_report_writer  = LiveExcelReport(NO_REV_REPORT, no_rev_headers, "No Revision Date")
    skipped_report_writer = LiveExcelReport(SKIPPED_REPORT, skipped_headers, "Skipped & Failed")

    print(f"\n[3/4] Processing {len(items)} files (saving live to Excel)...")

    success_count = fail_count = no_rev_count = skipped_count = 0

    def process_downloaded_file(item, temp_file):
        nonlocal success_count, no_rev_count
        doc_id    = item["doc_id"]
        product   = item["product"]
        url       = item["url"]
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        pdf_text = extract_pdf_text(temp_file)
        sds_date, date_type, is_found = extract_sds_date_info(pdf_text)

        if is_found and sds_date:
            final_file = f"{doc_id}_{sds_date}.pdf"
            dest       = os.path.join(OUTPUT_DIR, final_file)
            shutil.move(temp_file, dest)
            print(f"[SUCCESS] {date_type}: {sds_date} -> {final_file}")

            main_report_writer.append_row({
                "Document ID":   doc_id,
                "Product Name":  product,
                "URL":           url,
                "Status":        "SUCCESS",
                "Date Type":     date_type,
                "Revision Date": sds_date,
                "Saved As":      final_file,
                "Folder":        os.path.basename(OUTPUT_DIR),
                "Timestamp":     timestamp,
                "Error":         "",
            })
            success_count += 1
        else:
            # Saved with Document ID only into no_revisions folder
            final_file = f"{doc_id}.pdf"
            dest       = os.path.join(NO_REV_DIR, final_file)
            shutil.move(temp_file, dest)
            print(f"[NO REV DATE] Saved with Document ID -> {final_file}")

            main_report_writer.append_row({
                "Document ID":   doc_id,
                "Product Name":  product,
                "URL":           url,
                "Status":        "NO REVISION DATE",
                "Date Type":     "Not Found",
                "Revision Date": "Not Found",
                "Saved As":      final_file,
                "Folder":        os.path.join(os.path.basename(OUTPUT_DIR), "no_revisions"),
                "Timestamp":     timestamp,
                "Error":         "",
            })
            no_rev_report_writer.append_row({
                "Document ID":  doc_id,
                "Product Name": product,
                "URL":          url,
                "Saved As":     final_file,
                "Folder":       os.path.join(os.path.basename(OUTPUT_DIR), "no_revisions"),
                "Timestamp":    timestamp,
            })
            no_rev_count += 1

    def record_skipped_or_failed(item, err_msg):
        nonlocal fail_count, skipped_count
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        is_non_pdf = "not a valid pdf" in err_msg.lower() or "html" in err_msg.lower()

        status = "SKIPPED (NON-PDF)" if is_non_pdf else "FAILED"
        print(f"[{status}] {err_msg}")

        main_report_writer.append_row({
            "Document ID":   item["doc_id"],
            "Product Name":  item["product"],
            "URL":           item["url"],
            "Status":        status,
            "Date Type":     "",
            "Revision Date": "",
            "Saved As":      "",
            "Folder":        "",
            "Timestamp":     timestamp,
            "Error":         err_msg,
        })

        if is_non_pdf:
            skipped_count += 1
            skipped_report_writer.append_row({
                "Document ID":   item["doc_id"],
                "Product Name":  item["product"],
                "URL":           item["url"],
                "Status":        "SKIPPED",
                "Reason":        err_msg,
                "Timestamp":     timestamp,
            })
        else:
            fail_count += 1
            skipped_report_writer.append_row({
                "Document ID":   item["doc_id"],
                "Product Name":  item["product"],
                "URL":           item["url"],
                "Status":        "FAILED",
                "Reason":        err_msg,
                "Timestamp":     timestamp,
            })

    try:
        for item in items:
            row_num   = item["row"]
            doc_id    = item["doc_id"]
            url       = item["url"]
            temp_file = os.path.join(TEMP_DIR, f"{doc_id}_temp.pdf")

            print(f"Row {row_num} [DocID: {doc_id}] ", end="", flush=True)

            try:
                download_regular(url, temp_file)
                process_downloaded_file(item, temp_file)
            except Exception as e:
                err_msg = str(e)
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                record_skipped_or_failed(item, err_msg)
    finally:
        close_headless_driver()
        # Finalize and auto-fit column widths
        main_report_writer.finalize()
        no_rev_report_writer.finalize()
        skipped_report_writer.finalize()

    # -------------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------------
    print()
    print("[4/4] === SUMMARY " + "=" * 50)
    print(f"   Total processed         : {len(items)}")
    print(f"   Saved with date         : {success_count}  ->  {OUTPUT_DIR}")
    print(f"   Saved without rev date  : {no_rev_count}  ->  {NO_REV_DIR}")
    print(f"   Skipped (Non-PDF/HTML)  : {skipped_count}")
    print(f"   Failed to download      : {fail_count}")
    print(f"   Main report             : {REPORT_PATH}")
    if no_rev_count > 0:
        print(f"   No-Rev report           : {NO_REV_REPORT}")
    if skipped_count + fail_count > 0:
        print(f"   Skipped / Missed report : {SKIPPED_REPORT}")
    print("=" * 68)
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="SDS Renewal Downloader")
    parser.add_argument("--limit", type=int, default=0,
                        help="Limit number of files to process (0 = all)")
    args = parser.parse_args()
    main(limit=args.limit)
