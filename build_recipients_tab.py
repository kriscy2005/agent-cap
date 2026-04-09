"""
Builds (or refreshes) a 'Recipients' tab in the main sheet.
Uses Selenium (logged-in Edge) to download private sheets as CSV,
then writes the result via the Apps Script endpoint.
"""
import os, csv, io, time, re, json, requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
MAIN_SHEET_ID  = "19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY"
EMP_SHEET_ID   = "1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8"
EMP_GID        = "1165838487"
BRAND_SHEET_ID = "1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho"
BRAND_GID      = "1996161867"
OUTPUT_TAB     = "Recipients"
APPS_SCRIPT_URL = os.getenv("APPS_SCRIPT_URL", "")
EDGE_PROFILE    = r"C:\Users\i_ojasvi.singh\AppData\Local\Microsoft\Edge\User Data"

ALLOWED_DESIGNATIONS = [
    "product manager",
    "associate product manager",
    "director",
    "associate director",
    "platform",
    "product management",
]

# ── Selenium CSV fetch ────────────────────────────────────────────────────────
def fetch_csv_via_selenium(sheet_id, gid, label):
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options

    download_dir = str(Path(__file__).parent / "tmp_downloads")
    Path(download_dir).mkdir(exist_ok=True)

    # Clear old CSVs
    for f in Path(download_dir).glob("*.csv"):
        f.unlink()

    opts = Options()
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    print(f"  Downloading {label}...")

    driver = webdriver.Edge(options=opts)
    try:
        driver.get(url)
        # Wait for download
        deadline = time.time() + 30
        while time.time() < deadline:
            csvs = list(Path(download_dir).glob("*.csv"))
            if csvs:
                time.sleep(1)  # let file finish writing
                content = csvs[0].read_text(encoding="utf-8", errors="replace")
                print(f"  Downloaded: {len(content)} chars, {content.count(chr(10))} lines")
                return list(csv.reader(io.StringIO(content)))
            time.sleep(1)
        raise RuntimeError(f"Download timed out for {label}")
    finally:
        driver.quit()

# ── Column detection ──────────────────────────────────────────────────────────
def col_idx(header, candidates):
    for c in candidates:
        for i, h in enumerate(header):
            if c == h: return i
    for c in candidates:
        for i, h in enumerate(header):
            if c in h or h in c: return i
    return -1

def safe(row, idx):
    return row[idx].strip() if idx >= 0 and idx < len(row) else ""

# ── Parse employees ───────────────────────────────────────────────────────────
def parse_employees(rows):
    header = [h.strip().lower() for h in rows[0]]
    print(f"  Employee headers: {header}")

    i_name   = col_idx(header, ["name","employee name","full name","emp name","member name","employee"])
    i_email  = col_idx(header, ["email","work email","email address","corporate email","innovaccer email","official email"])
    i_desig  = col_idx(header, ["designation","title","job title","role","position","level","job role"])
    i_status = col_idx(header, ["status","active","employment status","emp status","employee status"])

    print(f"  Cols → name:{i_name} email:{i_email} desig:{i_desig} status:{i_status}")

    people = []
    for i, row in enumerate(rows[1:], 2):
        name  = safe(row, i_name)
        email = safe(row, i_email)
        desig = safe(row, i_desig)
        status = safe(row, i_status).lower()

        if not name:
            continue
        if i_status >= 0 and status and "active" not in status and status != "":
            continue
        if not matches_desig(desig):
            continue

        people.append({"name": name, "email": email, "designation": desig})

    return people

def matches_desig(desig):
    d = desig.lower().strip()
    return any(allowed in d for allowed in ALLOWED_DESIGNATIONS)

# ── Parse brand map ───────────────────────────────────────────────────────────
def parse_brand_map(rows):
    header = [h.strip().lower() for h in rows[0]]
    print(f"  Brand headers: {header}")

    i_name   = col_idx(header, ["name","employee name","full name","pm name","member name","person"])
    i_brand1 = col_idx(header, ["brand","solution","brand 1","solution 1","brand name","account","product"])
    i_brand2 = col_idx(header, ["brand 2","solution 2","brand2","secondary brand","solution2"])
    i_brand3 = col_idx(header, ["brand 3","solution 3","brand3","solution3"])

    print(f"  Brand cols → name:{i_name} b1:{i_brand1} b2:{i_brand2} b3:{i_brand3}")

    brand_map = {}
    for row in rows[1:]:
        name = safe(row, i_name).lower().strip()
        if not name:
            continue
        brands = [b for b in [safe(row, i_brand1), safe(row, i_brand2), safe(row, i_brand3)] if b]
        brand_map[name] = brands

    return brand_map

def fuzzy_lookup(name, brand_map):
    parts = name.lower().split()
    if len(parts) < 2:
        return []
    first, last = parts[0], parts[-1]
    for k, v in brand_map.items():
        if first in k and last in k:
            return v
    return []

# ── Write to main sheet via Apps Script ───────────────────────────────────────
def write_recipients_tab(people):
    if not APPS_SCRIPT_URL:
        raise RuntimeError("APPS_SCRIPT_URL not set in .env")

    # Build payload: use special action to create/replace a full tab
    rows = []
    for p in people:
        rows.append({
            "name":        p["name"],
            "email":       p["email"],
            "designation": p["designation"],
            "brand1":      p["brands"][0] if len(p["brands"]) > 0 else "",
            "brand2":      p["brands"][1] if len(p["brands"]) > 1 else "",
            "brand3":      p["brands"][2] if len(p["brands"]) > 2 else "",
        })

    payload = {"action": "createRecipientsTab", "rows": rows}
    resp = requests.post(APPS_SCRIPT_URL, json=payload, timeout=30)
    resp.raise_for_status()
    result = resp.json()
    print(f"  Apps Script response: {result}")
    return result

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  Building Recipients Tab")
    print("=" * 55)

    print("\n[1/4] Fetching employee sheet...")
    emp_rows = fetch_csv_via_selenium(EMP_SHEET_ID, EMP_GID, "Employees")

    print("\n[2/4] Fetching brand mapping sheet...")
    brand_rows = fetch_csv_via_selenium(BRAND_SHEET_ID, BRAND_GID, "Brand Map")

    print("\n[3/4] Processing data...")
    employees  = parse_employees(emp_rows)
    brand_map  = parse_brand_map(brand_rows)

    print(f"\n  Active PMs/Directors found: {len(employees)}")
    for e in employees:
        brands = brand_map.get(e["name"].lower()) or fuzzy_lookup(e["name"], brand_map)
        e["brands"] = brands
        print(f"    {e['name']:40} {e['designation']:45} brands={brands}")

    print(f"\n[4/4] Writing {len(employees)} rows to '{OUTPUT_TAB}' tab...")
    write_recipients_tab(employees)

    print("\nDone! Check the Recipients tab in your Google Sheet.")

if __name__ == "__main__":
    main()
