"""
Full automation:
 1. Download employee + brand sheets via Selenium (logged-in Edge)
 2. Update Apps Script code + redeploy
 3. POST processed data to create Recipients tab
"""
import os, csv, io, sys, time, re, json, requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(override=True)

MAIN_SHEET_ID   = "19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY"
EMP_SHEET_ID    = "1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8"
EMP_GID         = "1165838487"
BRAND_SHEET_ID  = "1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho"
BRAND_GID       = "1996161867"
SCRIPT_ID       = "1P1sZMNyE4wFCmidQy2kRHEa1gb1ozuODaknOgYlqJlZxn2cdoLhzQccC"
APPS_SCRIPT_URL = os.getenv("APPS_SCRIPT_URL", "")
EDGE_PROFILE    = r"C:\Users\i_ojasvi.singh\AppData\Local\Microsoft\Edge\User Data"
DOWNLOAD_DIR    = str(Path(__file__).parent / "tmp_downloads")
SHOTS_DIR       = Path(__file__).parent / "debug_screenshots"

ALLOWED_DESIGNATIONS = [
    "product manager",
    "associate product manager",
    "director",
    "associate director",
    "platform",
    "product management",
    "vp of product",
    "vp, product",
    "head of product",
]

NEW_APPS_SCRIPT = r"""
const SHEET_ID = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const LOG_TAB  = 'Bandwidth_Log';
const HEADERS  = ['Ticket Key','Summary','Brand','Person Name','Designation','Role','Bandwidth %','Submitted At'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'createRecipientsTab') return _createRecipientsTab(data.rows || []);
    const entries = data.entries || [];
    if (!entries.length) return _json({ error: 'No entries' });
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(LOG_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_TAB);
      sheet.appendRow(HEADERS);
      sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#0066CC').setFontColor('#ffffff');
    }
    const now  = new Date().toISOString();
    const rows = entries.map(r => [r.key, r.summary||'', r.brand||'', r.personName, r.designation||'', r.role||'Other', r.bandwidth, now]);
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, HEADERS.length).setValues(rows);
    return _json({ success: true, rows: rows.length });
  } catch(err) { return _json({ error: err.message }); }
}

function _createRecipientsTab(rows) {
  const RCPT_TAB = 'Recipients';
  const RCPT_HEADERS = ['Name','Email','Designation','Brand 1','Brand 2','Brand 3'];
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(RCPT_TAB);
  if (sheet) { sheet.clearContents(); } else { sheet = ss.insertSheet(RCPT_TAB); }
  sheet.appendRow(RCPT_HEADERS);
  sheet.getRange(1,1,1,RCPT_HEADERS.length).setFontWeight('bold').setBackground('#0066CC').setFontColor('#ffffff');
  const data = rows.map(r => [r.name, r.email, r.designation, r.brand1||'', r.brand2||'', r.brand3||'']);
  if (data.length) sheet.getRange(2,1,data.length,RCPT_HEADERS.length).setValues(data);
  sheet.setFrozenRows(1);
  for (let c=1; c<=RCPT_HEADERS.length; c++) sheet.autoResizeColumn(c);
  return _json({ success: true, tab: RCPT_TAB, rows: data.length });
}

function doGet() { return _json({ status: 'ok', tab: LOG_TAB }); }
function _json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
""".strip()

# ── Selenium helpers ──────────────────────────────────────────────────────────
def make_driver(download=False):
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options
    opts = Options()
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    if download:
        Path(DOWNLOAD_DIR).mkdir(exist_ok=True)
        opts.add_experimental_option("prefs", {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        })
    return webdriver.Edge(options=opts)

def shot(driver, name):
    SHOTS_DIR.mkdir(exist_ok=True)
    try:
        driver.save_screenshot(str(SHOTS_DIR / f"recip_{name}.png"))
    except: pass

def click_js(driver, text, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = driver.execute_script(f"""
            var els = document.querySelectorAll('*');
            for (var i=0; i<els.length; i++) {{
                var t = (els[i].innerText||'').trim();
                if (t === {json.dumps(text)}) {{
                    var r = els[i].getBoundingClientRect();
                    if (r.width>0 && r.height>0) {{ els[i].click(); return 'ok'; }}
                }}
            }}
            return 'not_found';
        """)
        if result == 'ok':
            return True
        time.sleep(0.8)
    return False

# ── Step 1: Download CSVs ─────────────────────────────────────────────────────
def download_csv(driver, sheet_id, gid, label, save_as):
    # Clear old downloads
    for f in Path(DOWNLOAD_DIR).glob("*.csv"):
        f.unlink(missing_ok=True)

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    print(f"  Fetching {label}...")
    driver.get(url)
    deadline = time.time() + 30
    while time.time() < deadline:
        csvs = list(Path(DOWNLOAD_DIR).glob("*.csv"))
        if csvs:
            time.sleep(0.8)
            dest = Path(DOWNLOAD_DIR) / save_as
            csvs[0].rename(dest)
            content = dest.read_text(encoding="utf-8", errors="replace")
            rows = list(csv.reader(io.StringIO(content)))
            print(f"  Got {len(rows)} rows, headers: {rows[0][:6]}")
            return rows
        time.sleep(1)
    raise RuntimeError(f"Download timed out for {label}")

# ── Step 2: Update Apps Script ────────────────────────────────────────────────
def update_apps_script():
    print("\n[2/4] Updating Apps Script...")
    driver = make_driver(download=False)
    try:
        driver.maximize_window()
        editor_url = f"https://script.google.com/d/{SCRIPT_ID}/edit"
        driver.get(editor_url)
        print("  Waiting for Monaco editor...")
        deadline = time.time() + 60
        while time.time() < deadline:
            try:
                ready = driver.execute_script("""
                    try { return monaco.editor.getModels().length > 0 ? 'ready' : 'wait'; }
                    catch(e) { return 'loading'; }
                """)
                if ready == 'ready':
                    break
            except: pass
            time.sleep(2)

        time.sleep(2)
        shot(driver, "01_editor")

        # Inject new code
        code_escaped = json.dumps(NEW_APPS_SCRIPT)
        driver.execute_script(f"monaco.editor.getModels()[0].setValue({code_escaped});")
        time.sleep(1)

        # Save
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
        time.sleep(3)
        shot(driver, "02_saved")
        print("  Code saved.")

        # Deploy -> Manage deployments
        print("  Opening Manage Deployments...")
        driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var i=0; i<els.length; i++) {
                var t = (els[i].innerText||'').trim();
                if (t === 'Deploy') {
                    var r = els[i].getBoundingClientRect();
                    if (r.width > 30) { els[i].click(); return; }
                }
            }
        """)
        time.sleep(2)
        shot(driver, "03_deploy_menu")

        clicked = click_js(driver, "Manage deployments", timeout=8)
        if not clicked:
            driver.execute_script("""
                var els = document.querySelectorAll('*');
                for (var i=0; i<els.length; i++) {
                    if ((els[i].innerText||'').includes('Manage')) {
                        var r = els[i].getBoundingClientRect();
                        if (r.width>0) { els[i].click(); return; }
                    }
                }
            """)
        time.sleep(3)
        shot(driver, "04_manage_dialog")

        driver.execute_script("""
            var btns = document.querySelectorAll('button, [role="button"], svg, [aria-label]');
            for (var i=0; i<btns.length; i++) {
                var label = (btns[i].getAttribute('aria-label')||'').toLowerCase();
                var title = (btns[i].getAttribute('title')||'').toLowerCase();
                if (label.includes('edit') || title.includes('edit')) {
                    btns[i].click(); return;
                }
            }
            var paths = document.querySelectorAll('path');
            for (var i=0; i<paths.length; i++) {
                var parent = paths[i].closest('button, [role="button"]');
                if (parent) { parent.click(); return; }
            }
        """)
        time.sleep(2)
        shot(driver, "05_edit_clicked")

        click_js(driver, "New version", timeout=6)
        time.sleep(1)

        driver.execute_script("""
            var btns = document.querySelectorAll('button, [role="button"]');
            for (var i=btns.length-1; i>=0; i--) {
                var t = (btns[i].innerText||'').trim();
                if (t === 'Deploy') {
                    var r = btns[i].getBoundingClientRect();
                    if (r.width>0 && r.y>100) { btns[i].click(); return; }
                }
            }
        """)
        time.sleep(6)
        shot(driver, "06_deployed")

        click_js(driver, "Done", timeout=8)
        time.sleep(1)
        print("  Apps Script updated and redeployed.")
    finally:
        driver.quit()

# ── Column helpers ────────────────────────────────────────────────────────────
def col_idx(header, candidates):
    """Exact match first, then partial."""
    for c in candidates:
        for i, h in enumerate(header):
            if c == h: return i
    for c in candidates:
        for i, h in enumerate(header):
            if c in h or h in c: return i
    return -1

def safe(row, idx):
    return row[idx].strip() if 0 <= idx < len(row) else ""

# ── Build email map from employee sheet ──────────────────────────────────────
def build_email_map(emp_rows):
    """
    Employee sheet actual headers include:
      'Display Name', 'Employee Status', and some email column.
    Returns dict: lowercase_name -> email
    """
    header = [h.strip().lower() for h in emp_rows[0]]
    i_name   = col_idx(header, ["display name", "name", "employee name", "full name", "emp name"])
    i_email  = col_idx(header, ["email", "work email", "email address", "corporate email",
                                 "innovaccer email", "official email", "business email",
                                 "user email", "primary email"])
    i_status = col_idx(header, ["employee status", "status", "active", "employment status", "emp status"])
    print(f"  Emp cols -> name:{i_name} email:{i_email} status:{i_status}")

    email_map = {}
    for row in emp_rows[1:]:
        name   = safe(row, i_name)
        email  = safe(row, i_email)
        status = safe(row, i_status).lower()
        if not name:
            continue
        # Skip inactive
        if i_status >= 0 and status and "active" not in status and status != "":
            continue
        if email:
            email_map[name.lower().strip()] = email
    print(f"  Email map size: {len(email_map)}")
    return email_map

# ── Parse brand/designation sheet ────────────────────────────────────────────
def parse_brand_sheet(brand_rows, email_map):
    """
    Brand mapping sheet actual headers:
      EMP ID, Name, Region, Designation, Solution 1, Solution 2, Solution 3
    Returns list of dicts: {name, email, designation, brands:[]}
    """
    header = [h.strip().lower() for h in brand_rows[0]]
    i_name  = col_idx(header, ["name", "employee name", "full name", "pm name", "member name"])
    i_desig = col_idx(header, ["designation", "title", "job title", "role", "position"])
    i_b1    = col_idx(header, ["solution 1", "brand 1", "brand", "solution", "brand name", "account", "product"])
    i_b2    = col_idx(header, ["solution 2", "brand 2", "brand2", "secondary brand"])
    i_b3    = col_idx(header, ["solution 3", "brand 3", "brand3"])
    print(f"  Brand cols -> name:{i_name} desig:{i_desig} b1:{i_b1} b2:{i_b2} b3:{i_b3}")

    result = []
    for row in brand_rows[1:]:
        name  = safe(row, i_name)
        desig = safe(row, i_desig)
        if not name:
            continue
        if not any(a in desig.lower() for a in ALLOWED_DESIGNATIONS):
            continue
        brands = [b for b in [safe(row, i_b1), safe(row, i_b2), safe(row, i_b3)] if b]
        # Look up email from employee sheet
        email = email_map.get(name.lower().strip(), "")
        if not email:
            # Try first+last name fuzzy match
            parts = name.lower().split()
            if len(parts) >= 2:
                first, last = parts[0], parts[-1]
                for k, v in email_map.items():
                    kparts = k.split()
                    if len(kparts) >= 2 and first == kparts[0] and last == kparts[-1]:
                        email = v
                        break
        result.append({"name": name, "email": email, "designation": desig, "brands": brands})
    return result

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  Build Recipients Tab -- Full Automation")
    print("=" * 55)

    # Step 1: Download both sheets
    print("\n[1/4] Downloading source sheets via Edge...")
    driver = make_driver(download=True)
    try:
        driver.maximize_window()
        emp_rows   = download_csv(driver, EMP_SHEET_ID,   EMP_GID,   "Employee Sheet",      "employees.csv")
        brand_rows = download_csv(driver, BRAND_SHEET_ID, BRAND_GID, "Brand Mapping Sheet", "brands.csv")
    finally:
        driver.quit()

    # Step 2: Update Apps Script
    update_apps_script()

    # Step 3: Process data
    print("\n[3/4] Processing data...")
    email_map = build_email_map(emp_rows)
    people    = parse_brand_sheet(brand_rows, email_map)
    print(f"  Filtered people (PMs/Directors): {len(people)}")

    for p in people:
        b_str = str(p["brands"]) if p["brands"] else "(none)"
        print(f"    {p['name'][:40]:40} {p['designation'][:35]:35} brands={b_str}")

    # Step 4: Write Recipients tab
    print(f"\n[4/4] Writing {len(people)} rows to Recipients tab...")
    if not APPS_SCRIPT_URL:
        print("ERROR: APPS_SCRIPT_URL not set in .env")
        sys.exit(1)

    payload = {
        "action": "createRecipientsTab",
        "rows": [
            {
                "name":        p["name"],
                "email":       p["email"],
                "designation": p["designation"],
                "brand1":      p["brands"][0] if len(p["brands"]) > 0 else "",
                "brand2":      p["brands"][1] if len(p["brands"]) > 1 else "",
                "brand3":      p["brands"][2] if len(p["brands"]) > 2 else "",
            }
            for p in people
        ]
    }

    resp = requests.post(APPS_SCRIPT_URL, json=payload, timeout=30)
    print(f"  HTTP {resp.status_code}: {resp.text[:200]}")

    if resp.status_code == 200:
        result = resp.json()
        if result.get("success"):
            print(f"\nSUCCESS! {result['rows']} people written to '{result['tab']}' tab.")
        else:
            print(f"\nApps Script error: {result}")
    else:
        print(f"\nHTTP error: {resp.text[:300]}")

if __name__ == "__main__":
    main()
