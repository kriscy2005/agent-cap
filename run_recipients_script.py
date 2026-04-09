"""
Opens Apps Script editor via Selenium, injects buildRecipientsTab code, saves,
then runs it directly — no deployment needed.
"""
import time, json
from pathlib import Path

SCRIPT_ID    = "1P1sZMNyE4wFCmidQy2kRHEa1gb1ozuODaknOgYlqJlZxn2cdoLhzQccC"
EDGE_PROFILE = r"C:\Users\i_ojasvi.singh\AppData\Local\Microsoft\Edge\User Data"
SHOTS_DIR    = Path(__file__).parent / "debug_screenshots"

SCRIPT_CODE = r"""
const MAIN_SHEET_ID  = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const EMP_SHEET_ID   = '1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8';
const BRAND_SHEET_ID = '1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho';
const EMP_GID        = '1165838487';
const BRAND_GID      = '1996161867';
const OUTPUT_TAB     = 'Recipients';
const ALLOWED = ['product manager','associate product manager','director','associate director','platform','product management','vp of product','vp, product','head of product'];
const LOG_TAB  = 'Bandwidth_Log';
const LOG_HEADERS = ['Ticket Key','Summary','Brand','Person Name','Designation','Role','Bandwidth %','Submitted At'];

function buildRecipientsTab() {
  Logger.log('START buildRecipientsTab');
  const empRows   = _readCsv(EMP_SHEET_ID, EMP_GID);
  const brandRows = _readCsv(BRAND_SHEET_ID, BRAND_GID);
  Logger.log('Emp rows: ' + empRows.length);
  Logger.log('Brand rows: ' + brandRows.length);
  const emailMap = _buildEmailMap(empRows);
  const people   = _parseBrandSheet(brandRows, emailMap);
  Logger.log('People to write: ' + people.length);
  _writeTab(people);
  Logger.log('DONE');
}

function _readCsv(sheetId, gid) {
  const url  = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?format=csv&gid=' + gid;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return Utilities.parseCsv(resp.getContentText());
}

function _buildEmailMap(rows) {
  const h = rows[0].map(x => x.toString().trim().toLowerCase());
  const iName   = _col(h, ['display name','name','employee name','full name']);
  const iEmail  = _col(h, ['email','work email','email address','corporate email','innovaccer email','official email','business email','user email','primary email']);
  const iStatus = _col(h, ['employee status','status','active','employment status']);
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];
    const name   = iName   >= 0 ? (row[iName]   || '').toString().trim() : '';
    const email  = iEmail  >= 0 ? (row[iEmail]  || '').toString().trim() : '';
    const status = iStatus >= 0 ? (row[iStatus] || '').toString().trim().toLowerCase() : 'active';
    if (!name) continue;
    if (iStatus >= 0 && status && !status.includes('active') && status !== '') continue;
    if (email) map[name.toLowerCase()] = email;
  }
  return map;
}

function _parseBrandSheet(rows, emailMap) {
  const h = rows[0].map(x => x.toString().trim().toLowerCase());
  const iName  = _col(h, ['name','employee name','full name','pm name','member name']);
  const iDesig = _col(h, ['designation','title','job title','role','position']);
  const iB1    = _col(h, ['solution 1','brand 1','brand','solution','brand name','account','product']);
  const iB2    = _col(h, ['solution 2','brand 2','brand2','secondary brand']);
  const iB3    = _col(h, ['solution 3','brand 3','brand3']);
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const row   = rows[i];
    const name  = iName  >= 0 ? (row[iName]  || '').toString().trim() : '';
    const desig = iDesig >= 0 ? (row[iDesig] || '').toString().trim() : '';
    if (!name) continue;
    const dl = desig.toLowerCase();
    if (!ALLOWED.some(a => dl.includes(a))) continue;
    const brands = [];
    if (iB1 >= 0 && (row[iB1] || '').toString().trim()) brands.push((row[iB1] || '').toString().trim());
    if (iB2 >= 0 && (row[iB2] || '').toString().trim()) brands.push((row[iB2] || '').toString().trim());
    if (iB3 >= 0 && (row[iB3] || '').toString().trim()) brands.push((row[iB3] || '').toString().trim());
    let email = emailMap[name.toLowerCase()] || '';
    if (!email) {
      const parts = name.toLowerCase().split(/\s+/).filter(Boolean);
      if (parts.length >= 2) {
        const first = parts[0], last = parts[parts.length - 1];
        for (const k of Object.keys(emailMap)) {
          const kp = k.split(/\s+/);
          if (kp.length >= 2 && kp[0] === first && kp[kp.length-1] === last) { email = emailMap[k]; break; }
        }
      }
    }
    result.push({ name, email, designation: desig, brands });
  }
  return result;
}

function _writeTab(people) {
  const ss  = SpreadsheetApp.openById(MAIN_SHEET_ID);
  let tab   = ss.getSheetByName(OUTPUT_TAB);
  if (tab) { tab.clearContents(); } else { tab = ss.insertSheet(OUTPUT_TAB); }
  const headers = ['Name','Email','Designation','Brand 1','Brand 2','Brand 3'];
  tab.appendRow(headers);
  tab.getRange(1,1,1,headers.length).setBackground('#0066CC').setFontColor('#ffffff').setFontWeight('bold');
  const data = people.map(p => [p.name, p.email, p.designation, p.brands[0]||'', p.brands[1]||'', p.brands[2]||'']);
  if (data.length) tab.getRange(2,1,data.length,headers.length).setValues(data);
  tab.setFrozenRows(1);
  for (let c = 1; c <= headers.length; c++) tab.autoResizeColumn(c);
  Logger.log('Written ' + people.length + ' rows to ' + OUTPUT_TAB);
}

function _col(header, candidates) {
  for (const c of candidates) {
    const i = header.indexOf(c);
    if (i >= 0) return i;
  }
  for (const c of candidates) {
    const i = header.findIndex(h => h.includes(c) || c.includes(h));
    if (i >= 0) return i;
  }
  return -1;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'createRecipientsTab') {
      const rows = data.rows || [];
      const ss = SpreadsheetApp.openById(MAIN_SHEET_ID);
      const RCPT_HEADERS = ['Name','Email','Designation','Brand 1','Brand 2','Brand 3'];
      let sheet = ss.getSheetByName(OUTPUT_TAB);
      if (sheet) { sheet.clearContents(); } else { sheet = ss.insertSheet(OUTPUT_TAB); }
      sheet.appendRow(RCPT_HEADERS);
      sheet.getRange(1,1,1,RCPT_HEADERS.length).setFontWeight('bold').setBackground('#0066CC').setFontColor('#ffffff');
      const d = rows.map(r => [r.name, r.email, r.designation, r.brand1||'', r.brand2||'', r.brand3||'']);
      if (d.length) sheet.getRange(2,1,d.length,RCPT_HEADERS.length).setValues(d);
      sheet.setFrozenRows(1);
      for (let c=1; c<=RCPT_HEADERS.length; c++) sheet.autoResizeColumn(c);
      return _json({ success: true, tab: OUTPUT_TAB, rows: d.length });
    }
    const entries = data.entries || [];
    if (!entries.length) return _json({ error: 'No entries' });
    const ss    = SpreadsheetApp.openById(MAIN_SHEET_ID);
    let   sheet = ss.getSheetByName(LOG_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_TAB);
      sheet.appendRow(LOG_HEADERS);
      sheet.getRange(1,1,1,LOG_HEADERS.length).setFontWeight('bold').setBackground('#0066CC').setFontColor('#ffffff');
    }
    const now  = new Date().toISOString();
    const rows = entries.map(r => [r.key, r.summary||'', r.brand||'', r.personName, r.designation||'', r.role||'Other', r.bandwidth, now]);
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
    return _json({ success: true, rows: rows.length });
  } catch(err) { return _json({ error: err.message }); }
}

function doGet() { return _json({ status: 'ok' }); }
function _json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
""".strip()


def make_driver():
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
    return webdriver.Edge(options=opts)


def shot(driver, name):
    SHOTS_DIR.mkdir(exist_ok=True)
    try: driver.save_screenshot(str(SHOTS_DIR / f"run_{name}.png"))
    except: pass


def safe_print(*args):
    """Print with Unicode chars replaced to avoid cp1252 errors."""
    msg = " ".join(str(a) for a in args)
    print(msg.encode('ascii', errors='replace').decode('ascii'))


def main():
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains

    print("Opening Apps Script editor...")
    driver = make_driver()
    try:
        driver.maximize_window()
        driver.get(f"https://script.google.com/d/{SCRIPT_ID}/edit")

        # Wait for Monaco
        print("Waiting for editor...")
        deadline = time.time() + 90
        while time.time() < deadline:
            try:
                ready = driver.execute_script(
                    "try { return monaco.editor.getModels().length > 0 ? 'ready' : 'wait'; } catch(e) { return 'loading'; }"
                )
                if ready == 'ready':
                    break
            except: pass
            time.sleep(2)
        else:
            raise RuntimeError("Monaco editor never loaded")

        time.sleep(2)
        shot(driver, "01_loaded")
        print("Editor loaded. Injecting code...")

        # Inject code
        driver.execute_script(f"monaco.editor.getModels()[0].setValue({json.dumps(SCRIPT_CODE)});")
        time.sleep(1)

        # Save
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
        time.sleep(3)
        shot(driver, "02_saved")
        print("Code saved.")

        # Click "Run" menu in the toolbar
        print("Clicking Run menu...")
        run_menu = driver.execute_script("""
            var items = Array.from(document.querySelectorAll('div, li, span, a'));
            for (var i=0; i<items.length; i++) {
                var t = (items[i].innerText || '').trim();
                if (t === 'Run') {
                    var r = items[i].getBoundingClientRect();
                    if (r.width > 0 && r.height > 0 && r.y < 100) {
                        items[i].click();
                        return 'clicked:' + items[i].tagName + ':' + items[i].className;
                    }
                }
            }
            return 'not_found';
        """)
        print(f"  Run menu: {run_menu}")
        time.sleep(1)
        shot(driver, "03_run_menu")

        # Look for "Run function" or "buildRecipientsTab" in the menu
        fn_result = driver.execute_script("""
            var items = Array.from(document.querySelectorAll('*'));
            for (var i=0; i<items.length; i++) {
                var t = (items[i].innerText || '').trim();
                if (t === 'buildRecipientsTab' || t.includes('buildRecipients')) {
                    var r = items[i].getBoundingClientRect();
                    if (r.width > 0) { items[i].click(); return 'clicked:' + t; }
                }
            }
            // Look for "Run function" submenu
            for (var i=0; i<items.length; i++) {
                var t = (items[i].innerText || '').trim();
                if (t === 'Run function') {
                    items[i].click();
                    return 'clicked_run_function';
                }
            }
            return 'not_found';
        """)
        print(f"  Function select: {fn_result}")
        time.sleep(1)
        shot(driver, "04_fn_select")

        # If "Run function" was clicked, now look for buildRecipientsTab
        if fn_result == 'clicked_run_function':
            fn2 = driver.execute_script("""
                var items = Array.from(document.querySelectorAll('*'));
                for (var i=0; i<items.length; i++) {
                    var t = (items[i].innerText || '').trim();
                    if (t.includes('buildRecipients')) {
                        var r = items[i].getBoundingClientRect();
                        if (r.width > 0) { items[i].click(); return 'clicked:' + t; }
                    }
                }
                return 'not_found';
            """)
            print(f"  buildRecipientsTab click: {fn2}")
            time.sleep(1)

        # Try toolbar Run button (play icon)
        toolbar_run = driver.execute_script("""
            var btns = document.querySelectorAll('button, [role="button"]');
            for (var i=0; i<btns.length; i++) {
                var a = (btns[i].getAttribute('aria-label') || '').toLowerCase();
                var t = (btns[i].getAttribute('title') || '').toLowerCase();
                if (a === 'run' || t === 'run') {
                    var r = btns[i].getBoundingClientRect();
                    if (r.width > 0) { btns[i].click(); return 'toolbar_run_clicked'; }
                }
            }
            // SVG play button (triangle)
            var svgs = document.querySelectorAll('svg path[d*="M8"], svg path[d*="M6"]');
            for (var i=0; i<svgs.length; i++) {
                var btn = svgs[i].closest('button, [role="button"]');
                if (btn) { btn.click(); return 'svg_run_clicked'; }
            }
            return 'not_found';
        """)
        print(f"  Toolbar run: {toolbar_run}")
        time.sleep(2)
        shot(driver, "05_running")

        # Wait for execution to complete
        print("Waiting for script to run (up to 180 seconds)...")
        deadline = time.time() + 180
        completed = False
        while time.time() < deadline:
            status = driver.execute_script("""
                // Check for execution log or dialog
                var body = document.body.innerText || '';
                if (body.includes('Written') && body.includes('rows')) return 'written';
                if (body.includes('Execution completed')) return 'completed';
                if (body.includes('Recipients')) return 'recipients_mentioned';
                if (body.includes('Error') && body.includes('line')) return 'error';
                // Check for running indicator
                if (body.includes('Running') || body.includes('Executing')) return 'running';
                return 'unknown';
            """)
            print(f"  [{int(time.time() % 100)}] Status: {status}")
            if status in ('written', 'completed', 'recipients_mentioned'):
                completed = True
                break
            if status == 'error':
                shot(driver, "06_error")
                break
            time.sleep(5)

        shot(driver, "07_final")

        # Capture log
        log = driver.execute_script("""
            var body = document.body.innerText || '';
            // Find the execution log section
            var panels = document.querySelectorAll('[class*="log"], [class*="console"], [class*="output"], [class*="panel"]');
            var logs = Array.from(panels).map(p => p.innerText).filter(t => t && t.length > 0);
            return logs.length > 0 ? logs.join('\\n---\\n').substring(0, 1000) : body.substring(0, 500);
        """)
        print(f"\nExecution output:\n{log[:500]}")

        if completed:
            print("\nSUCCESS: buildRecipientsTab ran! Check the Recipients tab in Google Sheets.")
        else:
            print("\nScript may still be running or completed — check the Recipients tab manually.")
            print("Screenshots saved to:", SHOTS_DIR)

        time.sleep(8)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
