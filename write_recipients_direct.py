"""
Direct write to Google Sheet Recipients tab using Selenium + logged-in Edge.
Reads already-downloaded CSVs, processes data, opens the Google Sheet,
and writes rows via Apps Script editor execution (runs buildRecipientsTab).
"""
import os, csv, io, sys, time, json, requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(override=True)

MAIN_SHEET_ID   = "19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY"
EMP_SHEET_ID    = "1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8"
EMP_GID         = "1165838487"
BRAND_SHEET_ID  = "1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho"
BRAND_GID       = "1996161867"
SCRIPT_ID       = "1P1sZMNyE4wFCmidQy2kRHEa1gb1ozuODaknOgYlqJlZxn2cdoLhzQccC"
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

# ── Apps Script code to inject (includes buildRecipientsTab) ──────────────────
RECIPIENTS_SCRIPT = """
const MAIN_SHEET_ID  = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const EMP_SHEET_ID   = '1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8';
const BRAND_SHEET_ID = '1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho';
const EMP_GID        = '1165838487';
const BRAND_GID      = '1996161867';
const OUTPUT_TAB     = 'Recipients';
const ALLOWED = ['product manager','associate product manager','director','associate director','platform','product management','vp of product','vp, product','head of product'];

function buildRecipientsTab() {
  Logger.log('START');
  const empRows   = _readCsv(EMP_SHEET_ID, EMP_GID);
  const brandRows = _readCsv(BRAND_SHEET_ID, BRAND_GID);
  Logger.log('Emp rows: ' + empRows.length + '  Brand rows: ' + brandRows.length);
  Logger.log('Emp headers: ' + empRows[0].join(' | '));
  Logger.log('Brand headers: ' + brandRows[0].join(' | '));

  const emailMap = _buildEmailMap(empRows);
  Logger.log('Email map size: ' + Object.keys(emailMap).length);

  const people = _parseBrandSheet(brandRows, emailMap);
  Logger.log('Filtered people: ' + people.length);

  _writeTab(people);
  Logger.log('DONE');
  try { SpreadsheetApp.getUi().alert('Done! ' + people.length + ' recipients written.'); } catch(e) {}
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
  Logger.log('Emp col idx: name=' + iName + ' email=' + iEmail + ' status=' + iStatus);
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
  Logger.log('Brand col idx: name=' + iName + ' desig=' + iDesig + ' b1=' + iB1 + ' b2=' + iB2 + ' b3=' + iB3);
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
      const parts = name.toLowerCase().split(/\\s+/).filter(Boolean);
      if (parts.length >= 2) {
        const first = parts[0], last = parts[parts.length - 1];
        for (const k of Object.keys(emailMap)) {
          const kp = k.split(/\\s+/);
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

// Also keep the doPost handler for the bandwidth bot
const LOG_SHEET_ID = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const LOG_TAB  = 'Bandwidth_Log';
const HEADERS  = ['Ticket Key','Summary','Brand','Person Name','Designation','Role','Bandwidth %','Submitted At'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'createRecipientsTab') return _createRecipientsTabViaPost(data.rows || []);
    const entries = data.entries || [];
    if (!entries.length) return _json({ error: 'No entries' });
    const ss    = SpreadsheetApp.openById(LOG_SHEET_ID);
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

function _createRecipientsTabViaPost(rows) {
  const RCPT_TAB = 'Recipients';
  const RCPT_HEADERS = ['Name','Email','Designation','Brand 1','Brand 2','Brand 3'];
  const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
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
    try: driver.save_screenshot(str(SHOTS_DIR / f"direct_{name}.png"))
    except: pass

def wait_monaco(driver, timeout=60):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            ready = driver.execute_script("""
                try { return monaco.editor.getModels().length > 0 ? 'ready' : 'wait'; }
                catch(e) { return 'loading'; }
            """)
            if ready == 'ready':
                return True
        except: pass
        time.sleep(2)
    return False

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

def run_in_editor():
    """Open Apps Script editor, inject buildRecipientsTab code, run it."""
    print("\nOpening Apps Script editor to run buildRecipientsTab()...")
    driver = make_driver()
    try:
        driver.maximize_window()
        editor_url = f"https://script.google.com/d/{SCRIPT_ID}/edit"
        driver.get(editor_url)

        print("  Waiting for Monaco editor...")
        if not wait_monaco(driver, timeout=60):
            print("  Monaco did not load, retrying...")
            driver.refresh()
            if not wait_monaco(driver, timeout=60):
                raise RuntimeError("Monaco editor never loaded")

        time.sleep(2)
        shot(driver, "01_editor_loaded")

        # Inject the full script
        print("  Injecting script code...")
        code_escaped = json.dumps(RECIPIENTS_SCRIPT)
        driver.execute_script(f"monaco.editor.getModels()[0].setValue({code_escaped});")
        time.sleep(1)

        # Save
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('s').key_up(Keys.CONTROL).perform()
        time.sleep(3)
        shot(driver, "02_saved")
        print("  Code saved.")

        # Select function buildRecipientsTab from the dropdown
        print("  Selecting function buildRecipientsTab...")
        # Try clicking the function dropdown
        driver.execute_script("""
            // Look for the function selector dropdown
            var selects = document.querySelectorAll('select');
            for (var i=0; i<selects.length; i++) {
                var opts = selects[i].options;
                for (var j=0; j<opts.length; j++) {
                    if (opts[j].text.includes('buildRecipients')) {
                        selects[i].value = opts[j].value;
                        selects[i].dispatchEvent(new Event('change'));
                        return 'found_select';
                    }
                }
            }
            // Try the custom dropdown
            var items = document.querySelectorAll('[role="option"], [role="menuitem"], li');
            for (var i=0; i<items.length; i++) {
                if ((items[i].innerText||'').includes('buildRecipients')) {
                    items[i].click();
                    return 'found_item';
                }
            }
            return 'not_found';
        """)
        time.sleep(1)

        # Try to find and click the function name selector button
        driver.execute_script("""
            // Find the Run button dropdown for function selection
            var btns = document.querySelectorAll('div[role="button"], button');
            for (var i=0; i<btns.length; i++) {
                var t = (btns[i].innerText || btns[i].getAttribute('aria-label') || '').trim();
                if (t && (t.includes('function') || t.includes('Function') || t.includes('Select'))) {
                    console.log('Found selector:', t);
                }
            }
        """)

        # Click the function name in the toolbar (it's usually a dropdown)
        # Find the toolbar input that shows current function
        result = driver.execute_script("""
            // The function selector in Apps Script editor is a custom element
            // Look for an element that shows function names
            var allText = Array.from(document.querySelectorAll('div, span, button'))
                .filter(el => {
                    var t = (el.innerText || '').trim();
                    return t && t.length < 50 && (t.includes('doPost') || t.includes('doGet') || t.includes('main') || t.includes('build'));
                });
            return allText.map(el => ({tag: el.tagName, text: el.innerText, cls: el.className})).slice(0, 10);
        """)
        print("  Function selector candidates:", result)
        time.sleep(1)
        shot(driver, "03_before_run")

        # Click Run button to run the currently selected function
        # The Run button in Apps Script has a specific look - try multiple approaches
        run_clicked = driver.execute_script("""
            // Look for run button by aria-label or title
            var btns = document.querySelectorAll('button, div[role="button"], [aria-label]');
            for (var i=0; i<btns.length; i++) {
                var label = (btns[i].getAttribute('aria-label') || '').toLowerCase();
                var title = (btns[i].getAttribute('title') || '').toLowerCase();
                var text  = (btns[i].innerText || '').trim();
                if (label === 'run' || title === 'run' || text === 'Run') {
                    var r = btns[i].getBoundingClientRect();
                    if (r.width > 0) { btns[i].click(); return 'run_clicked:' + label + ':' + title + ':' + text; }
                }
            }
            return 'not_found';
        """)
        print(f"  Run click result: {run_clicked}")
        time.sleep(2)
        shot(driver, "04_after_run_click")

        # Now we need to select the right function - look for function list
        # In Apps Script, there's a function dropdown next to the Run button
        fn_result = driver.execute_script("""
            // Find function selector - it's usually near the run button
            var dropdowns = document.querySelectorAll('[role="combobox"], [role="listbox"], select');
            for (var i=0; i<dropdowns.length; i++) {
                var t = dropdowns[i].innerText || dropdowns[i].textContent || '';
                console.log('Dropdown:', i, t.substring(0, 100));
            }
            // Also look for function dropdown that might show function list
            var allEls = document.querySelectorAll('*');
            for (var i=0; i<allEls.length; i++) {
                var t = (allEls[i].innerText || '').trim();
                if (t === 'buildRecipientsTab') {
                    allEls[i].click();
                    return 'clicked_buildRecipientsTab';
                }
            }
            return 'not_found';
        """)
        print(f"  Function select: {fn_result}")
        time.sleep(1)

        # Try keyboard shortcut to run: Ctrl+R in some versions
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('r').key_up(Keys.CONTROL).perform()
        time.sleep(2)
        shot(driver, "05_ctrl_r")

        # Wait for execution
        print("  Waiting for script execution (up to 120 seconds)...")
        deadline = time.time() + 120
        last_log = ""
        while time.time() < deadline:
            # Check for success/completion indicators
            status = driver.execute_script("""
                var els = document.querySelectorAll('*');
                for (var i=0; i<els.length; i++) {
                    var t = (els[i].innerText || '').trim();
                    if (t.includes('Done!') || t.includes('recipients written') || t.includes('Execution completed') || t.includes('DONE')) {
                        return 'done:' + t.substring(0, 100);
                    }
                    if (t.includes('Error') && t.length < 200 && t.includes('line')) {
                        return 'error:' + t.substring(0, 200);
                    }
                }
                // Check execution log
                var logEls = document.querySelectorAll('.execution-log, [class*="log"], [class*="Log"]');
                if (logEls.length > 0) return 'log:' + (logEls[0].innerText || '').substring(0, 200);
                return 'running';
            """)
            if status != last_log and status != 'running':
                print(f"  Status: {status}")
                last_log = status
            if status.startswith('done:') or status.startswith('error:'):
                break
            time.sleep(3)

        shot(driver, "06_execution_result")

        # Get final log content
        log_content = driver.execute_script("""
            // Try to get execution log
            var logArea = document.querySelector('.execution-log, [class*="executionLog"], [class*="transcript"]');
            if (logArea) return logArea.innerText;
            // Try iframe content
            var iframes = document.querySelectorAll('iframe');
            for (var i=0; i<iframes.length; i++) {
                try {
                    var content = iframes[i].contentDocument.body.innerText;
                    if (content && content.length > 10) return content.substring(0, 500);
                } catch(e) {}
            }
            // Get any visible text that looks like output
            var bottom = document.querySelector('[class*="console"], [class*="output"], [class*="panel"]');
            if (bottom) return bottom.innerText.substring(0, 500);
            return 'No log captured';
        """)
        print(f"  Execution log: {log_content[:300] if log_content else 'none'}")

        print("\nScript execution attempted. Check the Google Sheet for the Recipients tab.")
        print("If it worked, you'll see the Recipients tab populated.")
        time.sleep(5)
    finally:
        driver.quit()


if __name__ == "__main__":
    run_in_editor()
