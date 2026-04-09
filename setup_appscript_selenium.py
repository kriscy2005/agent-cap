"""
Opens the existing Apps Script project, pastes the code, saves, deploys.
Run with Edge fully closed: python setup_appscript_selenium.py
"""
import os, sys, time, re
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()

SCRIPT_ID  = "1P1sZMNyE4wFCmidQy2kRHEa1gb1ozuODaknOgYlqJlZxn2cdoLhzQccC"
EDITOR_URL = f"https://script.google.com/d/{SCRIPT_ID}/edit"
SHEET_ID   = os.getenv("SHEET_ID", "19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY")
ENV_FILE   = Path(__file__).parent / ".env"
SHOTS_DIR  = Path(__file__).parent / "debug_screenshots"
SHOTS_DIR.mkdir(exist_ok=True)

SCRIPT_CODE = (
    'const SHEET_ID = "' + SHEET_ID + '";\n'
    'const LOG_TAB  = "Bandwidth_Log";\n'
    'const HEADERS  = ["Ticket Key","Summary","Brand","Person Name","Designation","Role","Bandwidth %","Submitted At"];\n\n'
    'function doPost(e) {\n'
    '  try {\n'
    '    const data    = JSON.parse(e.postData.contents);\n'
    '    const entries = data.entries || [];\n'
    '    if (!entries.length) return _json({ error: "No entries" });\n'
    '    const ss    = SpreadsheetApp.openById(SHEET_ID);\n'
    '    let   sheet = ss.getSheetByName(LOG_TAB);\n'
    '    if (!sheet) {\n'
    '      sheet = ss.insertSheet(LOG_TAB);\n'
    '      sheet.appendRow(HEADERS);\n'
    '      sheet.getRange(1,1,1,HEADERS.length).setFontWeight("bold")\n'
    '           .setBackground("#0066CC").setFontColor("#ffffff");\n'
    '    }\n'
    '    const now  = new Date().toISOString();\n'
    '    const rows = entries.map(r => [\n'
    '      r.key, r.summary||"", r.brand||"",\n'
    '      r.personName, r.designation||"", r.role||"Other",\n'
    '      r.bandwidth, now\n'
    '    ]);\n'
    '    sheet.getRange(sheet.getLastRow()+1,1,rows.length,HEADERS.length).setValues(rows);\n'
    '    return _json({ success: true, rows: rows.length });\n'
    '  } catch(err) { return _json({ error: err.message }); }\n'
    '}\n'
    'function doGet()  { return _json({ status: "ok" }); }\n'
    'function _json(o) {\n'
    '  return ContentService.createTextOutput(JSON.stringify(o))\n'
    '    .setMimeType(ContentService.MimeType.JSON);\n'
    '}\n'
)

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException

EDGE_PROFILE = r"C:\Users\i_ojasvi.singh\AppData\Local\Microsoft\Edge\User Data"

def shot(driver, name):
    p = str(SHOTS_DIR / f"{name}.png")
    try:
        driver.save_screenshot(p)
        print(f"  [screenshot: {name}.png]")
    except:
        pass

def try_click(driver, selectors, timeout=15):
    for by, sel in selectors:
        try:
            el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, sel)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.3)
            el.click()
            return el
        except TimeoutException:
            continue
    return None

def save_url(url):
    content = ENV_FILE.read_text()
    if "APPS_SCRIPT_URL=" in content:
        content = re.sub(r"APPS_SCRIPT_URL=.*", f"APPS_SCRIPT_URL={url}", content)
    else:
        content += f"\nAPPS_SCRIPT_URL={url}\n"
    ENV_FILE.write_text(content)
    print(f"  Saved to .env: APPS_SCRIPT_URL={url}")

def wait_for_editor(driver, timeout=60):
    """Wait until the Monaco editor is loaded and ready."""
    print("  Waiting for Monaco editor...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            result = driver.execute_script("""
                try {
                    var m = monaco.editor.getModels();
                    return m && m.length > 0 ? 'ready:' + m.length : 'no_models';
                } catch(e) { return 'not_loaded'; }
            """)
            if str(result).startswith("ready"):
                print(f"  Monaco ready ({result})")
                return True
        except:
            pass
        time.sleep(2)
    return False

def inject_code(driver):
    """Set code in Monaco editor."""
    result = driver.execute_script("""
        try {
            var models = monaco.editor.getModels();
            if (models && models.length > 0) {
                models[0].setValue(arguments[0]);
                return 'ok';
            }
            return 'no_models';
        } catch(e) { return 'err:' + e.message; }
    """, SCRIPT_CODE)
    return result

def main():
    print("=" * 55)
    print("  Bandwidth Bot - Apps Script Setup")
    print("=" * 55)

    opts = Options()
    opts.add_argument(f"--user-data-dir={EDGE_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_experimental_option("detach", True)

    print("\n[1/4] Launching Edge...")
    driver = webdriver.Edge(options=opts)
    driver.maximize_window()

    try:
        # ── Step 1: Sign in first via accounts.google.com ────────────
        print("[2/4] Checking Google sign-in...")
        driver.get("https://accounts.google.com/ServiceLogin?continue=" +
                   "https://script.google.com")
        time.sleep(5)
        shot(driver, "01_login_check")

        current = driver.current_url
        if "script.google.com" not in current and "myaccount" not in current:
            # Need to sign in - wait up to 2 minutes
            print("  Please sign into your Google account in the Edge window...")
            print("  Waiting up to 2 minutes...")
            try:
                WebDriverWait(driver, 120).until(
                    lambda d: "script.google.com" in d.current_url
                              or "myaccount.google.com" in d.current_url
                )
                print("  Sign-in detected.")
            except TimeoutException:
                print("  Timed out waiting for sign-in. Proceeding anyway...")
        else:
            print("  Already signed in.")

        # ── Step 2: Navigate directly to the editor ──────────────────
        print(f"\n[3/4] Opening editor: {EDITOR_URL}")
        driver.get(EDITOR_URL)
        time.sleep(5)
        shot(driver, "02_editor_navigate")

        # If redirected to sign-in again, wait
        if "accounts.google.com" in driver.current_url:
            print("  Sign-in required again, waiting...")
            WebDriverWait(driver, 120).until(
                lambda d: "script.google.com" in d.current_url
            )
            time.sleep(5)

        shot(driver, "03_editor_loaded")

        # Wait for Monaco to be ready
        if not wait_for_editor(driver, timeout=60):
            shot(driver, "error_no_monaco")
            print("\n  Monaco editor not detected. Check debug_screenshots/03_editor_loaded.png")
            print("  The browser is left open — you can paste the code manually.")
            return

        # ── Step 3: Inject code ──────────────────────────────────────
        print("\n  Injecting code into editor...")
        result = inject_code(driver)
        print(f"  Result: {result}")

        if result != "ok":
            print("  Retrying in 3s...")
            time.sleep(3)
            result = inject_code(driver)
            print(f"  Retry result: {result}")

        # Save
        print("  Saving (Ctrl+S)...")
        time.sleep(1)
        body = driver.find_element(By.TAG_NAME, "body")
        body.click()
        time.sleep(0.3)
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("s").key_up(Keys.CONTROL).perform()
        time.sleep(4)
        shot(driver, "04_saved")

        # ── Step 4: Deploy ────────────────────────────────────────────
        print("\n[4/4] Clicking Deploy...")

        deployed = try_click(driver, [
            (By.XPATH, "//div[@role='button'][.//span[normalize-space(text())='Deploy']]"),
            (By.XPATH, "//button[.//span[normalize-space(text())='Deploy']]"),
            (By.XPATH, "//span[normalize-space(text())='Deploy']/parent::div[@role='button']"),
            (By.XPATH, "//span[normalize-space(text())='Deploy']/parent::button"),
            (By.XPATH, "//*[@data-tooltip='Deploy' or @aria-label='Deploy']"),
        ], timeout=15)

        if not deployed:
            shot(driver, "error_no_deploy")
            print("\n  Could not find Deploy button automatically.")
            print("  The browser is open — please manually:")
            print("  1. Click Deploy -> New deployment")
            print("  2. Type: Web app, Access: Anyone, click Deploy")
            print("  3. Copy the Web App URL")
            url = input("\n  Paste the Web App URL here: ").strip()
            if url:
                save_url(url)
            return

        shot(driver, "05_deploy_menu")
        time.sleep(2)

        # New deployment
        try_click(driver, [
            (By.XPATH, "//*[normalize-space(text())='New deployment']"),
            (By.XPATH, "//*[contains(text(),'New deployment')]"),
        ], timeout=12)
        time.sleep(3)
        shot(driver, "06_new_deploy_dialog")

        # Type = Web app
        try_click(driver, [
            (By.XPATH, "//div[@aria-label='Select type']"),
            (By.XPATH, "//*[@jsaction and contains(@class,'type')][@role='button']"),
        ], timeout=8)
        time.sleep(1)
        try_click(driver, [
            (By.XPATH, "//*[normalize-space(text())='Web app']"),
            (By.XPATH, "//li[contains(.,'Web app')]"),
        ], timeout=8)
        time.sleep(1)

        # Access = Anyone
        try:
            from selenium.webdriver.support.ui import Select as Sel
            for s in driver.find_elements(By.TAG_NAME, "select"):
                opts = [o.text for o in s.find_elements(By.TAG_NAME, "option")]
                if any("Anyone" in o for o in opts):
                    Sel(s).select_by_visible_text("Anyone")
                    print("  Set access to Anyone")
                    break
        except:
            pass

        try_click(driver, [
            (By.XPATH, "//div[@role='option'][contains(.,'Anyone')]"),
            (By.XPATH, "//li[normalize-space()='Anyone']"),
            (By.XPATH, "//*[normalize-space(text())='Anyone'][@role='option']"),
        ], timeout=8)
        time.sleep(1)

        shot(driver, "07_access_set")

        # Final Deploy
        try_click(driver, [
            (By.XPATH, "//button[.//span[normalize-space(text())='Deploy']]"),
            (By.XPATH, "//div[@role='button'][.//span[normalize-space(text())='Deploy']]"),
        ], timeout=15)

        time.sleep(8)
        shot(driver, "08_deployed")

        # ── Grab URL ──────────────────────────────────────────────────
        web_app_url = None
        for el in driver.find_elements(By.TAG_NAME, "input"):
            v = el.get_attribute("value") or ""
            if "script.google.com/macros/s/" in v:
                web_app_url = v
                break

        if not web_app_url:
            m = re.search(
                r'https://script\.google\.com/macros/s/[A-Za-z0-9_\-]+/exec',
                driver.page_source
            )
            if m:
                web_app_url = m.group(0)

        if web_app_url:
            save_url(web_app_url)
            print(f"\nSUCCESS!")
            print(f"URL: {web_app_url}")
            print("\nRestart the bot server to activate sheet writes.")
        else:
            shot(driver, "error_no_url")
            print("\nDeployed but URL not captured. Copy it from the browser:")
            url = input("Web App URL: ").strip()
            if url:
                save_url(url)

    except KeyboardInterrupt:
        print("\nCancelled.")
    except Exception as ex:
        shot(driver, "error_exception")
        print(f"\nError: {ex}")
        print("Browser left open for manual completion.")

if __name__ == "__main__":
    main()
