"""
Deploy the already-saved Apps Script as a Web App.
Run AFTER setup_appscript_selenium.py has saved the code.
"""
import os, sys, time, re
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()

SCRIPT_ID  = "1P1sZMNyE4wFCmidQy2kRHEa1gb1ozuODaknOgYlqJlZxn2cdoLhzQccC"
EDITOR_URL = f"https://script.google.com/d/{SCRIPT_ID}/edit"
ENV_FILE   = Path(__file__).parent / ".env"
SHOTS_DIR  = Path(__file__).parent / "debug_screenshots"
SHOTS_DIR.mkdir(exist_ok=True)

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

EDGE_PROFILE = r"C:\Users\i_ojasvi.singh\AppData\Local\Microsoft\Edge\User Data"

def shot(driver, name):
    p = str(SHOTS_DIR / f"deploy_{name}.png")
    try:
        driver.save_screenshot(p)
        print(f"  [screenshot: deploy_{name}.png]")
    except:
        pass

def save_url(url):
    content = ENV_FILE.read_text()
    if "APPS_SCRIPT_URL=" in content:
        content = re.sub(r"APPS_SCRIPT_URL=.*", f"APPS_SCRIPT_URL={url}", content)
    else:
        content += f"\nAPPS_SCRIPT_URL={url}\n"
    ENV_FILE.write_text(content)
    print(f"  Saved to .env: APPS_SCRIPT_URL={url}")

def find_and_click_deploy(driver):
    """Find Deploy button using JavaScript DOM traversal."""
    result = driver.execute_script("""
        // Search all elements for Deploy button
        var allElements = document.querySelectorAll('*');
        var candidates = [];
        for (var i = 0; i < allElements.length; i++) {
            var el = allElements[i];
            var text = el.innerText || el.textContent || '';
            text = text.trim();
            // Look for element with exactly 'Deploy' text that is clickable
            if (text === 'Deploy' && (
                el.tagName === 'BUTTON' ||
                el.tagName === 'SPAN' ||
                el.getAttribute('role') === 'button'
            )) {
                var rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) {
                    candidates.push({
                        tag: el.tagName,
                        role: el.getAttribute('role'),
                        id: el.id,
                        className: el.className ? el.className.substring(0, 80) : '',
                        x: rect.x, y: rect.y, w: rect.width, h: rect.height
                    });
                }
            }
        }
        return JSON.stringify(candidates);
    """)
    print(f"  Deploy candidates: {result}")
    return result

def click_deploy_js(driver):
    """Click the Deploy button using JavaScript."""
    clicked = driver.execute_script("""
        var allElements = document.querySelectorAll('*');
        for (var i = 0; i < allElements.length; i++) {
            var el = allElements[i];
            var text = (el.innerText || el.textContent || '').trim();
            if (text === 'Deploy') {
                var rect = el.getBoundingClientRect();
                if (rect.width > 30 && rect.height > 10) {
                    // Find the nearest clickable ancestor or self
                    var target = el;
                    while (target && target.tagName !== 'BUTTON' && target.getAttribute('role') !== 'button') {
                        target = target.parentElement;
                        if (!target || target === document.body) {
                            target = el;
                            break;
                        }
                    }
                    target.click();
                    return 'clicked:' + target.tagName + '/' + (target.getAttribute('role') || '');
                }
            }
        }
        return 'not_found';
    """)
    return clicked

def wait_for_text(driver, text, timeout=15):
    """Wait until page contains specific text."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if text in driver.page_source:
            return True
        time.sleep(0.5)
    return False

def main():
    print("=" * 55)
    print("  Apps Script — Deploy Step")
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

    print("\n[1/4] Launching Edge and opening editor...")
    driver = webdriver.Edge(options=opts)
    driver.maximize_window()

    try:
        driver.get(EDITOR_URL)
        time.sleep(6)
        shot(driver, "01_editor")

        # Wait for the editor to load (Monaco)
        print("  Waiting for editor to load...")
        deadline = time.time() + 60
        while time.time() < deadline:
            try:
                result = driver.execute_script("""
                    try { return monaco.editor.getModels().length > 0 ? 'ready' : 'wait'; }
                    catch(e) { return 'loading'; }
                """)
                if result == 'ready':
                    print("  Editor ready.")
                    break
            except:
                pass
            time.sleep(2)

        time.sleep(2)

        # ── Inspect Deploy button DOM ──────────────────────────────────
        print("\n[2/4] Inspecting Deploy button...")
        candidates_json = find_and_click_deploy(driver)
        shot(driver, "02_before_deploy")

        # ── Click Deploy ───────────────────────────────────────────────
        print("\n[3/4] Clicking Deploy button...")
        clicked = click_deploy_js(driver)
        print(f"  Click result: {clicked}")

        if clicked == 'not_found':
            # Try Selenium selectors as fallback
            print("  JS click failed. Trying Selenium selectors...")
            selectors = [
                (By.XPATH, "//div[contains(@class,'goog-button') and normalize-space(.)='Deploy']"),
                (By.XPATH, "//div[@role='button' and normalize-space(.)='Deploy']"),
                (By.CSS_SELECTOR, "button.btn-deploy, [data-action='deploy']"),
                (By.XPATH, "//span[text()='Deploy']/ancestor::div[@role='button'][1]"),
                (By.XPATH, "//span[text()='Deploy']/ancestor::button[1]"),
            ]
            for by, sel in selectors:
                try:
                    el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, sel)))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    time.sleep(0.3)
                    el.click()
                    clicked = f"selenium:{sel[:40]}"
                    print(f"  Clicked via Selenium: {sel[:60]}")
                    break
                except TimeoutException:
                    continue

        time.sleep(3)
        shot(driver, "03_after_deploy_click")

        # Check if a menu appeared
        if "New deployment" in driver.page_source or "Manage deployments" in driver.page_source:
            print("  Deploy menu opened!")
        else:
            print("  Deploy menu may not have opened. Check deploy_03_after_deploy_click.png")
            # Log page source snippet for debugging
            body_text = driver.execute_script("return document.body.innerText.substring(0, 500);")
            print(f"  Page text (first 500 chars): {body_text}")

        # ── New deployment ─────────────────────────────────────────────
        print("\n[4/4] Completing deployment dialog...")

        # Click "New deployment"
        new_deploy_clicked = driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                var t = (els[i].innerText || '').trim();
                if (t === 'New deployment') {
                    var r = els[i].getBoundingClientRect();
                    if (r.width > 0 && r.height > 0) {
                        els[i].click();
                        return 'clicked_new_deploy';
                    }
                }
            }
            return 'not_found';
        """)
        print(f"  New deployment click: {new_deploy_clicked}")
        time.sleep(3)
        shot(driver, "04_new_deploy_dialog")

        # Select "Web app" type
        driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                var t = (els[i].innerText || '').trim();
                if (t === 'Select type' || t === 'Web app') {
                    var r = els[i].getBoundingClientRect();
                    if (r.width > 0 && r.height > 0 && (
                        els[i].getAttribute('role') === 'button' ||
                        els[i].tagName === 'BUTTON' ||
                        els[i].tagName === 'SELECT'
                    )) {
                        els[i].click();
                        return 'clicked:' + t;
                    }
                }
            }
        """)
        time.sleep(1)

        driver.execute_script("""
            var els = document.querySelectorAll('*');
            for (var i = 0; i < els.length; i++) {
                var t = (els[i].innerText || '').trim();
                if (t === 'Web app') {
                    var r = els[i].getBoundingClientRect();
                    if (r.width > 0) { els[i].click(); break; }
                }
            }
        """)
        time.sleep(1)
        shot(driver, "05_type_selected")

        # Set access to Anyone
        driver.execute_script("""
            var selects = document.querySelectorAll('select');
            for (var i = 0; i < selects.length; i++) {
                var opts = selects[i].querySelectorAll('option');
                for (var j = 0; j < opts.length; j++) {
                    if (opts[j].text.indexOf('Anyone') >= 0) {
                        selects[i].value = opts[j].value;
                        selects[i].dispatchEvent(new Event('change'));
                        break;
                    }
                }
            }
            // Also try role=option elements
            var els = document.querySelectorAll('[role="option"], [role="listitem"]');
            for (var i = 0; i < els.length; i++) {
                var t = (els[i].innerText || '').trim();
                if (t.indexOf('Anyone') >= 0) {
                    els[i].click();
                    break;
                }
            }
        """)
        time.sleep(1)
        shot(driver, "06_access_set")

        # Click final Deploy button in dialog
        driver.execute_script("""
            // Find Deploy button in the dialog (not the toolbar one)
            var buttons = document.querySelectorAll('button, [role="button"]');
            for (var i = buttons.length - 1; i >= 0; i--) {
                var t = (buttons[i].innerText || '').trim();
                var r = buttons[i].getBoundingClientRect();
                if ((t === 'Deploy' || t === 'Deploy ') && r.width > 0 && r.y > 100) {
                    buttons[i].click();
                    return 'deploy_clicked';
                }
            }
        """)
        time.sleep(8)
        shot(driver, "07_deployed")

        # Grab URL
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
            shot(driver, "08_no_url")
            print("\nDeployment dialog opened but URL not captured.")
            print("Check deploy_07_deployed.png in debug_screenshots/")
            print("\nPlease copy the Web App URL from the browser and paste it:")
            url = input("Web App URL: ").strip()
            if url:
                save_url(url)

    except KeyboardInterrupt:
        print("\nCancelled.")
    except Exception as ex:
        shot(driver, "error")
        print(f"\nError: {ex}")
        import traceback
        traceback.print_exc()
        print("Browser left open.")

if __name__ == "__main__":
    main()
