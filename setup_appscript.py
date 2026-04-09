"""
One-shot setup script:
  1. OAuth2 browser login (Google account that owns the sheet)
  2. Creates an Apps Script project bound to the sheet
  3. Uploads the bandwidth log code
  4. Deploys it as a public Web App
  5. Writes APPS_SCRIPT_URL to .env automatically

Requirements:
  - Place client_secrets.json (OAuth2 Desktop credentials) in this folder
  - Run: python setup_appscript.py

How to get client_secrets.json (30 seconds):
  1. https://console.cloud.google.com  →  select your project
  2. APIs & Services → Enable: "Apps Script API" + "Google Sheets API"
  3. APIs & Services → Credentials → + Create Credentials → OAuth 2.0 Client ID
  4. Application type: Desktop app  →  Name: bandwidth-bot  →  Create
  5. Download JSON  →  rename to client_secrets.json  →  drop in this folder
"""

import os
import sys
import json
import time
import re
import webbrowser
from pathlib import Path
from dotenv import load_dotenv, set_key

load_dotenv()

import requests
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

SHEET_ID       = os.getenv("SHEET_ID", "19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY")
CLIENT_SECRETS = Path("client_secrets.json")
TOKEN_FILE     = Path("token.pkl")
ENV_FILE       = Path(".env")
SCOPES         = [
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

APPS_SCRIPT_CODE = r'''
const SHEET_ID = "''' + SHEET_ID + r'''";
const LOG_TAB  = "Bandwidth_Log";
const HEADERS  = ["Ticket Key","Summary","Brand","Person Name","Designation","Role","Bandwidth %","Submitted At"];

function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const entries = data.entries || [];
    if (!entries.length) return _json({ error: "No entries" });
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(LOG_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_TAB);
      sheet.appendRow(HEADERS);
      sheet.getRange(1,1,1,HEADERS.length).setFontWeight("bold")
           .setBackground("#0066CC").setFontColor("#ffffff");
    }
    const now  = new Date().toISOString();
    const rows = entries.map(r => [
      r.key, r.summary||"", r.brand||"",
      r.personName, r.designation||"", r.role||"Other",
      r.bandwidth, now
    ]);
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, HEADERS.length).setValues(rows);
    return _json({ success: true, rows: rows.length });
  } catch(err) {
    return _json({ error: err.message });
  }
}

function doGet() {
  return _json({ status: "ok", tab: LOG_TAB });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
'''


def get_credentials():
    """Load cached token or run browser OAuth2 flow."""
    creds = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
        return creds

    if not CLIENT_SECRETS.exists():
        print("\n❌  client_secrets.json not found.")
        print("   Follow the instructions at the top of this file to create it.")
        sys.exit(1)

    flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRETS), SCOPES)
    print("\n🌐  Opening browser for Google authorisation…")
    creds = flow.run_local_server(port=0, open_browser=True)
    with open(TOKEN_FILE, "wb") as f:
        pickle.dump(creds, f)
    return creds


def api(creds, method, url, **kwargs):
    headers = {"Authorization": f"Bearer {creds.token}", "Content-Type": "application/json"}
    resp = getattr(requests, method)(url, headers=headers, **kwargs)
    if not resp.ok:
        print(f"  ⚠  {method.upper()} {url.split('/')[-1]}: {resp.status_code} {resp.text[:200]}")
    return resp


def main():
    print("=" * 55)
    print("  Bandwidth Bot — Apps Script Setup")
    print("=" * 55)

    # ── 1. Auth ────────────────────────────────────────────────
    print("\n[1/5] Authenticating with Google…")
    creds = get_credentials()
    print("      ✅  Authenticated")

    # ── 2. Create Apps Script project ─────────────────────────
    print("\n[2/5] Creating Apps Script project…")
    create_resp = api(creds, "post",
        "https://script.googleapis.com/v1/projects",
        json={
            "title": "Bandwidth Bot — Log Writer",
            "parentId": SHEET_ID,
        }
    )
    if not create_resp.ok:
        print("  Failed to create project. Make sure 'Apps Script API' is enabled in Cloud Console.")
        sys.exit(1)
    script_id = create_resp.json()["scriptId"]
    print(f"      ✅  Script ID: {script_id}")

    # ── 3. Upload code ─────────────────────────────────────────
    print("\n[3/5] Uploading code…")
    content_resp = api(creds, "put",
        f"https://script.googleapis.com/v1/projects/{script_id}/content",
        json={
            "files": [
                {
                    "name": "Code",
                    "type": "SERVER_JS",
                    "source": APPS_SCRIPT_CODE,
                },
                {
                    "name": "appsscript",
                    "type": "JSON",
                    "source": json.dumps({
                        "timeZone": "Asia/Kolkata",
                        "exceptionLogging": "STACKDRIVER",
                        "runtimeVersion": "V8",
                        "webapp": {
                            "access": "ANYONE_ANONYMOUS",
                            "executeAs": "USER_DEPLOYING",
                        },
                    }),
                },
            ]
        }
    )
    if not content_resp.ok:
        sys.exit(1)
    print("      ✅  Code uploaded")

    # ── 4. Deploy as Web App ───────────────────────────────────
    print("\n[4/5] Deploying as Web App (Anyone can access)…")
    deploy_resp = api(creds, "post",
        f"https://script.googleapis.com/v1/projects/{script_id}/deployments",
        json={
            "versionNumber": 1,
            "manifestFileName": "appsscript",
            "description": "Bandwidth Bot log writer",
        }
    )
    if not deploy_resp.ok:
        # Try creating a version first
        version_resp = api(creds, "post",
            f"https://script.googleapis.com/v1/projects/{script_id}/versions",
            json={"description": "v1"}
        )
        version_number = version_resp.json().get("versionNumber", 1) if version_resp.ok else 1
        deploy_resp = api(creds, "post",
            f"https://script.googleapis.com/v1/projects/{script_id}/deployments",
            json={
                "versionNumber": version_number,
                "manifestFileName": "appsscript",
                "description": "v1",
            }
        )

    if not deploy_resp.ok:
        print("\n  Could not auto-deploy. Manual fallback:")
        print(f"  1. Open: https://script.google.com/d/{script_id}/edit")
        print("  2. Deploy → New deployment → Web App → Anyone → Deploy")
        print("  3. Copy the URL and add it to .env as APPS_SCRIPT_URL=<url>")
        sys.exit(1)

    deploy_data = deploy_resp.json()
    deployment_id = deploy_data.get("deploymentId") or deploy_data.get("deploymentConfig", {}).get("deploymentId")
    web_app_url = f"https://script.google.com/macros/s/{deployment_id}/exec"
    print(f"      ✅  Deployed: {web_app_url}")

    # ── 5. Save URL to .env ────────────────────────────────────
    print("\n[5/5] Saving APPS_SCRIPT_URL to .env…")
    env_path = str(ENV_FILE.resolve())
    env_content = ENV_FILE.read_text()
    if "APPS_SCRIPT_URL=" in env_content:
        new_content = re.sub(r"APPS_SCRIPT_URL=.*", f"APPS_SCRIPT_URL={web_app_url}", env_content)
        ENV_FILE.write_text(new_content)
    else:
        with open(ENV_FILE, "a") as f:
            f.write(f"\nAPPS_SCRIPT_URL={web_app_url}\n")
    print(f"      ✅  Saved to .env")

    # ── Done ───────────────────────────────────────────────────
    print("\n" + "=" * 55)
    print("  ✅  Setup complete!")
    print("=" * 55)
    print(f"\n  Web App URL : {web_app_url}")
    print(f"  Script URL  : https://script.google.com/d/{script_id}/edit")
    print("\n  Restart the server: the bot will now write to Bandwidth_Log")
    print("  on every form submission.\n")


if __name__ == "__main__":
    main()
