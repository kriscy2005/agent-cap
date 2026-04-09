"""
Write bandwidth log rows via Google Apps Script Web App.
No service account needed — the script runs as the sheet owner.
Set APPS_SCRIPT_URL in .env after deploying apps_script.js.
"""
import os
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env", override=True)

APPS_SCRIPT_URL = os.getenv("APPS_SCRIPT_URL", "")


def log_bandwidth_submissions(entries: list):
    """POST entries to the Apps Script web app, which appends rows to Bandwidth_Log."""
    if not entries:
        return
    if not APPS_SCRIPT_URL:
        raise RuntimeError(
            "APPS_SCRIPT_URL not set in .env. "
            "Deploy apps_script.js as a Web App and paste the URL."
        )
    payload = {"entries": entries}
    resp = requests.post(APPS_SCRIPT_URL, json=payload, timeout=20)
    resp.raise_for_status()
    result = resp.json()
    if not result.get("success"):
        raise RuntimeError(f"Apps Script error: {result.get('error')}")
    print(f"[Sheets] Logged {result.get('rows')} rows via Apps Script.")
