"""Jira Cloud REST API v3 — comment posting."""
import os
import base64
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env", override=True)

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL", "https://innovaccer.atlassian.net")
JIRA_EMAIL = os.getenv("JIRA_EMAIL", "")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN", "")
QUARTER_LABEL = os.getenv("QUARTER_LABEL", "Q1'26")


def _auth_header() -> str:
    token = base64.b64encode(f"{JIRA_EMAIL}:{JIRA_API_TOKEN}".encode()).decode()
    return f"Basic {token}"


def post_comment(ticket_key: str, creator_name: str, bandwidth: float) -> bool:
    """Posts an ADF comment on a Jira ticket. Returns True on success."""
    url = f"{JIRA_BASE_URL}/rest/api/3/issue/{ticket_key}/comment"
    payload = {
        "body": {
            "type": "doc",
            "version": 1,
            "content": [{
                "type": "paragraph",
                "content": [{
                    "type": "text",
                    "text": f"Bandwidth allocated by {creator_name}: {bandwidth}% \u2014 {QUARTER_LABEL}"
                }]
            }]
        }
    }
    try:
        resp = requests.post(
            url, json=payload,
            headers={
                "Authorization": _auth_header(),
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            timeout=15,
        )
        if resp.status_code in (200, 201):
            return True
        print(f"[Jira] {ticket_key}: HTTP {resp.status_code} — {resp.text[:200]}")
        return False
    except Exception as e:
        print(f"[Jira] {ticket_key}: {e}")
        return False


def post_comments_for_submission(allocations: list, creator_name: str) -> dict:
    """
    allocations: [{key, bandwidth}]
    Returns {success: [key, ...], failed: [key, ...]}
    """
    success, failed = [], []
    for item in allocations:
        ok = post_comment(item["key"], creator_name, item["bandwidth"])
        (success if ok else failed).append(item["key"])
    return {"success": success, "failed": failed}
