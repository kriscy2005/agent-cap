"""Email sending via Gmail SMTP with IPv4 forced (fixes Railway IPv6 routing)."""
import os
import socket
import smtplib
import json
import time
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env", override=True)

GMAIL_USER = os.getenv("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
BASE_URL = os.getenv("BASE_URL", "http://localhost:3000")
QUARTER_LABEL = os.getenv("QUARTER_LABEL", "Q1'26")
TEST_MODE = os.getenv("TEST_MODE", "true").lower() == "true"
TEST_EMAIL = os.getenv("TEST_EMAIL", "")

SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
SA_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "service-account.json")


# ── Force IPv4 to avoid Railway IPv6 routing issues ──────────────────────────

_original_getaddrinfo = socket.getaddrinfo

def _ipv4_getaddrinfo(*args, **kwargs):
    results = _original_getaddrinfo(*args, **kwargs)
    ipv4 = [r for r in results if r[0] == socket.AF_INET]
    return ipv4 if ipv4 else results

socket.getaddrinfo = _ipv4_getaddrinfo


def _build_html(first_name: str, form_url: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bandwidth Allocation — {QUARTER_LABEL}</title>
  <style>
    body {{ margin:0; padding:0; background:#f5f7fa; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif; }}
    .wrapper {{ max-width:600px; margin:40px auto; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.08); }}
    .header {{ background:#0066CC; padding:32px 40px; }}
    .header h1 {{ margin:0; color:#fff; font-size:22px; font-weight:600; }}
    .header p {{ margin:6px 0 0; color:rgba(255,255,255,0.8); font-size:14px; }}
    .body {{ padding:36px 40px; }}
    .body p {{ margin:0 0 16px; color:#333; font-size:15px; line-height:1.6; }}
    .cta {{ text-align:center; margin:32px 0; }}
    .cta a {{ display:inline-block; background:#0066CC; color:#fff; text-decoration:none; padding:14px 36px; border-radius:6px; font-size:15px; font-weight:600; }}
    .note {{ background:#f0f6ff; border-left:4px solid #0066CC; border-radius:4px; padding:12px 16px; margin:24px 0 0; }}
    .note p {{ margin:0; font-size:13px; color:#444; }}
    .footer {{ border-top:1px solid #eee; padding:20px 40px; }}
    .footer p {{ margin:0; font-size:12px; color:#999; line-height:1.5; }}
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="header">
      <h1>{QUARTER_LABEL} Bandwidth Allocation</h1>
      <p>Action Required — Product Operations</p>
    </div>
    <div class="body">
      <p>Hi {first_name},</p>
      <p>It&rsquo;s time to allocate your bandwidth % for <strong>{QUARTER_LABEL} capitalised tickets</strong>. This helps us accurately track engineering effort for financial reporting.</p>
      <p>Please open the form below, review your tickets (as Creator, Assignee, or PM/Director), and enter your estimated bandwidth % for each.</p>
      <div class="cta"><a href="{form_url}">Open Bandwidth Form &rarr;</a></div>
      <div class="note">
        <p><strong>Note:</strong> You can resubmit anytime — each submission is logged separately with a timestamp.</p>
      </div>
    </div>
    <div class="footer">
      <p>Sent by the Bandwidth Allocation Bot on behalf of Product Operations.</p>
    </div>
  </div>
</body>
</html>"""


# ── Gmail API sender (HTTP — fallback when SMTP is fully blocked) ────────────

def _get_gmail_service():
    """Build Gmail API client using service account with domain-wide delegation."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/gmail.send"]
    if SA_JSON:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(SA_JSON), scopes=scopes
        )
    else:
        creds = service_account.Credentials.from_service_account_file(
            SA_PATH, scopes=scopes
        )
    creds = creds.with_subject(GMAIL_USER)
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def _send_via_gmail_api(msg, gmail_service):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    gmail_service.users().messages().send(userId="me", body={"raw": raw}).execute()


# ── SMTP sender ──────────────────────────────────────────────────────────────

def _connect_smtp():
    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=15)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    return server


# ── Main send functions ──────────────────────────────────────────────────────

def send_allocation_email(name: str, email: str) -> dict:
    """Sends a personalised email. Returns {sent: bool, to: str, error?: str}"""
    first_name = name.split()[0]
    form_url = f"{BASE_URL}/form?creator={quote(name)}"
    to_address = TEST_EMAIL if TEST_MODE else email

    msg = MIMEMultipart("alternative")
    msg["Subject"] = (
        f"Action Required: {QUARTER_LABEL} Bandwidth Allocation"
        + (f" [TEST — actual: {name} <{email}>]" if TEST_MODE and to_address != email else "")
    )
    msg["From"] = f"Product Operations <{GMAIL_USER}>"
    msg["To"] = to_address
    msg.attach(MIMEText(_build_html(first_name, form_url), "html"))

    try:
        server = _connect_smtp()
        server.sendmail(GMAIL_USER, to_address, msg.as_string())
        server.quit()
        print(f"[Email] Sent to {to_address} ({name})")
        return {"sent": True, "to": to_address}
    except Exception as e:
        print(f"[Email] Failed for {name}: {e}")
        return {"sent": False, "to": to_address, "error": str(e)}


def send_all_emails(recipients: list) -> dict:
    """recipients: [{name, email}]. Returns {sent: [], failed: []}"""
    sent, failed = [], []

    # Try SMTP first
    smtp_server = None
    use_api = False
    try:
        smtp_server = _connect_smtp()
    except Exception as smtp_err:
        print(f"[Email] SMTP failed ({smtp_err}), trying Gmail API...")
        try:
            gmail_service = _get_gmail_service()
            use_api = True
        except Exception as api_err:
            err = f"SMTP: {smtp_err} | Gmail API: {api_err}"
            return {"sent": [], "failed": [{"name": r["name"], "email": r["email"], "error": err} for r in recipients]}

    for r in recipients:
        name = r["name"]
        email = r["email"]
        first_name = name.split()[0]
        form_url = f"{BASE_URL}/form?creator={quote(name)}"
        to_address = TEST_EMAIL if TEST_MODE else email

        msg = MIMEMultipart("alternative")
        msg["Subject"] = (
            f"Action Required: {QUARTER_LABEL} Bandwidth Allocation"
            + (f" [TEST — actual: {name} <{email}>]" if TEST_MODE and to_address != email else "")
        )
        msg["From"] = f"Product Operations <{GMAIL_USER}>"
        msg["To"] = to_address
        msg.attach(MIMEText(_build_html(first_name, form_url), "html"))

        try:
            if use_api:
                _send_via_gmail_api(msg, gmail_service)
            else:
                smtp_server.sendmail(GMAIL_USER, to_address, msg.as_string())
            print(f"[Email] Sent to {to_address} ({name})")
            sent.append({"name": name, "email": to_address})
        except smtplib.SMTPServerDisconnected:
            try:
                smtp_server = _connect_smtp()
                smtp_server.sendmail(GMAIL_USER, to_address, msg.as_string())
                sent.append({"name": name, "email": to_address})
            except Exception as e:
                print(f"[Email] Failed for {name}: {e}")
                failed.append({"name": name, "email": to_address, "error": str(e)})
        except Exception as e:
            print(f"[Email] Failed for {name}: {e}")
            failed.append({"name": name, "email": to_address, "error": str(e)})
        time.sleep(1)

    if smtp_server:
        try:
            smtp_server.quit()
        except Exception:
            pass
    return {"sent": sent, "failed": failed}
