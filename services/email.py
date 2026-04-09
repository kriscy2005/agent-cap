"""Email sending via Gmail SMTP (app password)."""
import os
import smtplib
import time
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
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_USER, to_address, msg.as_string())
        print(f"[Email] Sent to {to_address} ({name})")
        return {"sent": True, "to": to_address}
    except Exception as e:
        print(f"[Email] Failed for {name}: {e}")
        return {"sent": False, "to": to_address, "error": str(e)}


def send_all_emails(recipients: list) -> dict:
    """recipients: [{name, email}]. Returns {sent: [], failed: []}"""
    sent, failed = [], []
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    except Exception as e:
        return {"sent": [], "failed": [{"name": r["name"], "email": r["email"], "error": str(e)} for r in recipients]}

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
            server.sendmail(GMAIL_USER, to_address, msg.as_string())
            print(f"[Email] Sent to {to_address} ({name})")
            sent.append({"name": name, "email": to_address})
        except smtplib.SMTPServerDisconnected:
            # Reconnect and retry once
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                server.sendmail(GMAIL_USER, to_address, msg.as_string())
                sent.append({"name": name, "email": to_address})
            except Exception as e:
                print(f"[Email] Failed for {name}: {e}")
                failed.append({"name": name, "email": to_address, "error": str(e)})
        except Exception as e:
            print(f"[Email] Failed for {name}: {e}")
            failed.append({"name": name, "email": to_address, "error": str(e)})
        time.sleep(1)  # 1s between emails to respect Gmail rate limits

    server.quit()
    return {"sent": sent, "failed": failed}
