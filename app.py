"""Flask server — routes: /health, /form, /submit, /admin"""
import os
import json
import re
from pathlib import Path
from flask import Flask, request, jsonify, Response, redirect
from dotenv import load_dotenv

_ENV_FILE = Path(__file__).parent / ".env"
load_dotenv(_ENV_FILE, override=True)

from services import jira

# Use public CSV reader when no service account exists; fall back to auth'd client for writes
import os as _os
_SA = _os.path.abspath(_os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "./service-account.json"))
if _os.path.exists(_SA):
    from services import sheets as _sheet_reader
else:
    from services import sheets_csv as _sheet_reader  # no-auth public CSV reads

# For writes: prefer Apps Script webhook (no SA needed), fall back to sheets.py (needs SA)
if os.getenv("APPS_SCRIPT_URL"):
    from services import sheets_appscript as _sheet_writer
elif _os.path.exists(_SA):
    from services import sheets as _sheet_writer
else:
    _sheet_writer = None  # writes disabled until configured

# ── Mock data for ?mock=true testing (no credentials needed) ──────────────────

MOCK_FORM_DATA = {
    "personName": "Ojasvi Singh",
    "designation": "APM",
    "quarterLabel": "Q1'26",
    "brands": {
        "Care Management": {
            "asCreator": [
                {"key": "IPD-1001", "summary": "Build patient risk stratification engine for care gaps"},
                {"key": "IPD-1002", "summary": "Implement care plan auto-generation from clinical data"},
            ],
            "asAssignee": [
                {"key": "IPD-1003", "summary": "Integrate HL7 FHIR R4 patient data feed into care platform"},
            ],
            "other": [
                {"key": "IPD-1004", "summary": "Dashboard redesign for care coordinator workflow"},
                {"key": "IPD-1005", "summary": "SDoH screening tool integration with care plans"},
            ],
        },
        "Population Health": {
            "asCreator": [
                {"key": "IPD-2001", "summary": "Chronic disease predictive model — diabetes cohort v2"},
            ],
            "asAssignee": [],
            "other": [
                {"key": "IPD-2002", "summary": "Population segmentation by risk score and attribution"},
                {"key": "IPD-2003", "summary": "Quality measure reporting for HEDIS 2026 compliance"},
            ],
        },
    },
}

PORT = int(os.getenv("PORT", 3000))
app = Flask(__name__, static_folder="public")

FORM_TEMPLATE = Path(__file__).parent / "public" / "form.html"


# ── Health ────────────────────────────────────────────────────────────────────



@app.get("/")
def index():
    return redirect("/admin")


@app.get("/health")
def health():
    from datetime import datetime, timezone
    return jsonify({"status": "ok", "timestamp": datetime.now(timezone.utc).isoformat()})


# ── Form ──────────────────────────────────────────────────────────────────────

@app.get("/form")
def form():
    creator_name = (request.args.get("creator") or "").strip()
    use_mock = request.args.get("mock", "").lower() in ("1", "true", "yes")

    if not creator_name and not use_mock:
        return Response("<h2>Missing ?creator= parameter</h2>", status=400, mimetype="text/html")

    if use_mock and not creator_name:
        # Pure mock: use hardcoded sample data
        form_data = {**MOCK_FORM_DATA, "isMock": True}
    else:
        # Real data from sheet — clear caches so we always get fresh data
        _sheet_reader.clear_caches()
        try:
            rows = _sheet_reader.get_capitalised_rows()
            people_map = _sheet_reader.build_people_map()
            if use_mock:
                # mock=true but creator name given: use real sheet data, skip writes on submit
                form_data = _sheet_reader.get_form_data_for_person(creator_name, rows, people_map)
                if form_data:
                    form_data = {**form_data, "isMock": True}
            else:
                form_data = _sheet_reader.get_form_data_for_person(creator_name, rows, people_map)
        except Exception as e:
            return Response(f"<h2>Server error</h2><pre>{e}</pre>", status=500, mimetype="text/html")

        if form_data is None:
            return Response(
                f'<h2>Person not found: "{creator_name}"</h2>'
                f"<p>This person was not found in the recipients list or the sheet for {os.getenv('QUARTER_LABEL','Q1&#39;26')}.</p>",
                status=404, mimetype="text/html",
            )

    template = FORM_TEMPLATE.read_text(encoding="utf-8")

    # Safely inject FORM_DATA — escape </script> inside JSON string
    safe_json = json.dumps(form_data).replace("</script>", "<\\/script>")
    injected = f"window.FORM_DATA = JSON.parse({json.dumps(safe_json)});"
    html = template.replace("/* __FORM_DATA_PLACEHOLDER__ */", injected)

    return Response(html, status=200, mimetype="text/html; charset=utf-8")


# ── Submit ─────────────────────────────────────────────────────────────────────

@app.post("/submit")
def submit():
    body = request.get_json(force=True, silent=True) or {}
    person_name = (body.get("personName") or "").strip()
    designation = (body.get("designation") or "").strip()
    entries = body.get("entries") or []

    if not person_name:
        return jsonify({"error": "Missing personName"}), 400
    if not entries:
        return jsonify({"error": "No entries provided"}), 400

    # Validate entries
    valid = []
    for e in entries:
        try:
            bw = float(e.get("bandwidth", ""))
        except (TypeError, ValueError):
            continue
        if e.get("key") and 0 <= bw <= 100:
            valid.append({**e, "bandwidth": bw})

    if not valid:
        return jsonify({"error": "No valid entries (bandwidth must be 0–100)"}), 400

    total_bw = sum(e["bandwidth"] for e in valid)
    if total_bw > 100:
        return jsonify({"error": f"Total bandwidth is {total_bw:.0f}% — must not exceed 100%. Please reduce your allocations."}), 400

    is_mock = body.get("mock") is True

    log_entries = [
        {
            "key": e["key"],
            "summary": e.get("summary", ""),
            "brand": e.get("brand", ""),
            "personName": person_name,
            "designation": designation,
            "role": e.get("role", "Other"),
            "bandwidth": e["bandwidth"],
        }
        for e in valid
    ]

    if is_mock:
        print(f"[/submit][MOCK] {person_name} submitted {len(log_entries)} entries:")
        for e in log_entries:
            print(f"  {e['key']} — {e['bandwidth']}% ({e['role']})")
        return jsonify({
            "saved": len(log_entries),
            "sheetError": None,
            "jiraSuccess": [e["key"] for e in log_entries],
            "jiraFailed": [],
            "mock": True,
        })

    # 1. Log to sheet
    sheet_error = None
    if _sheet_writer is None:
        sheet_error = "No sheet writer configured. Set APPS_SCRIPT_URL or add service-account.json."
        print(f"[/submit] {sheet_error}")
    else:
        try:
            _sheet_writer.log_bandwidth_submissions(log_entries)
        except Exception as ex:
            sheet_error = str(ex)
            print(f"[/submit] Sheet error: {ex}")

    # 2. Jira comments
    allocations = [{"key": e["key"], "bandwidth": e["bandwidth"]} for e in valid]
    jira_result = jira.post_comments_for_submission(allocations, person_name)

    return jsonify({
        "saved": 0 if sheet_error else len(log_entries),
        "sheetError": sheet_error,
        "jiraSuccess": jira_result["success"],
        "jiraFailed": jira_result["failed"],
    })


# ── Admin ─────────────────────────────────────────────────────────────────────

ADMIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bandwidth Bot — Admin</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;background:#f5f7fa;min-height:100vh;padding:40px 20px}
    .card{max-width:680px;margin:0 auto;background:#fff;border-radius:10px;box-shadow:0 2px 12px rgba(0,0,0,0.08);overflow:hidden}
    .header{background:#0066CC;padding:28px 36px}
    .header h1{color:#fff;font-size:20px;font-weight:600}
    .header p{color:rgba(255,255,255,0.75);font-size:13px;margin-top:4px}
    .body{padding:32px 36px}
    .field{margin-bottom:20px}
    label{display:block;font-size:13px;font-weight:600;color:#444;margin-bottom:6px}
    input[type=text]{width:100%;padding:9px 12px;border:1px solid #ddd;border-radius:6px;font-size:14px;color:#333}
    .toggle-row{display:flex;align-items:center;gap:10px;margin-bottom:20px}
    .toggle-row label{margin:0;font-size:14px;color:#333;font-weight:400}
    input[type=checkbox]{width:16px;height:16px;cursor:pointer}
    .recipients{background:#f8f9fb;border:1px solid #e5e7eb;border-radius:6px;padding:12px 16px;max-height:220px;overflow-y:auto;margin-bottom:24px;font-size:13px;color:#555;line-height:1.7}
    .recipients b{color:#222}
    .btn{background:#0066CC;color:#fff;border:none;padding:12px 28px;border-radius:6px;font-size:15px;font-weight:600;cursor:pointer;width:100%}
    .btn:hover{background:#0055aa}
    .btn:disabled{background:#aaa;cursor:not-allowed}
    .result{margin-top:24px;padding:16px;border-radius:6px;font-size:14px;line-height:1.7}
    .result.ok{background:#f0fdf4;border:1px solid #86efac;color:#166534}
    .result.err{background:#fef2f2;border:1px solid #fca5a5;color:#991b1b}
    .badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600;margin-left:6px}
    .badge.test{background:#fef3c7;color:#92400e}
    .badge.live{background:#dcfce7;color:#166534}
    .spinner{display:none;margin:0 auto 0 10px;width:16px;height:16px;border:2px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin .6s linear infinite;vertical-align:middle}
    @keyframes spin{to{transform:rotate(360deg)}}
  </style>
</head>
<body>
<div class="card">
  <div class="header">
    <h1>Bandwidth Bot — Admin</h1>
    <p>Send quarterly bandwidth allocation emails to all creators</p>
  </div>
  <div class="body">
    <div class="field">
      <label>Quarter Label</label>
      <input type="text" id="quarter" value="__QUARTER__" placeholder="e.g. Q1'26">
    </div>
    <div class="toggle-row">
      <input type="checkbox" id="testMode" __TEST_CHECKED__>
      <label for="testMode">Test mode — send all emails to <strong>__TEST_EMAIL__</strong> instead</label>
      <span class="badge __MODE_CLASS__">__MODE_LABEL__</span>
    </div>
    <div class="field">
      <label>Recipients (__COUNT__ people from sheet)</label>
      <div class="recipients">__RECIPIENTS_HTML__</div>
    </div>
    <button class="btn" id="sendBtn" onclick="sendEmails()">
      Send __COUNT__ Emails &nbsp;<span class="spinner" id="spinner"></span>
    </button>
    <div id="result"></div>
  </div>
</div>
<script>
async function sendEmails() {
  const btn = document.getElementById('sendBtn');
  const spinner = document.getElementById('spinner');
  const result = document.getElementById('result');
  const quarter = document.getElementById('quarter').value.trim();
  const testMode = document.getElementById('testMode').checked;

  btn.disabled = true;
  spinner.style.display = 'inline-block';
  result.innerHTML = '';

  try {
    const resp = await fetch('/admin/send-emails', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({quarter, testMode})
    });
    const data = await resp.json();
    if (data.error) {
      result.className = 'result err';
      result.innerHTML = '<b>Error:</b> ' + data.error;
    } else {
      result.className = 'result ok';
      result.innerHTML =
        '<b>Done!</b> Sent: ' + data.sent.length + ' &nbsp;|&nbsp; Failed: ' + data.failed.length + '<br><br>' +
        data.sent.map(r => '✓ ' + r.name + ' &rarr; ' + r.email).join('<br>') +
        (data.failed.length ? '<br><br><b>Failed:</b><br>' + data.failed.map(r => '✗ ' + r.name + ': ' + r.error).join('<br>') : '');
    }
  } catch(e) {
    result.className = 'result err';
    result.innerHTML = '<b>Network error:</b> ' + e.message;
  }
  btn.disabled = false;
  spinner.style.display = 'none';
}
</script>
</body>
</html>"""


@app.get("/admin")
def admin():
    load_dotenv(_ENV_FILE, override=True)
    _sheet_reader.clear_caches()
    try:
        recipients = _sheet_reader.get_recipients_from_tab()
    except Exception:
        recipients = []

    quarter = os.getenv("QUARTER_LABEL", "Q1'26")
    test_mode = os.getenv("TEST_MODE", "true").lower() == "true"
    test_email = os.getenv("TEST_EMAIL", "")

    PLACEHOLDER = "i_ojasvi.singh@innovaccer.com"
    recipients_html = "".join(
        f"<div><b>{r['name']}</b> &lt;{r['email']}&gt;"
        + (" <span style='color:#f59e0b;font-size:11px;font-weight:600;'>⚠ placeholder email</span>" if r['email'] == PLACEHOLDER else "")
        + "</div>"
        for r in recipients
    ) or "<i>No recipients found — check sheet data.</i>"

    html = (ADMIN_HTML
        .replace("__QUARTER__", quarter)
        .replace("__TEST_CHECKED__", "checked" if test_mode else "")
        .replace("__TEST_EMAIL__", test_email)
        .replace("__MODE_CLASS__", "test" if test_mode else "live")
        .replace("__MODE_LABEL__", "TEST" if test_mode else "LIVE")
        .replace("__COUNT__", str(len(recipients)))
        .replace("__RECIPIENTS_HTML__", recipients_html))

    return Response(html, mimetype="text/html; charset=utf-8")


@app.post("/admin/send-emails")
def admin_send_emails():
    from services.email import send_all_emails
    from services import email as email_svc
    import importlib

    from dotenv import load_dotenv as _load_dotenv
    _load_dotenv(_ENV_FILE, override=True)
    _sheet_reader.clear_caches()

    body = request.get_json(force=True, silent=True) or {}
    quarter = (body.get("quarter") or os.getenv("QUARTER_LABEL", "Q1'26")).strip()
    test_mode = body.get("testMode", True)

    # Temporarily override env values for this send
    os.environ["QUARTER_LABEL"] = quarter
    os.environ["TEST_MODE"] = "true" if test_mode else "false"
    importlib.reload(email_svc)

    gmail_user = os.getenv("GMAIL_USER", "")
    gmail_pass = os.getenv("GMAIL_APP_PASSWORD", "")
    if not gmail_user or not gmail_pass:
        return jsonify({"error": "Gmail credentials not configured. Set GMAIL_USER and GMAIL_APP_PASSWORD in .env"}), 400

    try:
        recipients = _sheet_reader.get_recipients_from_tab()
    except Exception as e:
        return jsonify({"error": f"Could not load recipients from sheet: {e}"}), 500

    if not recipients:
        return jsonify({"error": "No recipients found in sheet."}), 400

    result = email_svc.send_all_emails(recipients)
    print(f"[/admin/send-emails] Sent: {len(result['sent'])}, Failed: {len(result['failed'])}")
    return jsonify(result)


# ── Auto-detect LAN IP and keep .env in sync ──────────────────────────────────

def _sync_base_url():
    """Detect current LAN IP and update BASE_URL in .env if it has changed."""
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        return
    new_url = f"http://{ip}:{PORT}"
    old_url = os.getenv("BASE_URL", "")
    if new_url != old_url:
        text = _ENV_FILE.read_text(encoding="utf-8")
        text = re.sub(r"BASE_URL=http://[^\n]+", f"BASE_URL={new_url}", text)
        _ENV_FILE.write_text(text, encoding="utf-8")
        load_dotenv(_ENV_FILE, override=True)
        print(f"[Bandwidth Bot] IP updated: {old_url} → {new_url}")


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    _sync_base_url()
    current_url = os.getenv("BASE_URL", f"http://localhost:{PORT}")
    print(f"\n[Bandwidth Bot] Starting on {current_url}")
    print(f"  Admin     : {current_url}/admin")
    print(f"  TEST_MODE : {os.getenv('TEST_MODE')}")
    print(f"  SHEET_ID  : {os.getenv('SHEET_ID')}\n")
    app.run(host="0.0.0.0", port=PORT, debug=True, use_reloader=False)
