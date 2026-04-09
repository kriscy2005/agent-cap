"""Google Sheets service — all read/write logic."""
import os
import re
import json
from functools import lru_cache
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

SHEET_ID = os.getenv("SHEET_ID")
SA_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
SA_PATH = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "./service-account.json")
QUARTER_LABEL = os.getenv("QUARTER_LABEL", "Q1'26")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Module-level caches
_service = None
_brand_map = None
_people_map = None
_email_cache = None


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_service():
    global _service
    if _service:
        return _service
    # Try env var first (Railway), then file (local)
    if SA_JSON:
        creds = service_account.Credentials.from_service_account_info(json.loads(SA_JSON), scopes=SCOPES)
    else:
        sa_path = os.path.abspath(SA_PATH)
        creds = service_account.Credentials.from_service_account_file(sa_path, scopes=SCOPES)
    _service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return _service


def _get_values(range_: str):
    """Fetch a range from the sheet, returning list-of-lists (empty cells = '')."""
    svc = get_service()
    res = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=range_
    ).execute()
    return res.get("values", [])


# ── Helpers ───────────────────────────────────────────────────────────────────

KEY_RE = re.compile(r"([A-Z]+-\d+)")


def extract_key(raw: str) -> str | None:
    m = KEY_RE.search(str(raw or ""))
    return m.group(1) if m else None


def norm(s: str) -> str:
    return str(s or "").strip().lower()


def is_valid_email(email: str) -> bool:
    s = str(email or "").strip()
    return "@" in s and "#N/A" not in s and s != ""


def safe_get(row: list, idx: int, default="") -> str:
    try:
        return str(row[idx]).strip() if idx < len(row) else default
    except Exception:
        return default


# ── buildBrandMap ─────────────────────────────────────────────────────────────

def build_brand_map() -> dict:
    """Returns {product_team_lower: brand_name}"""
    global _brand_map
    if _brand_map is not None:
        return _brand_map
    rows = _get_values("Brand<>Product!A:B")
    result = {}
    for row in rows[1:]:  # skip header
        if len(row) < 2:
            continue
        product_team, brand = row[0].strip(), row[1].strip()
        if product_team and brand:
            result[norm(product_team)] = brand
    _brand_map = result
    return _brand_map


# ── buildEmailCache ───────────────────────────────────────────────────────────

def build_email_cache() -> dict:
    """Returns {name_lower: email} from Emails tab (Col B = name, Col D = email)."""
    global _email_cache
    if _email_cache is not None:
        return _email_cache
    rows = _get_values("Emails!A:D")
    result = {}
    for row in rows[1:]:
        name = safe_get(row, 1)
        email = safe_get(row, 3)
        if name and is_valid_email(email):
            result[norm(name)] = email
    _email_cache = result
    return _email_cache


# ── buildPeopleMap ────────────────────────────────────────────────────────────

def build_people_map() -> dict:
    """
    Returns {name_lower: {name, designation, brands: [], email}}
    Columns: EMP ID|Name|Region|Designation|Solution 1|Solution 2|Solution 3|Email
                0    1    2       3           4          5          6          7
    """
    global _people_map
    if _people_map is not None:
        return _people_map
    rows = _get_values("People<>Soln!A:H")
    result = {}
    for row in rows[1:]:
        name = safe_get(row, 1)
        if not name:
            continue
        designation = safe_get(row, 3)
        brands = [b for b in [safe_get(row, 4), safe_get(row, 5), safe_get(row, 6)] if b]
        email = safe_get(row, 7)
        result[norm(name)] = {"name": name, "designation": designation, "brands": brands, "email": email}
    _people_map = result
    return _people_map


# ── getEmailForPerson ─────────────────────────────────────────────────────────

def get_email_for_person(name: str) -> str | None:
    email_cache = build_email_cache()
    people_map = build_people_map()
    # Primary: Emails tab
    email = email_cache.get(norm(name))
    if email:
        return email
    # Fallback: People<>Soln
    person = people_map.get(norm(name))
    if person and is_valid_email(person["email"]):
        return person["email"]
    return None


# ── getCapitalisedRows ────────────────────────────────────────────────────────

def get_capitalised_rows() -> list:
    """
    Reads the quarter tab and returns enriched dicts for Capitalisation=Yes rows.
    Columns (0-indexed):
      0  Key         7  Creator        12 Assignee
      3  Summary     9  Capitalisation 13 Assignee Email
      10 Product Team 16 Brand         17 PM + Director
      14 Creator Email
    """
    brand_map = build_brand_map()
    rows = _get_values(f"{QUARTER_LABEL}!A:R")
    if not rows:
        return []

    # Build column index from header row (resilient to reordering)
    header = [norm(h) for h in rows[0]]

    def col(name):
        try:
            return header.index(name)
        except ValueError:
            return -1

    i_key = col("key")
    i_summary = col("summary")
    i_cap = col("capitalisation")
    i_product_team = col("product team")
    i_creator = col("creator")
    i_assignee = col("assignee")
    i_assignee_email = col("assignee email")
    i_creator_email = col("creator email")
    i_pm_director = col("pm + director")
    i_brand = col("brand")

    results = []
    for row in rows[1:]:
        cap = norm(safe_get(row, i_cap))
        if cap != "yes":
            continue
        raw_key = safe_get(row, i_key)
        key = extract_key(raw_key)
        if not key:
            continue

        product_team = safe_get(row, i_product_team)
        brand = safe_get(row, i_brand) if i_brand >= 0 else ""
        if not brand:
            brand = brand_map.get(norm(product_team), product_team)

        pm_raw = safe_get(row, i_pm_director) if i_pm_director >= 0 else ""
        pm_directors = [p.strip() for p in pm_raw.split(",") if p.strip()]

        results.append({
            "key": key,
            "summary": safe_get(row, i_summary),
            "product_team": product_team,
            "brand": brand,
            "creator": safe_get(row, i_creator),
            "assignee": safe_get(row, i_assignee),
            "assignee_email": safe_get(row, i_assignee_email),
            "creator_email": safe_get(row, i_creator_email),
            "pm_directors": pm_directors,
        })
    return results


# ── getAllRecipients ───────────────────────────────────────────────────────────

def get_all_recipients(rows: list, people_map: dict) -> list:
    """Returns [{name, email, designation}] — allowlisted by People<>Soln."""
    seen = set()
    recipients = []

    def add(name):
        if not name:
            return
        k = norm(name)
        if k in seen:
            return
        seen.add(k)
        person = people_map.get(k)
        if not person:
            return
        email = get_email_for_person(name)
        if not is_valid_email(email):
            return
        recipients.append({"name": person["name"], "email": email, "designation": person["designation"]})

    for row in rows:
        add(row["creator"])
        add(row["assignee"])
        for pm in row["pm_directors"]:
            add(pm)

    return recipients


# ── getFormDataForPerson ──────────────────────────────────────────────────────

def get_form_data_for_person(name: str, rows: list, people_map: dict) -> dict | None:
    """
    Returns structured form data dict:
    {
      personName, designation, quarterLabel,
      brands: {brand: {asCreator: [...], asAssignee: [...], other: [...]}}
    }
    Priority: Creator > Assignee > Other.
    """
    k = norm(name)
    person = people_map.get(k)
    if not person:
        return None

    brand_data = {b: {"asCreator": [], "asAssignee": [], "other": []} for b in person["brands"]}

    for row in rows:
        brand = row["brand"]
        if brand not in brand_data:
            continue

        is_creator = norm(row["creator"]) == k
        is_assignee = norm(row["assignee"]) == k
        ticket = {"key": row["key"], "summary": row["summary"]}

        if is_creator:
            brand_data[brand]["asCreator"].append(ticket)
        elif is_assignee:
            brand_data[brand]["asAssignee"].append(ticket)
        else:
            brand_data[brand]["other"].append(ticket)

    # Remove brands with zero tickets
    filtered = {
        b: buckets for b, buckets in brand_data.items()
        if buckets["asCreator"] or buckets["asAssignee"] or buckets["other"]
    }

    return {
        "personName": person["name"],
        "designation": person["designation"],
        "brands": filtered,
        "quarterLabel": QUARTER_LABEL,
    }


# ── ensureBandwidthLogTab ─────────────────────────────────────────────────────

def ensure_bandwidth_log_tab():
    svc = get_service()
    meta = svc.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if "Bandwidth_Log" in existing:
        return

    # Create tab
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": "Bandwidth_Log"}}}]}
    ).execute()

    # Write headers
    svc.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range="Bandwidth_Log!A1:H1",
        valueInputOption="RAW",
        body={"values": [["Ticket Key", "Summary", "Brand", "Person Name",
                          "Designation", "Role", "Bandwidth %", "Submitted At"]]}
    ).execute()


# ── logBandwidthSubmissions ───────────────────────────────────────────────────

def log_bandwidth_submissions(entries: list):
    """Appends rows to Bandwidth_Log. entries: [{key,summary,brand,personName,designation,role,bandwidth}]"""
    if not entries:
        return
    ensure_bandwidth_log_tab()
    from datetime import datetime, timezone
    now = datetime.now(timezone.utc).isoformat()
    values = [
        [e["key"], e.get("summary", ""), e.get("brand", ""),
         e["personName"], e.get("designation", ""),
         e.get("role", "Other"), e["bandwidth"], now]
        for e in entries
    ]
    svc = get_service()
    svc.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range="Bandwidth_Log!A:H",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    ).execute()


# ── getRecipientsFromTab ──────────────────────────────────────────────────────

def get_recipients_from_tab() -> list:
    """Read all recipients from Recipients tab. Returns [{name, email, designation}]."""
    people_map = build_people_map()
    rows = _get_values("Recipients!A:D")
    if not rows:
        return []

    result = []
    seen_email = set()
    for row in rows[1:]:  # skip header
        name = safe_get(row, 1)
        designation = safe_get(row, 2)
        email = safe_get(row, 3)

        if not name:
            continue
        # Fall back to email cache if tab cell is empty
        if not is_valid_email(email):
            email = build_email_cache().get(norm(name), "")
        if not is_valid_email(email):
            person = people_map.get(norm(name))
            if person:
                email = person.get("email", "")
        if not is_valid_email(email):
            continue
        if email in seen_email:
            continue
        seen_email.add(email)
        result.append({"name": name, "email": email, "designation": designation})

    return result


# ── Cache clear (testing) ─────────────────────────────────────────────────────

def clear_caches():
    global _service, _brand_map, _people_map, _email_cache
    _service = _brand_map = _people_map = _email_cache = None
