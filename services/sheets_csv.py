"""
Read-only Google Sheets access via public CSV export.
Works for any sheet shared as 'Anyone with the link can view'.
No service account required. Used for reads; writes still need sheets.py.
"""
import csv
import re
import io
import os
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env", override=True)

SHEET_ID = os.getenv("SHEET_ID")
QUARTER_LABEL = os.getenv("QUARTER_LABEL", "Q1'26")
BASE_URL = "https://docs.google.com/spreadsheets/d/{id}/gviz/tq?tqx=out:csv&sheet={sheet}"

# Module-level caches
_brand_map = None
_people_map = None
_email_cache = None
_cap_rows = None
_recipients_tab = None
_recipient_brands_map = None

KEY_RE = re.compile(r"([A-Z]+-\d+)")
NAME_EMAIL_RE = re.compile(r"^(.*?)\s*<[^>]+>$")


def _fetch_csv(sheet_name: str) -> list[list[str]]:
    url = BASE_URL.format(id=SHEET_ID, sheet=requests.utils.quote(sheet_name))
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    reader = csv.reader(io.StringIO(resp.text))
    return list(reader)


def norm(s: str) -> str:
    return str(s or "").strip().lower()


def safe(row: list, i: int) -> str:
    return row[i].strip() if i < len(row) else ""


def extract_key(raw: str) -> str | None:
    m = KEY_RE.search(raw)
    return m.group(1) if m else None


def is_valid_email(e: str) -> bool:
    s = str(e or "").strip()
    return "@" in s and "#N/A" not in s and s != ""


def extract_name(cell: str) -> str:
    """Extract plain name from 'Name <email@x.com>' or just 'Name'."""
    m = NAME_EMAIL_RE.match(cell.strip())
    return m.group(1).strip() if m else cell.strip()


# ── buildBrandMap ─────────────────────────────────────────────────────────────

def build_brand_map() -> dict:
    global _brand_map
    if _brand_map is not None:
        return _brand_map
    rows = _fetch_csv("Brand<>Product")
    result = {}
    for row in rows[1:]:
        if len(row) < 2:
            continue
        product_team, brand = row[0].strip(), row[1].strip()
        if product_team and brand:
            result[norm(product_team)] = brand
    _brand_map = result
    return _brand_map


# ── buildEmailCache ───────────────────────────────────────────────────────────
# The Emails tab has an unusual layout: the email for the person in row[i]
# is stored in col D (index 3) of row[i-1] (one row up).

def build_email_cache() -> dict:
    global _email_cache
    if _email_cache is not None:
        return _email_cache
    rows = _fetch_csv("Emails")
    result = {}
    # rows[0] is the header row — but col D of header row holds the first person's email.
    # rows[1] is first data row (person 1) — col D holds person 2's email, etc.
    # So: email for person at rows[i] (i >= 1) = safe(rows[i-1], 3)
    for i in range(1, len(rows)):
        name = safe(rows[i], 1)
        email = safe(rows[i - 1], 3)
        if name and is_valid_email(email):
            result[norm(name)] = email
    _email_cache = result
    return _email_cache


# ── buildPeopleMap ────────────────────────────────────────────────────────────

def build_people_map() -> dict:
    global _people_map
    if _people_map is not None:
        return _people_map
    rows = _fetch_csv("People<>Soln")
    result = {}
    for row in rows[1:]:
        name = safe(row, 1)
        if not name:
            continue
        designation = safe(row, 3)
        brands = [b for b in [safe(row, 4), safe(row, 5), safe(row, 6)] if b]
        email = safe(row, 7)
        result[norm(name)] = {
            "name": name,
            "designation": designation,
            "brands": brands,
            "email": email,
        }
    _people_map = result
    return _people_map


def get_email_for_person(name: str) -> str | None:
    ec = build_email_cache()
    pm = build_people_map()
    email = ec.get(norm(name))
    if email:
        return email
    person = pm.get(norm(name))
    if person and is_valid_email(person["email"]):
        return person["email"]
    return None


# ── getCapitalisedRows ────────────────────────────────────────────────────────

def get_capitalised_rows() -> list:
    global _cap_rows
    if _cap_rows is not None:
        return _cap_rows

    brand_map = build_brand_map()
    rows = _fetch_csv(QUARTER_LABEL)
    if not rows:
        return []

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
        cap = norm(safe(row, i_cap))
        if cap != "yes":
            continue
        key = extract_key(safe(row, i_key))
        if not key:
            continue

        product_team = safe(row, i_product_team)
        brand = safe(row, i_brand) if i_brand >= 0 else ""
        if not brand:
            brand = brand_map.get(norm(product_team), product_team)

        # PM+Director column has entries like "Name <email>, Name2 <email2>"
        pm_raw = safe(row, i_pm_director) if i_pm_director >= 0 else ""
        pm_directors = []
        for part in pm_raw.split(","):
            name = extract_name(part)
            if name:
                pm_directors.append(name)

        results.append({
            "key": key,
            "summary": safe(row, i_summary),
            "product_team": product_team,
            "brand": brand,
            "creator": safe(row, i_creator),
            "assignee": safe(row, i_assignee),
            "assignee_email": safe(row, i_assignee_email),
            "creator_email": safe(row, i_creator_email),
            "pm_directors": pm_directors,
        })

    _cap_rows = results
    return _cap_rows


# ── getAllRecipients ───────────────────────────────────────────────────────────

def get_all_recipients(rows: list, people_map: dict) -> list:
    """Returns every unique person who appears on any capitalised ticket.
    Falls back to sheet creator_email/assignee_email for people not in People<>Soln.
    Deduplicates by real email so reversed/shortened names don't create two entries."""
    seen_norm = set()        # normalised names already processed
    seen_email = {}          # real email → recipient already added
    PLACEHOLDER = "i_ojasvi.singh@innovaccer.com"
    recipients = []

    def add(name, fallback_email="", fallback_designation=""):
        if not name:
            return
        k = norm(name)
        if k in seen_norm:
            return
        seen_norm.add(k)

        person = people_map.get(k)
        if person:
            email = get_email_for_person(name)
            display_name = person["name"]
            designation = person["designation"]
        else:
            email = get_email_for_person(name) or fallback_email
            display_name = name
            designation = fallback_designation

        if not is_valid_email(email):
            return

        # If this is a placeholder email, check if a better entry already exists
        if email == PLACEHOLDER:
            return  # skip — a People<>Soln entry with real email will cover this person

        # Deduplicate by real email — keep the first (People<>Soln preferred)
        if email in seen_email:
            return
        seen_email[email] = True

        recipients.append({
            "name": display_name,
            "email": email,
            "designation": designation,
        })

    # Add People<>Soln entries first so they take priority over sheet name variants
    for person in people_map.values():
        add(person["name"])

    # Then add anyone from tickets not yet seen
    for row in rows:
        add(row["creator"], fallback_email=row["creator_email"])
        add(row["assignee"], fallback_email=row["assignee_email"])
        for pm in row["pm_directors"]:
            add(pm)

    return recipients


# ── _buildRecipientBrandsMap ──────────────────────────────────────────────────

def _build_recipient_brands_map() -> dict:
    """Returns {norm(name): [brand, ...]} from brand columns in the Recipients tab.
    Used as fallback when person is not found in People<>Soln."""
    global _recipient_brands_map
    if _recipient_brands_map is not None:
        return _recipient_brands_map
    result = {}
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={RECIPIENTS_GID}"
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        rows_raw = list(csv.reader(io.StringIO(resp.text)))
    except Exception:
        _recipient_brands_map = result
        return result

    if not rows_raw or len(rows_raw) < 2:
        _recipient_brands_map = result
        return result

    header = [norm(h) for h in rows_raw[0]]
    i_name = next((i for i, h in enumerate(header) if h == "name"), 1)
    brand_cols = [i for i, h in enumerate(header) if "brand" in h]

    for row in rows_raw[1:]:
        name = safe(row, i_name)
        if not name:
            continue
        brands = [safe(row, i) for i in brand_cols if safe(row, i)]
        if brands:
            result[norm(name)] = brands

    _recipient_brands_map = result
    return result


# ── getFormDataForPerson ──────────────────────────────────────────────────────

def get_form_data_for_person(name: str, rows: list, people_map: dict) -> dict | None:
    k = norm(name)
    person = people_map.get(k)

    if person:
        display_name = person["name"]
        designation = person["designation"]
        person_brands = set(b for b in person["brands"] if b)
    else:
        # Not in People<>Soln — derive everything from their tickets directly
        display_name = name
        designation = ""
        person_brands = set()

    # If still no brands, try the brand columns in the Recipients tab
    if not person_brands:
        person_brands = set(_build_recipient_brands_map().get(k, []))

    brand_data: dict = {}

    # ── Pass 1: tickets where person is creator / assignee / pm ──────────────
    seen_keys: set = set()
    for row in rows:
        brand = row["brand"]
        is_creator  = norm(row["creator"])  == k
        is_assignee = norm(row["assignee"]) == k
        is_pm       = any(norm(pm) == k for pm in row["pm_directors"])

        if not (is_creator or is_assignee or is_pm):
            continue

        # PM/Director: only show tickets from brands they own
        # Creator/Assignee: always show, regardless of brand
        if is_pm and not is_creator and not is_assignee:
            if person_brands and brand not in person_brands:
                continue

        if brand not in brand_data:
            brand_data[brand] = {"asCreator": [], "asAssignee": [], "other": []}

        ticket = {"key": row["key"], "summary": row["summary"]}
        seen_keys.add(row["key"])
        if is_creator:
            brand_data[brand]["asCreator"].append(ticket)
        elif is_assignee:
            brand_data[brand]["asAssignee"].append(ticket)
        else:
            brand_data[brand]["other"].append(ticket)

    # ── Pass 2: remaining brand tickets → "other" (rest) ─────────────────────
    # Ensures people with no creator/assignee tickets still see their brand's work
    for row in rows:
        brand = row["brand"]
        if brand not in person_brands:
            continue
        if row["key"] in seen_keys:
            continue
        if brand not in brand_data:
            brand_data[brand] = {"asCreator": [], "asAssignee": [], "other": []}
        brand_data[brand]["other"].append({"key": row["key"], "summary": row["summary"]})
        seen_keys.add(row["key"])

    if not brand_data:
        # No brand info and no tickets — still return a usable form (custom entries only)
        return {
            "personName": display_name,
            "designation": designation,
            "brands": {},
            "quarterLabel": QUARTER_LABEL,
        }

    return {
        "personName": display_name,
        "designation": designation,
        "brands": brand_data,
        "quarterLabel": QUARTER_LABEL,
    }


# ── getRecipientsFromTab ──────────────────────────────────────────────────────

# GID of the Recipients tab in the main sheet
RECIPIENTS_GID = "1121046660"

def get_recipients_from_tab() -> list:
    """Read recipients from the Recipients tab (GID 1121046660).
    Columns: Employee ID, Name, Designation, Brand 1, Brand 2, Brand 3.
    Email is looked up from the Emails/People<>Soln cache."""
    global _recipients_tab
    if _recipients_tab is not None:
        return _recipients_tab

    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={RECIPIENTS_GID}"
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        import csv as _csv
        import io as _io
        rows = list(_csv.reader(_io.StringIO(resp.text)))
    except Exception:
        rows = []

    if not rows or len(rows) < 2:
        _recipients_tab = []
        return _recipients_tab

    header = [norm(h) for h in rows[0]]
    i_name  = next((i for i, h in enumerate(header) if h == "name"), 1)
    i_desig = next((i for i, h in enumerate(header) if "designation" in h), 2)
    # Accept "email", "emails", "email address", "work email", etc.
    i_email = next((i for i, h in enumerate(header) if "email" in h), -1)

    result = []
    seen_email: set = set()
    for row in rows[1:]:
        name  = safe(row, i_name)
        desig = safe(row, i_desig)
        email = safe(row, i_email) if i_email >= 0 else ""
        if not name:
            continue
        # Fall back to cache lookup if tab email cell is empty
        if not is_valid_email(email):
            email = build_email_cache().get(norm(name)) or ""
        if not is_valid_email(email):
            person = build_people_map().get(norm(name))
            if person:
                email = person.get("email", "")
        if not is_valid_email(email):
            continue
        if email in seen_email:
            continue
        seen_email.add(email)
        result.append({"name": name, "email": email, "designation": desig})

    _recipients_tab = result
    return _recipients_tab


def clear_caches():
    global _brand_map, _people_map, _email_cache, _cap_rows, _recipients_tab, _recipient_brands_map
    _brand_map = _people_map = _email_cache = _cap_rows = _recipients_tab = _recipient_brands_map = None
