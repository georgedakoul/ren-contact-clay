"""discover_contacts.py — Clay contact discovery driven by actionable_contacts.json.

Usage:
  python discover_contacts.py                          → print status report
  python discover_contacts.py save                     → save the BATCH dict below
  python discover_contacts.py enrich_emails Brand1 … → print JSON for find-and-enrich-list-of-contacts
  python discover_contacts.py export                   → export all contacts to Excel
  python discover_contacts.py sync_sheets              → push new contacts (with email) to shared Google Sheet
  python discover_contacts.py mark_empty   Brand1 [Brand2 ...]
  python discover_contacts.py unmark_empty Brand1 [Brand2 ...]

3-phase Clay workflow:
  Phase 1 — Employee discovery:
    find-and-enrich-contacts-at-company(companyIdentifier=<domain>, numberOfContacts=PHASE1_NUM_CONTACTS, contactFilters={locations:["Greece"]})
    → raw employees (name + title + LinkedIn URL, no email needed yet)
  Phase 2 — Save + title filter:
    Edit BATCH dict, run `python discover_contacts.py save`
    → writes 00-BrandName.json, BANNED_TITLES applied automatically
  Phase 3 — Email enrichment:
    run `python discover_contacts.py enrich_emails BrandName`
    → prints JSON list; pass to find-and-enrich-list-of-contacts with dataPoints:{contactDataPoints:[{type:"Email"}]}
    → save result via BATCH as usual

File prefix convention in output/employees/:
  BrandName.json    → ✓ Covered   (has at least one email)
  00-BrandName.json → ∅ LinkedIn  (contacts found, no emails yet)
  ZZ-BrandName.json → ✗ Empty     (Clay returned 0 contacts — skip)
  (no file)         → ✗ Missing   (never searched)
"""
import json, sys, unicodedata, re, subprocess
from datetime import datetime, timezone
from pathlib import Path

ROOT          = Path(__file__).resolve().parents[3]
STORE_DIR     = ROOT / "AI Sales Agent System" / "output" / "employees"
CONTACTS_FILE = ROOT / "AI Sales Agent System" / "actionable_contacts.json"
EXPORT_PATH   = ROOT / "AI Sales Agent System" / "output" / "contacts_export.xlsx"

# Google Sheets sync — set SHEET_ID after creating the sheet (see SKILL.md setup guide)
SHEET_ID               = "1g3rpo6P2drPwhcDHWkLdmAqRCy3W6y6jbN_T5_THirw"
SHEETS_SERVICE_ACCOUNT = Path(__file__).parent / "service_account.json"
TODAY               = datetime.now(timezone.utc).strftime("%Y-%m-%d")
PHASE1_NUM_CONTACTS = 11   # numberOfContacts passed to find-and-enrich-contacts-at-company
STORE_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# DOMAIN_MAP: preferred companyIdentifier (domain) per brand.
# Always use domain over LinkedIn slug — Clay resolves the correct entity from domain.
# Add any brand here once its domain is known; brands not listed fall back to linkedin_slug.
# ---------------------------------------------------------------------------
DOMAIN_MAP = {
    # Slug resolves to wrong entity or returns 0 — domain is the fix
    "COSMOTE":           "cosmote.gr",
    "Village Cinemas":   "villagecinemas.gr",    # slug → Australian entity
    "Cosmos Sport":      "cosmossport.gr",
    "Coca-Cola":         "coca-colahellenic.com",
    "Psichogios Books":  "psichogios.gr",
    "Pame Stoixima":     "opap.gr",              # brand lives under OPAP
    "ION":               "ion.gr",
    "instacar":          "instacar.gr",          # slug → US company
    "Apivita":           "apivita.com",
    "more.com":          "more.com",
    "Wind":              "wind.gr",
    "Alterlife":         "alterlife.gr",
    "SKY express":       "skyexpress.gr",
    "Carroten":          "carroten.gr",          # owned by Sarantis; no standalone LinkedIn
    "Fresh Line":        "freshline.gr",
    "Alumil":            "alumil.com",
    "BSB Fashion":       "bsbfashion.com",
    "La Vie en Rose":    "lavieenrose.com",      # slug → Swiss NGO
    "Mind Your Style":   "mindyourstyle.gr",
    "Protergia":         "protergia.gr",
    # Known good domains from previous batches
    "Douleutaras":       "douleutaras.gr",
    "Box Now":           "boxnow.gr",
    "Germanos":          "germanos.gr",
    "MEVGAL":            "mevgal.gr",
    "Vitex":             "vitex.gr",
    "Three Cents":       "threecents.com",
    "Snappi":            "snappibank.com",
    "LG":                "lg.com",
    "JYSK":              "jysk.com",
    "Kinder":            "ferrerocareers.com",
    "NIVEA":             "beiersdorf.com",
    "Dove":              "unilever.com",
    "Converse":          "converse.com",
    "BMW":               "bmwgroup.com",
    "Ferryhopper":       "ferryhopper.com",
    "Vans":              "vans.com",
    "ANT1":              "ant1.gr",
    "AEK FC":            "aekfc.gr",
    "Jumbo":             "e-jumbo.gr",
    "FAGE":              "home.fage",
    "New Balance":       "newbalance.com",
    "Zara":              "inditexpeople.com",
    "Puma":              "puma.com",
    "CarVertical":       "carvertical.com",
    "Muagreece":              "muagreece.com",   # parent co BVELO operates MUA Greece brand
    "Emmanouela Cosmetics":   "emmanouelacosmetics.com",
    "Elixir Makeup":          "elixirmakeup.gr",
}

# All searches now use locations=["Greece"] — this is a Greek outreach list.
# GLOBAL_BRANDS kept for reference only (no longer drives filter logic).
GLOBAL_BRANDS = {
    "Samsung", "Apple", "Herbalife", "IKEA", "LEGO", "Huawei",
    "Motorola", "Starbucks", "Red Bull", "Wolt", "FREENOW",
    "Hertz", "Sony Music Entertainment", "Chanel", "Visa",
    "Nespresso", "Monster Energy", "La Vie en Rose",
}

# Job titles to ignore — contacts whose title contains any of these strings
# (case-insensitive, substring match) are skipped on save and purged on export.
BANNED_TITLES = {
    # Confirmed by user (specific titles)
    "chief of staff to the ceo",
    "design lead",
    "performance marketing manager",
    "assistant general manager",
    "chief executive officer at cosmote payments",   # specific person only, not all CEOs
    # Business development / commercial (not marketing decision-makers)
    "business development",         # BDM, head of BD, chief strategy & BD officer
    "commercial director",
    "commercial manager",
    "international business director",
    # Trade / shopper / channel (not influencer marketing decision-makers)
    "trade marketing",              # specialist, coordinator, lead, section manager, manager
    "shopper",                      # shopper marketing manager, head of shopper marketing
    # Exports (distribution, not brand marketing)
    "exports director",
    "export marketing manager",
    # Operations
    "director of partners operations",
    # HR / people / admin roles
    "head of people",               # head of people rewards, head of people & culture, etc.
    "organisational development",
    "executive assistant",
    # Analytics (not decision-makers)
    "business data analyst",
    # Keyword catch-all
    "sales",                        # any title containing "sales"
    # Noise / non-decision-maker
    "marketing department",         # not a real title
    "marketing team member",        # too junior / generic
    "jysk influencer",              # brand ambassador role, not marketing decision-maker
    "assistant brand marketing skip",  # support role
    # FLA-specific / internal codes
    "fla marketing director",
    # Performance/paid specialists that are executional, not decision-makers
    "performance marketing specialist",
}

# Exact-match bans (full title only — substring would catch valid roles).
BANNED_TITLES_EXACT = {
    "social media",                 # bare "Social Media" with no qualifier is noise
}

# Titles always kept regardless of BANNED_TITLES substring matches.
# Paid media roles are valid decision-makers for influencer spend conversations.
WANTED_TITLES = {
    "paid media manager",
    "paid media specialist",
}

_brand_names_cache = None

def _brand_names():
    global _brand_names_cache
    if _brand_names_cache is None:
        try:
            data = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
            _brand_names_cache = {normalize(b["brand"]) for b in data["brands"]}
        except Exception:
            _brand_names_cache = set()
    return _brand_names_cache


def _is_banned(title):
    t = normalize(title)
    if t in BANNED_TITLES_EXACT:
        return True
    if t in WANTED_TITLES:
        return False
    # Brand-prefixed titles (e.g. "BMW Marketing Manager") always kept.
    for brand in _brand_names():
        if t.startswith(brand + " "):
            return False
    return any(banned in t for banned in BANNED_TITLES)


def _strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def normalize(s):
    return re.sub(r"\s+", " ", _strip_accents(s).lower().strip())


def get_identifier(brand_name, linkedin_slug, website=None):
    if brand_name in DOMAIN_MAP:
        return DOMAIN_MAP[brand_name]
    if website:
        return website
    if linkedin_slug:
        return f"https://www.linkedin.com/company/{linkedin_slug}"
    return None  # no identifier — needs domain discovery via web search


def _employee_path(brand_name):
    """Return existing employee file; checks all prefix variants."""
    for prefix in ("", "00-", "ZZ-"):
        p = STORE_DIR / f"{prefix}{brand_name}.json"
        if p.exists():
            return p
    return STORE_DIR / f"{brand_name}.json"  # default for new files


def _sync_actionable(brand_name):
    """Update actionable_contacts.json stats for one brand from the employee file on disk."""
    try:
        data = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return
    brands = data if isinstance(data, list) else data.get("brands", [])
    entry = next((b for b in brands if b.get("brand") == brand_name), None)
    if entry is None:
        return
    info = _scan_store().get(brand_name)
    if info is None:
        entry["has_employee_file"] = False
        entry["employee_count"]    = 0
        entry["emails_known"]      = 0
        entry["employee_file"]     = None
        entry["coverage_status"]   = "No coverage"
    elif info["_prefix"] == "ZZ-":
        entry["has_employee_file"] = True
        entry["employee_count"]    = 0
        entry["emails_known"]      = 0
        entry["employee_file"]     = f"ZZ-{brand_name}.json"
        entry["coverage_status"]   = "Empty"
    elif info["emails"] > 0:
        entry["has_employee_file"] = True
        entry["employee_count"]    = info["total"]
        entry["emails_known"]      = info["emails"]
        entry["employee_file"]     = f"{brand_name}.json"
        entry["coverage_status"]   = "Covered"
    else:
        entry["has_employee_file"] = True
        entry["employee_count"]    = info["total"]
        entry["emails_known"]      = 0
        entry["employee_file"]     = f"00-{brand_name}.json"
        entry["coverage_status"]   = "LinkedIn-only"
    CONTACTS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def save_contacts(brand_name, contacts, domain=None):
    if not contacts:
        zz = STORE_DIR / f"ZZ-{brand_name}.json"
        zz.write_text("[]", encoding="utf-8")
        for old_prefix in ("00-",):
            old = STORE_DIR / f"{old_prefix}{brand_name}.json"
            if old.exists():
                old.unlink()
        _sync_actionable(brand_name)
        print(f"  {brand_name}: Clay returned 0 contacts → ZZ-{brand_name}.json (skipped in future batches)")
        return
    # If we now have contacts, remove any ZZ- marker
    zz = STORE_DIR / f"ZZ-{brand_name}.json"
    if zz.exists():
        zz.unlink()
    path = _employee_path(brand_name)
    existing = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    by_name = {normalize(e["name"]): e for e in existing}
    # Purge any existing contacts with banned titles before merging
    by_name = {k: v for k, v in by_name.items() if not _is_banned(v.get("job_title") or "")}
    added = email_added = 0
    for c in contacts:
        name = (c.get("name") or "").strip()
        if not name:
            continue
        title  = (c.get("latest_experience_title") or "").strip()
        if _is_banned(title):
            continue
        key = normalize(name)
        li_url = c.get("url") or None
        em     = (c.get("email") or "").strip() or None
        dom    = c.get("domain") or domain or None
        if key in by_name:
            if title  and not by_name[key].get("job_title"):    by_name[key]["job_title"]    = title
            if li_url and not by_name[key].get("linkedin_url"): by_name[key]["linkedin_url"] = li_url
            if em     and not by_name[key].get("email"):
                by_name[key]["email"] = em
                email_added += 1
            by_name[key]["last_seen"] = TODAY
        else:
            by_name[key] = {
                "name": name, "job_title": title or None, "email": em,
                "linkedin_url": li_url, "verified": False, "active": True,
                "source": "clay", "domain": dom, "first_seen": TODAY, "last_seen": TODAY,
            }
            added += 1
            if em: email_added += 1
    merged = list(by_name.values())
    has_email = any(c.get("email") for c in merged)
    new_path = STORE_DIR / (f"{brand_name}.json" if has_email else f"00-{brand_name}.json")
    new_path.write_text(json.dumps(merged, indent=2, ensure_ascii=False), encoding="utf-8")
    if path.exists() and path != new_path:
        path.unlink()
    _sync_actionable(brand_name)
    print(f"  {brand_name}: {added} new contacts, {email_added} emails added → {len(merged)} total")


def list_for_enrichment(brands):
    out = []
    for b in brands:
        path = STORE_DIR / f"00-{b}.json"
        if not path.exists():
            print(f"  {b}: no 00-{b}.json (not linkedin-only)")
            continue
        contacts = json.loads(path.read_text(encoding="utf-8"))
        domain = DOMAIN_MAP.get(b)
        for c in contacts:
            if c.get("email"):
                continue
            out.append({"contactName": c["name"], "companyIdentifier": domain or c.get("domain") or b})
    print(json.dumps(out, indent=2, ensure_ascii=False))


def _scan_store():
    """Scan ALL employee files → {brand_name: {total, emails, last_seen, _prefix}}."""
    store = {}
    for f in sorted(STORE_DIR.glob("*.json")):
        if f.name == "desktop.ini":
            continue
        stem = f.stem
        if stem.startswith("ZZ-"):
            brand_name = stem[3:]
            store[brand_name] = {"total": 0, "emails": 0, "last_seen": None, "_prefix": "ZZ-"}
        elif stem.startswith("00-"):
            brand_name = stem[3:]
            contacts = json.loads(f.read_text(encoding="utf-8"))
            email_count = sum(1 for c in contacts if c.get("email"))
            last_seen = max((c.get("last_seen") or "1970-01-01") for c in contacts) if contacts else None
            store[brand_name] = {"total": len(contacts), "emails": email_count, "last_seen": last_seen, "_prefix": "00-"}
        else:
            brand_name = stem
            contacts = json.loads(f.read_text(encoding="utf-8"))
            email_count = sum(1 for c in contacts if c.get("email"))
            last_seen = max((c.get("last_seen") or "1970-01-01") for c in contacts) if contacts else None
            store[brand_name] = {"total": len(contacts), "emails": email_count, "last_seen": last_seen, "_prefix": ""}
    return store


BRANDS_LOOKUP_FILE = Path(__file__).parent / "data" / "brands_consolidated.csv"


def _load_brand_lookup():
    """brand name (lowercased) -> {num, times_advertised, unique_profiles, industry} from brands_consolidated.csv"""
    lookup = {}
    if not BRANDS_LOOKUP_FILE.exists():
        return lookup
    import csv
    with open(BRANDS_LOOKUP_FILE, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            name = (row.get("Brand") or "").strip()
            if not name:
                continue
            lookup[name.lower()] = {
                "num":              row.get("#", "").split(".")[0] if row.get("#") else "",
                "times_advertised": row.get("Times Advertised", "").split(".")[0] if row.get("Times Advertised") else "",
                "unique_profiles":  row.get("Unique Profiles", "").split(".")[0] if row.get("Unique Profiles") else "",
                "industry":         row.get("Industry", "").strip(),
            }
    return lookup


# Real shared-sheet header (Master tab). Columns beyond L are pipeline/report
# fields owned by other skills — never write into them from here.
SHEET_HEADER = ["#", "Company", "Times_Advertised", "Unique_Profiles", "Industry",
                "Full Name", "Job Title", "Email", "LinkedIn", "Tier", "Persona", "Status"]
SHEET_FMT_RANGE_COLS = "A:J"   # per-project convention: grey new rows on insert, cols A-J only
SHEET_FMT_BG = {"red": 0.906, "green": 0.902, "blue": 0.902}  # #e7e6e6

# Low-quality email signals — skip these contacts entirely, never sync to sheet.
# Confirmed 2026-07-01 cleanup: generic role inboxes and third-party (personal /
# law-tax-accounting-firm) domains are noise, not real marketing contacts.
GENERIC_LOCAL_PARTS = {"info", "contact", "sales", "support", "hello", "office",
                        "marketing", "pr", "press", "accounting", "invoice"}
PERSONAL_EMAIL_DOMAINS = {"gmail.com", "hotmail.com", "hotmail.gr", "yahoo.com", "yahoo.gr",
                           "outlook.com", "mac.com", "rocketmail.com", "icloud.com"}
# substrings, not exact domains — these are outside accountants/lawyers, not brand employees
PRO_SERVICES_DOMAIN_KEYWORDS = ["law", "legal", "tax", "ecovis", "andersen", "mcbainscooper",
                                 "martzoukos", "gt.com", "capitalpartners", "bernitsaslaw",
                                 "firstfloor", "syntaxis", "altaxis", "365taccs"]


def _is_low_quality_email(email):
    local, _, domain = email.lower().partition("@")
    if local in GENERIC_LOCAL_PARTS:
        return True
    if domain in PERSONAL_EMAIL_DOMAINS:
        return True
    if any(kw in domain for kw in PRO_SERVICES_DOMAIN_KEYWORDS):
        return True
    if local[:1].isdigit():
        return True
    return False


# IMPORTANT — do NOT add automatic "domain must match company name" filtering.
# Confirmed 2026-07-01: many brands' real marketing contacts sit on a parent/
# subsidiary domain that shares no substring with the storefront brand name —
# e.g. Pame Stoixima -> opap.gr, Hellmann's -> unilever.com, Vichy -> loreal.com,
# Pantene -> pg.com, Coca-Cola -> cchellenic.com, Minos EMI -> umusic.com,
# COSMOTE -> ote.gr. A naive brand/domain string match flags ~95% of these as
# "mismatches" and would delete legitimate contacts. Only the two signals above
# (generic role inbox, personal/professional-services domain) are safe to
# auto-filter; anything else needs a human to eyeball it.


def sync_to_sheets():
    """Insert new contacts into the shared Google Sheet, grouped under their brand's
    existing row block — never a bottom dump. A contact only qualifies if it has BOTH
    an email AND a LinkedIn URL (partial contacts are not synced), and the email isn't
    low-quality (generic role inbox, personal email provider, or a law/tax/accounting
    firm domain — see _is_low_quality_email). Dedup by email. Times_Advertised /
    Unique_Profiles / # are looked up by brand name from data/brands_consolidated.csv.

    HARD RULE: this function only ever inserts or appends rows. It must never delete,
    clear, or overwrite any existing cell in the sheet — this is company data owned
    outside this pipeline. If a future change needs to touch existing rows, stop and
    get explicit sign-off first; do not add delete/overwrite logic here."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("ERROR: gspread not installed. Run: pip3 install gspread google-auth"); return

    if not SHEETS_SERVICE_ACCOUNT.exists():
        print(f"ERROR: {SHEETS_SERVICE_ACCOUNT} not found — cannot authenticate."); return

    SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds  = Credentials.from_service_account_file(str(SHEETS_SERVICE_ACCOUNT), scopes=SCOPES)
    gc     = gspread.authorize(creds)

    try:
        sh = gc.open_by_key(SHEET_ID)
    except Exception as e:
        print(f"ERROR opening sheet {SHEET_ID}: {e}"); return

    ws = sh.sheet1
    existing = ws.get_all_values()
    if not existing:
        print(f"ERROR: sheet is empty — expected header row {SHEET_HEADER}. Not writing blind."); return
    if existing[0][:len(SHEET_HEADER)] != SHEET_HEADER:
        print("ERROR: sheet header doesn't match expected schema — refusing to write "
              "(would misalign columns). Check SHEET_HEADER against the live sheet."); return

    EMAIL_COL = SHEET_HEADER.index("Email")        # 7
    COMPANY_COL = SHEET_HEADER.index("Company")    # 1

    existing_emails = {row[EMAIL_COL].lower().strip() for row in existing[1:]
                        if len(row) > EMAIL_COL and row[EMAIL_COL]}

    # last existing sheet row (1-indexed) per brand — for grouped insertion
    brand_last_row = {}
    for i, row in enumerate(existing[1:], start=2):
        if len(row) > COMPANY_COL and row[COMPANY_COL].strip():
            brand_last_row[row[COMPANY_COL].strip().lower()] = i

    lookup = _load_brand_lookup()

    # brand -> list of new row-value lists (in SHEET_HEADER column order)
    per_brand_new_rows = {}
    for f in sorted(STORE_DIR.glob("*.json")):
        if f.stem.startswith(("ZZ-", "00-")):
            continue
        brand = f.stem
        try:
            contacts = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        info = lookup.get(brand.lower(), {})
        for c in contacts:
            em = (c.get("email") or "").strip()
            li = (c.get("linkedin_url") or "").strip()
            if not em or not li or em.lower() in existing_emails:
                continue
            if _is_low_quality_email(em):
                continue
            per_brand_new_rows.setdefault(brand, []).append([
                info.get("num", ""),
                brand,
                info.get("times_advertised", ""),
                info.get("unique_profiles", ""),
                info.get("industry", ""),
                c.get("name") or "",
                c.get("job_title") or "",
                em,
                c.get("linkedin_url") or "",
                "", "", "",  # Tier, Persona, Status — filled by downstream skills, not here
            ])
            existing_emails.add(em.lower())

    if not per_brand_new_rows:
        print("Google Sheet is already up-to-date — no new contacts to add.")
        return

    # Insert brand-by-brand, bottom-up by target row, so earlier inserts don't
    # shift the row indices of brands still queued for insertion.
    inserts = []  # (target_row, rows)
    tail_rows = []  # brands with no existing block — appended once at the very end
    for brand, rows in per_brand_new_rows.items():
        last_row = brand_last_row.get(brand.lower())
        if last_row:
            inserts.append((last_row, rows))
        else:
            tail_rows.extend(rows)
    inserts.sort(key=lambda x: x[0], reverse=True)

    total = 0
    formatted_ranges = []
    for target_row, rows in inserts:
        ws.insert_rows(rows, row=target_row + 1, value_input_option="USER_ENTERED")
        formatted_ranges.append((target_row + 1, target_row + len(rows)))
        total += len(rows)
        print(f"  Inserted {len(rows)} row(s) under '{rows[0][1]}' at row {target_row + 1} ({total} so far)...")

    if tail_rows:
        next_row = len(ws.get_all_values()) + 1
        ws.append_rows(tail_rows, value_input_option="USER_ENTERED")
        formatted_ranges.append((next_row, next_row + len(tail_rows) - 1))
        total += len(tail_rows)
        print(f"  Appended {len(tail_rows)} row(s) for brand(s) with no existing sheet block.")

    # Format new rows: grey background on A:J only, batched into one call
    fmt_body = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": start - 1,
                        "endRowIndex": end,
                        "startColumnIndex": 0,
                        "endColumnIndex": 10,  # A:J
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": SHEET_FMT_BG}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            }
            for start, end in formatted_ranges
        ]
    }
    if fmt_body["requests"]:
        sh.batch_update(fmt_body)

    print(f"Done. {total} new contacts inserted into Google Sheet, grouped by brand.")


def git_sync():
    result = subprocess.run(
        ["git", "add",
         str(STORE_DIR),
         str(CONTACTS_FILE),
         str(Path(__file__))],
        cwd=ROOT, capture_output=True, text=True
    )
    if result.returncode != 0:
        print(result.stderr); return
    status = subprocess.run(["git", "status", "--short"], cwd=ROOT, capture_output=True, text=True)
    if not status.stdout.strip():
        print("Nothing to commit."); return
    msg = f"clay-batch: sync employee store {datetime.now(timezone.utc).strftime('%Y-%m-%d')}"
    result = subprocess.run(["git", "commit", "-m", msg], cwd=ROOT, capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stderr); return
    print(result.stdout.strip())
    result = subprocess.run(["git", "push"], cwd=ROOT, capture_output=True, text=True)
    print(result.stdout.strip() or result.stderr.strip())


def status_report():
    store = _scan_store()
    total_contacts = sum(v["total"] for v in store.values())
    total_emails   = sum(v["emails"] for v in store.values())
    brands_with_emails = sum(1 for v in store.values() if v["emails"] > 0)

    print(f"\n=== EMPLOYEE STORE OVERVIEW ===")
    print(f"  Files on disk : {len(store):>4}  ({brands_with_emails} with emails)")
    print(f"  Total contacts: {total_contacts:>4}  ({total_emails} with emails)")

    data      = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
    ac_brands = {b["brand"]: b for b in data["brands"]}
    searchable = ac_brands

    covered      = []  # BrandName.json — has ≥1 email
    linkedin_only = [] # 00-BrandName.json — contacts but 0 emails
    empty        = []  # ZZ-BrandName.json — Clay returned 0 contacts
    missing      = []  # no file — has identifier (slug/override/website), never searched
    needs_domain = []  # no file — no identifier at all, needs web search to find domain

    for name in sorted(searchable):
        slug       = searchable[name].get("linkedin_slug") or ""
        website    = searchable[name].get("website") or None
        identifier = get_identifier(name, slug, website)
        info       = store.get(name)
        if identifier is None:
            needs_domain.append((name,))
        elif info is None:
            missing.append((name, identifier))
        elif info["_prefix"] == "ZZ-":
            empty.append((name, identifier))
        elif info["emails"] > 0:
            covered.append((name, identifier, info))
        else:
            linkedin_only.append((name, identifier, info))

    print(f"\n=== CLAY-SEARCHABLE BRANDS ({len(searchable)} total) ===")
    print(f"  ✓ Covered       (BrandName.json   — has emails)        : {len(covered)}")
    print(f"  ∅ LinkedIn-only (00-BrandName.json — contacts, 0 email): {len(linkedin_only)}")
    print(f"  ✗ Empty         (ZZ-BrandName.json — Clay returned 0)  : {len(empty)}")
    print(f"  ✗ Missing       (has identifier    — never searched)   : {len(missing)}")
    print(f"  ? Needs-domain  (no identifier     — web search needed): {len(needs_domain)}")

    if covered:
        print(f"\n--- ✓ COVERED (have emails) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'#C':>4} {'#E':>4} {'Last seen':<12}")
        print("  " + "─" * 100)
        for name, identifier, info in covered:
            print(f"  {name:<32} {identifier:<45} {info['total']:>4} {info['emails']:>4} {info['last_seen'] or '—':<12}")

    if linkedin_only:
        print(f"\n--- ∅ LINKEDIN-ONLY (contacts found, 0 emails) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'#C':>4}  {'Last seen':<12}")
        print("  " + "─" * 96)
        for name, identifier, info in linkedin_only:
            print(f"  {name:<32} {identifier:<45} {info['total']:>4}  {info['last_seen'] or '—':<12}")

    if empty:
        print(f"\n--- ✗ EMPTY (ZZ- files — Clay returned 0, excluded from next batch) ---")
        print(f"  {'Brand':<32} {'Identifier':<55}")
        print("  " + "─" * 90)
        for name, identifier in empty:
            print(f"  {name:<32} {identifier:<55}")

    if missing:
        print(f"\n--- ✗ MISSING (has identifier — never searched) ---")
        print(f"  {'Brand':<32} {'Identifier':<45}")
        print("  " + "─" * 80)
        for name, identifier in missing:
            print(f"  {name:<32} {identifier:<45}")

    if needs_domain:
        print(f"\n--- ? NEEDS-DOMAIN (no slug/website — web search required) ---")
        print(f"  {'Brand':<32}")
        print("  " + "─" * 34)
        for (name,) in needs_domain[:20]:
            print(f"  {name:<32}  [needs-domain]")
        if len(needs_domain) > 20:
            print(f"  ... and {len(needs_domain) - 20} more")

    # Next batch: missing first (have identifiers), then linkedin_only (retry for emails),
    # then needs_domain (require web search — listed last, agent handles via WebSearch).
    # Within each tier, Beauty industry brands come first.
    _is_beauty = lambda n: 0 if searchable.get(n, {}).get("industry", "").lower() == "beauty" else 1
    next_up  = sorted([(n, i, "missing")         for n, i    in missing],       key=lambda x: _is_beauty(x[0]))
    next_up += sorted([(n, i, "linkedin-only")   for n, i, _ in linkedin_only], key=lambda x: _is_beauty(x[0]))
    next_up += sorted([(n, None, "needs-domain") for (n,)    in needs_domain],  key=lambda x: _is_beauty(x[0]))

    if next_up:
        print(f"\n=== NEXT BATCH RECOMMENDATION (top 22) ===")
        print(f"  {'#':<4} {'Tier':<14} {'Brand':<32} {'Identifier'}")
        print("  " + "─" * 100)
        for idx, (name, identifier, tier) in enumerate(next_up[:22], 1):
            ident_str = identifier or "[needs-domain — use WebSearch]"
            print(f"  [{idx:>2}] {tier:<14}  {name:<32} {ident_str}")
    else:
        print("\nAll searchable brands are covered or pending LinkedIn-only retry.")



# ---------------------------------------------------------------------------
# BATCH — populated by Claude after each round of Clay MCP calls.
# Format: brand_name → (contacts_list, domain_hint)
# contacts_list items: {"name", "latest_experience_title", "url", "email", "domain"}
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# BATCH — populate before each run, then: python discover_contacts.py save
# Format: brand_name → (contacts_list, domain_hint)
# contacts_list items: {"name", "latest_experience_title", "url", "email", "domain"}
# After save completes, clear this dict (data is in employee files + git history).
# ---------------------------------------------------------------------------
BATCH = {
    "Kiko Milano": ([
        {"name": "Ioanna Angelis", "latest_experience_title": "Marketing Director Kalogirou S.A.", "url": "https://www.linkedin.com/in/ioanna-angelis-05724620/", "email": "ioanna.angelis@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Melina Zineli - Prastakou", "latest_experience_title": "Brand Marketing Manager-Luxury Division", "url": "https://www.linkedin.com/in/melina-zineli-prastakou-5ba516a3/", "email": "melina.zineli@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Paris Kapsalis", "latest_experience_title": "Luxury Buyer-Brand Manager at KALOGIROU", "url": "https://www.linkedin.com/in/paris-kapsalis/", "email": "paris.kapsalis@faisgroup.gr", "domain": "faisgroup.com"},
    ], "faisgroup.com"),
    "Prada": ([
        {"name": "Ioanna Angelis", "latest_experience_title": "Marketing Director Kalogirou S.A.", "url": "https://www.linkedin.com/in/ioanna-angelis-05724620/", "email": "ioanna.angelis@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Melina Zineli - Prastakou", "latest_experience_title": "Brand Marketing Manager-Luxury Division", "url": "https://www.linkedin.com/in/melina-zineli-prastakou-5ba516a3/", "email": "melina.zineli@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Paris Kapsalis", "latest_experience_title": "Luxury Buyer-Brand Manager at KALOGIROU", "url": "https://www.linkedin.com/in/paris-kapsalis/", "email": "paris.kapsalis@faisgroup.gr", "domain": "faisgroup.com"},
    ], "faisgroup.com"),
    "Dior": ([
        {"name": "Sevasti Polidouli", "latest_experience_title": "Retail Manager", "url": "https://www.linkedin.com/in/sevasti-polidouli-a0814b163/", "email": None, "domain": "dior.com"},
        {"name": "Ioanna Angelis", "latest_experience_title": "Marketing Director Kalogirou S.A.", "url": "https://www.linkedin.com/in/ioanna-angelis-05724620/", "email": "ioanna.angelis@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Melina Zineli - Prastakou", "latest_experience_title": "Brand Marketing Manager-Luxury Division", "url": "https://www.linkedin.com/in/melina-zineli-prastakou-5ba516a3/", "email": "melina.zineli@faisgroup.gr", "domain": "faisgroup.com"},
        {"name": "Paris Kapsalis", "latest_experience_title": "Luxury Buyer-Brand Manager at KALOGIROU", "url": "https://www.linkedin.com/in/paris-kapsalis/", "email": "paris.kapsalis@faisgroup.gr", "domain": "faisgroup.com"},
    ], "faisgroup.com"),
    "Emmanouela Cosmetics": ([
        {"name": "Athina Tsiringouli", "latest_experience_title": "Digital Marketing Executive", "url": "https://www.linkedin.com/in/athina-tsiringouli-570473195/", "email": None, "domain": "emmanouelacosmetics.com"},
    ], "emmanouelacosmetics.com"),
    "Seventeen Cosmetics": ([
        {"name": "Iro Prinia", "latest_experience_title": "Export Marketing Manager", "url": "https://www.linkedin.com/in/iro-prinia-33049541/", "email": "i.prinia@hellenica.gr", "domain": "hellenica.gr"},
        {"name": "Irma Fragkogianni - Matsa", "latest_experience_title": "eCommerce Growth Lead", "url": "https://www.linkedin.com/in/irma-fragkogianni-matsa-b59b401a3/", "email": "ifra@hellenica.gr", "domain": "hellenica.gr"},
        {"name": "Lydia Christoforidou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/lchristoforidou/", "email": None, "domain": "hellenica.gr"},
        {"name": "Silia Karatza", "latest_experience_title": "Digital Marketing & Social Media Specialist", "url": "https://www.linkedin.com/in/silia-karatza-79b726197/", "email": "s.karatza@hellenica.gr", "domain": "hellenica.gr"},
        {"name": "Katerina Efstathiou", "latest_experience_title": "Ecommerce Manager", "url": "https://www.linkedin.com/in/katerina-efstathiou-360887171/", "email": "k.efstathiou@hellenica.gr", "domain": "hellenica.gr"},
        {"name": "Electra Papanastasiou", "latest_experience_title": "Marketing Manager, Strategy Lead Seventeen Cosmetics", "url": "https://www.linkedin.com/in/papanastasiou-electra/", "email": "e.papanastasiou@hellenica.gr", "domain": "hellenica.gr"},
    ], "hellenica.gr"),
    "Coca-Cola": ([
        {"name": "ANTHIE DIMAKOU", "latest_experience_title": "Events Marketing Manager", "url": "https://www.linkedin.com/in/anthidimakou/", "email": "adimakou@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Ourania Xaplanteri", "latest_experience_title": "Channel Marketing Activation Executive", "url": "https://www.linkedin.com/in/ouraniaxaplanteri/", "email": "ourania.xaplanteri@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Ippokratis Kourkoumpas", "latest_experience_title": "Marketing Capability Development Manager", "url": "https://www.linkedin.com/in/ippokratis-kourkoumpas-915526b/", "email": "ippokratis.kourkoumpas@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Konstantinos Dolkas", "latest_experience_title": "Digital Factory Manager - Integrations", "url": "https://www.linkedin.com/in/konstantinos-dolkas-b308a132/", "email": "konstantinos.dolkas@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Dimitris Alevizopoulos", "latest_experience_title": "Digital Product Manager - Contact Center & Order Management", "url": "https://www.linkedin.com/in/dimitris-alevizopoulos-b1929927/", "email": "dimitris.alevizopoulos@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Eleni Seferi", "latest_experience_title": "Trade Marketing Manager Premium Spirits", "url": "https://www.linkedin.com/in/eleni-seferi-baaaa563/", "email": "eleni.seferi@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Dimitris Lianos", "latest_experience_title": "Digital Enterprise Platform Leader GR/CY", "url": "https://www.linkedin.com/in/dimitris-lianos-4243b7b3/", "email": "dimitris.lianos@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Dionysia Simatou", "latest_experience_title": "CCH Brands Marketing Director", "url": "https://www.linkedin.com/in/dionysia-simatou-50a38116/", "email": "dionysia.simatou@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Koukouvaos Georgios", "latest_experience_title": "Group Digital & Technology Platform Services IIoT Engineer", "url": "https://www.linkedin.com/in/koukouvaos-georgios-592b1063/", "email": "georgios.koukouvaos@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Danai Dikaiou", "latest_experience_title": "Senior Brand Manager Premium Spirits", "url": "https://www.linkedin.com/in/danai-dikaiou-332907146/", "email": "danai.dikaiou@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Dimitris Stathopoulos", "latest_experience_title": "Group Digital Platform Manager -Supply Chain Planning", "url": "https://www.linkedin.com/in/dimitris-stathopoulos-a7718222/", "email": "dimitris.stathopoulos@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Antonios Patarias", "latest_experience_title": "Data Science & AI Manager - Digital Commerce", "url": "https://www.linkedin.com/in/patarias-analytics/", "email": "antonios.patarias@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Jozsef Juhasz", "latest_experience_title": "Head of Digital Employee Platform", "url": "https://www.linkedin.com/in/jozsef-juhasz-4045243/", "email": "jozsef.juhasz@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Maria Konisioti", "latest_experience_title": "Brand Manager Beer", "url": "https://www.linkedin.com/in/maria-konisioti-2a501271/", "email": "maria.konisioti@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Ioanna Stroubi", "latest_experience_title": "CCH Brands Marketing Director", "url": "https://www.linkedin.com/in/ioanna-stroubi-a5470060/", "email": "ioanna.stroubi@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Christos Balatsos", "latest_experience_title": "Senior Brand Marketing Manager Water", "url": "https://www.linkedin.com/in/christos-balatsos-330a2546/", "email": "christos.balatsos@cchellenic.com", "domain": "coca-colahellenic.com"},
        {"name": "Chara Braouzi", "latest_experience_title": "Digital Enterprise Platform Lead", "url": "https://www.linkedin.com/in/charikleia-braouzi-a5aa835b/", "email": "chara.braouzi@cchellenic.com", "domain": "coca-colahellenic.com"},
    ], "coca-colahellenic.com"),
    "Estee Lauder": ([
        {"name": "Georgia Tselondre", "latest_experience_title": "Consumer Marketing Manager Balkans | Tom Ford", "url": "https://www.linkedin.com/in/georgia-tselondre-48607821/", "email": "gtselondre@estee.com", "domain": "elcompanies.com"},
        {"name": "Natalia Lamprou", "latest_experience_title": "Consumer Marketing Manager Clinique, Balkans", "url": "https://www.linkedin.com/in/natalialamprou/", "email": "nlamprou@estee.com", "domain": "elcompanies.com"},
        {"name": "Valia Kanellopoulou", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/vasiliki-kanellopoulou-b2651557/", "email": "vkanellopoulou@estee.com", "domain": "elcompanies.com"},
        {"name": "Michail - Aggelos Pappas", "latest_experience_title": "EMEA Senior Digital Merchandising Specialist", "url": "https://www.linkedin.com/in/michail-aggelos-pappas-460591131/", "email": "apappas@elcompanies.com", "domain": "elcompanies.com"},
        {"name": "Athanasia Makri", "latest_experience_title": "Trade Marketing Executive", "url": "https://www.linkedin.com/in/athanasia-makri-17b9bb134/", "email": "amakri@estee.com", "domain": "elcompanies.com"},
        {"name": "Effie Nezou", "latest_experience_title": "Consumer Marketing Manager Estee Lauder", "url": "https://www.linkedin.com/in/effie-nezou-74a91bb2/", "email": "enezou@elcompanies.com", "domain": "elcompanies.com"},
        {"name": "Maria Itsiou", "latest_experience_title": "Consumer Marketing Executive Tom Ford Beauty", "url": "https://www.linkedin.com/in/maria-itsiou-923797106/", "email": "mitsiou@estee.com", "domain": "elcompanies.com"},
        {"name": "Katerina Moschopoulou", "latest_experience_title": "Product Marketing Executive | Jo Malone London", "url": "https://www.linkedin.com/in/katerina-moschopoulou-177263180/", "email": "kmoschopoulou@estee.com", "domain": "elcompanies.com"},
        {"name": "Melina Mokka", "latest_experience_title": "Digital Paid Coordinator", "url": "https://www.linkedin.com/in/melina-mokka-0b7349178/", "email": "mmokka@estee.com", "domain": "elcompanies.com"},
        {"name": "Christina K.", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/christina-k-722838ab/", "email": "ckatic@elcompanies.com", "domain": "elcompanies.com"},
        {"name": "Alex Nikas", "latest_experience_title": "Digital Commerce & Marketing Director (Balkans)", "url": "https://www.linkedin.com/in/alex-nikas-0109423/", "email": "anikas@elcompanies.com", "domain": "elcompanies.com"},
        {"name": "Eirini Patsaki", "latest_experience_title": "360 Marketing Executive, Tom Ford & Balmain - Balkans", "url": "https://www.linkedin.com/in/eirini-patsaki/", "email": "epatsaki@estee.com", "domain": "elcompanies.com"},
        {"name": "Theodoros Arapogiannis", "latest_experience_title": "Trade Marketing Executive Clinique", "url": "https://www.linkedin.com/in/theodore-arapogiannis/", "email": "tarapogiannis@estee.com", "domain": "elcompanies.com"},
        {"name": "Alexandra Venouka", "latest_experience_title": "Consumer Marketing Manager Darphin & Origins", "url": "https://www.linkedin.com/in/alexandra-venouka-9b82ba142/", "email": "avenouka@aveda.co.uk", "domain": "elcompanies.com"},
        {"name": "Maria Saridaki", "latest_experience_title": "360 Marketing Executive Clinique", "url": "https://www.linkedin.com/in/maria-saridaki-4817551a3/", "email": None, "domain": "elcompanies.com"},
        {"name": "Freida Paisidou", "latest_experience_title": "360 Marketing Executive", "url": "https://www.linkedin.com/in/freidapaisidou/", "email": None, "domain": "elcompanies.com"},
    ], "elcompanies.com"),
    "MUA Makeup Academy": ([
        {"name": "Panagiotis Velonas", "latest_experience_title": "Founder | BVELO", "url": "https://www.linkedin.com/in/panagiotis-velonas-886254353/", "email": "pvelonas@bvelo.gr", "domain": "muagreece.com"},
        {"name": "Maria Oikonomou", "latest_experience_title": "Digital Marketing & Performance Strategist", "url": "https://www.linkedin.com/in/maria-oikonomou-28938a1aa/", "email": "maria.oikonomou@bravescale.gr", "domain": "muagreece.com"},
        {"name": "Angeliki Gidiotou", "latest_experience_title": "Influencer Marketing and UGC Specialist", "url": "https://www.linkedin.com/in/angeliki-gidiotou/", "email": None, "domain": "muagreece.com"},
    ], "muagreece.com"),
    "Benefit Cosmetics": ([{"name": "Marietta Fameli", "latest_experience_title": "Digital Project Manager", "url": "https://www.linkedin.com/in/marietta-fameli-622a692b/", "email": "mariettaf@benefitcosmetics.com", "domain": "benefitcosmetics.com"}, {"name": "Afroditi Delipanagioti", "latest_experience_title": "Digital/PR & Marketing manager", "url": "https://www.linkedin.com/in/afroditi-delipanagioti-198406244/", "email": "afroditid@benefitcosmetics.com", "domain": "benefitcosmetics.com"}], "benefitcosmetics.com"),
    "Borotalco": ([{"name": "Maria Kotsifakou", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/maria-kotsifakou/", "email": "mkotsifakou@boltonadhesives.com", "domain": "bolton.com"}, {"name": "Chrysa Printziou", "latest_experience_title": "Marketing Specialist UHU/Bison", "url": "https://www.linkedin.com/in/cprintziou/", "email": "cprintziou@boltonadhesives.com", "domain": "bolton.com"}], "bolton.com"),
    "Dust and Cream": ([{"name": "Despoina  E. Ananiadou", "latest_experience_title": "Marketing Coordinator", "url": "https://www.linkedin.com/in/despoina-e-ananiadou-0aa383120/", "email": None, "domain": "dustandcream.gr"}, {"name": "Nikolaos Apostolidis", "latest_experience_title": "Digital Marketing Specialist", "url": "https://www.linkedin.com/in/nikolaos-apostolidis-6363a0216/", "email": None, "domain": "dustandcream.gr"}, {"name": "Amalia Charalampopoulou", "latest_experience_title": "Brand Content Manager", "url": "https://www.linkedin.com/in/amalia-charalampopoulou-b258951a4/", "email": None, "domain": "dustandcream.gr"}], "dustandcream.gr"),
    "Froika": ([{"name": "Kelly Kottari", "latest_experience_title": "Marketing Coordinator", "url": "https://www.linkedin.com/in/kellykottari/", "email": "kkottari@froika.com", "domain": "froika.com"}], "froika.com"),
    "Helenvita": ([{"name": "Ioanna G.", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/ioanna-grigoriou-28159b29/", "email": "i.grigoriou@pharmex.gr", "domain": "pharmex.gr"}], "pharmex.gr"),
    "Hugo Boss": ([{"name": "Katerina Antoniou", "latest_experience_title": "Marketing and Communications", "url": "https://www.linkedin.com/in/katerina-antoniou-412694b2/", "email": "katerina_antoniou@hugoboss.com", "domain": "hugoboss.com"}, {"name": "Dr. Argyro Tsampis", "latest_experience_title": "Brand Ambassador", "url": "https://www.linkedin.com/in/dr-argyro-tsampis-89021378/", "email": None, "domain": "hugoboss.com"}], "hugoboss.com"),
    "Juliette Armand": ([{"name": "Elpida Fardogianni", "latest_experience_title": "Marketing Specialist", "url": "https://www.linkedin.com/in/elpida-fardogianni-124877b2/", "email": None, "domain": "juliettearmand.com"}, {"name": "Bill Papaefstratiou", "latest_experience_title": "Head of Marketing", "url": "https://www.linkedin.com/in/bill-papaefstratiou-26a4b258/", "email": "william@juliettearmand.com", "domain": "juliettearmand.com"}], "juliettearmand.com"),
    "Little Secrets Natural Cosmetics": ([{"name": "Eirini Amygdalia Petridou", "latest_experience_title": "Marketing Assistant", "url": "https://www.linkedin.com/in/eirini-amygdalia-petridou-596122297/", "email": None, "domain": "littlesecrets.gr"}], "littlesecrets.gr"),
    "MAC Cosmetics": ([{"name": "Eirini Nikolakea", "latest_experience_title": "Product Marketing Executive", "url": "https://www.linkedin.com/in/eirini-nikolakea/", "email": "enikolakea@maccosmetics.com", "domain": "maccosmetics.com"}, {"name": "Eleni Charitopoulou", "latest_experience_title": "Digital Marketing Specialist", "url": "https://www.linkedin.com/in/eleni-charitopoulou-b8b617162/", "email": None, "domain": "maccosmetics.com"}], "maccosmetics.com"),
    "Mastic Spa": ([{"name": "Apostle Mengoulis", "latest_experience_title": "Digital Marketing Manager", "url": "https://www.linkedin.com/in/apostle-mengoulis-9947519b/", "email": "apostle@sodisbrands.com", "domain": "masticspa.com"}, {"name": "Sofia Sodi", "latest_experience_title": "Co-founder & Marketing Director", "url": "https://www.linkedin.com/in/sofia-sodi-6aaa94a6/", "email": "sofia@sodisbrands.com", "domain": "masticspa.com"}], "masticspa.com"),
    "Medisei": ([{"name": "Popita Bikof", "latest_experience_title": "Head of Marketing@MEDISEI", "url": "https://www.linkedin.com/in/popita-bikof-2199aa50/", "email": "p.bikof@medisei.gr", "domain": "medisei.gr"}, {"name": "Stella Mathioudaki", "latest_experience_title": "Digital Marketing & E-Commerce Specialist", "url": "https://www.linkedin.com/in/stella-mathioudaki/", "email": "s.mathioudaki@medisei.gr", "domain": "medisei.gr"}, {"name": "Αμαλία Λυμπέρη", "latest_experience_title": "Marketing Assistant Trainee", "url": "https://www.linkedin.com/in/αμαλία-λυμπέρη-a30387396/", "email": None, "domain": "medisei.gr"}], "medisei.gr"),
    "Panthenol Extra": ([{"name": "Popita Bikof", "latest_experience_title": "Head of Marketing@MEDISEI", "url": "https://www.linkedin.com/in/popita-bikof-2199aa50/", "email": "p.bikof@medisei.gr", "domain": "medisei.gr"}, {"name": "Stella Mathioudaki", "latest_experience_title": "Digital Marketing & E-Commerce Specialist", "url": "https://www.linkedin.com/in/stella-mathioudaki/", "email": "s.mathioudaki@medisei.gr", "domain": "medisei.gr"}, {"name": "Αμαλία Λυμπέρη", "latest_experience_title": "Marketing Assistant Trainee", "url": "https://www.linkedin.com/in/αμαλία-λυμπέρη-a30387396/", "email": None, "domain": "medisei.gr"}], "medisei.gr"),
    "Mon Rêve Cosmetics": ([{"name": "Lydia Christoforidou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/lydia-christoforidou-9278a4122/", "email": None, "domain": "hellenica.gr"}, {"name": "Electra Papanastasiou", "latest_experience_title": "Marketing Manager, Strategy Lead Seventeen Cosmetics", "url": "https://www.linkedin.com/in/papanastasiou-electra/", "email": "e.papanastasiou@hellenica.gr", "domain": "hellenica.gr"}, {"name": "Silia Karatza", "latest_experience_title": "Digital Marketing & Social Media Specialist", "url": "https://www.linkedin.com/in/silia-karatza-79b726197/", "email": "s.karatza@hellenica.gr", "domain": "hellenica.gr"}], "hellenica.gr"),
    "Radiant Professional": ([{"name": "Lydia Christoforidou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/lydia-christoforidou-9278a4122/", "email": None, "domain": "hellenica.gr"}, {"name": "Electra Papanastasiou", "latest_experience_title": "Marketing Manager, Strategy Lead Seventeen Cosmetics", "url": "https://www.linkedin.com/in/papanastasiou-electra/", "email": "e.papanastasiou@hellenica.gr", "domain": "hellenica.gr"}, {"name": "Silia Karatza", "latest_experience_title": "Digital Marketing & Social Media Specialist", "url": "https://www.linkedin.com/in/silia-karatza-79b726197/", "email": "s.karatza@hellenica.gr", "domain": "hellenica.gr"}], "hellenica.gr"),
    "Radiant Professional Make Up": ([{"name": "Lydia Christoforidou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/lydia-christoforidou-9278a4122/", "email": None, "domain": "hellenica.gr"}, {"name": "Electra Papanastasiou", "latest_experience_title": "Marketing Manager, Strategy Lead Seventeen Cosmetics", "url": "https://www.linkedin.com/in/papanastasiou-electra/", "email": "e.papanastasiou@hellenica.gr", "domain": "hellenica.gr"}, {"name": "Silia Karatza", "latest_experience_title": "Digital Marketing & Social Media Specialist", "url": "https://www.linkedin.com/in/silia-karatza-79b726197/", "email": "s.karatza@hellenica.gr", "domain": "hellenica.gr"}], "hellenica.gr"),
    "Oriflame": ([{"name": "despoina noukari", "latest_experience_title": "oriflame marketing", "url": "https://www.linkedin.com/in/despoina-noukari-5a0b9aab/", "email": None, "domain": "oriflame.com"}], "oriflame.com"),
    "Priorin": ([{"name": "Sonia Mousavere", "latest_experience_title": "Head of Communications & PGA", "url": "https://www.linkedin.com/in/smousavere/", "email": "sonia.mousavere@bayer.com", "domain": "bayer.com"}, {"name": "Thekli Bourtzinou", "latest_experience_title": "Brand Communications and Digital Services Coordinator", "url": "https://www.linkedin.com/in/thekli-bourtzinou-340b3099/", "email": "thekli.bourtzinou@bayer.com", "domain": "bayer.com"}, {"name": "Kyriakos Nathanail", "latest_experience_title": "Brand Manager / Customer Engagement Team Lead", "url": "https://www.linkedin.com/in/kyriakos-nathanail-7118874/", "email": "kyriakos.nathanail@bayer.com", "domain": "bayer.com"}, {"name": "Fotis Tsopelas", "latest_experience_title": "Digital Media & eCommerce Manager Consumer Health Division", "url": "https://www.linkedin.com/in/fotis-tsopelas-41b72960/", "email": "fotis.tsopelas@bayer.com", "domain": "bayer.com"}, {"name": "Spiros Galousis", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/spiros-galousis-24781746/", "email": "sgalousis@bayer.com", "domain": "bayer.com"}, {"name": "Maria Davit", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/maria-davit-110006115/", "email": "maria.davit@bayer.com", "domain": "bayer.com"}, {"name": "Zoi Arachovitou", "latest_experience_title": "Marketing Assistant", "url": "https://www.linkedin.com/in/zoi-arachovitou/", "email": "zoi.arachovitou@bayer.com", "domain": "bayer.com"}, {"name": "Eleftherios Filios", "latest_experience_title": "Cluster Brand Manager Opthalmology (GR, CY, MA, IS, ROM, BG)", "url": "https://www.linkedin.com/in/eleftherios-filios-0959b35a/", "email": None, "domain": "bayer.com"}], "bayer.com"),
    "The Body Shop": ([{"name": "Dimitra Kaldi", "latest_experience_title": "Marketing, Communication & CSR Director at The Body Shop", "url": "https://www.linkedin.com/in/dimitra-kaldi-a8a36350/", "email": "dimitra.kaldi@thebodyshop.com", "domain": "thebodyshopcareers.com"}, {"name": "Eleni Vasilakopoulou", "latest_experience_title": "Marketing Coordinator", "url": "https://www.linkedin.com/in/eleni-vasilakopoulou-87ab4642/", "email": "eleni.vasilakopoulou@thebodyshop.com", "domain": "thebodyshopcareers.com"}], "thebodyshopcareers.com"),
    "Version Derm": ([{"name": "Chryssa Siaga", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/chryssa-siaga-37252484/", "email": "chryssa@mail.versionderm.gr", "domain": "versionderm.gr"}], "versionderm.gr"),
    "AEK FC": ([{"name": "Antonis Apostolopoulos", "latest_experience_title": "Event Management & Sports Marketing Assistant", "url": "https://www.linkedin.com/in/antonis-apostolopoulos-7b753b187/", "email": "aapostolopoulos@aekfc.gr", "domain": "aekfc.gr"}, {"name": "Christina Koromila", "latest_experience_title": "Social Media", "url": "https://www.linkedin.com/in/christina-koromila-768326b5/", "email": "c.koromila@kingbetmedia.com", "domain": "aekfc.gr"}], "aekfc.gr"),
    "ANT1": ([{"name": "Marco Struecker", "latest_experience_title": "Interim General Manager - ANT1+", "url": "https://www.linkedin.com/in/marcostruecker/", "email": None, "domain": "ant1.gr"}, {"name": "Konstantinos Bourounis", "latest_experience_title": "Chief Marketing Officer - Antenna Group Greece", "url": "https://www.linkedin.com/in/konstantinos-bourounis-8675a0/", "email": None, "domain": "ant1.gr"}, {"name": "Agapi Kantartzi", "latest_experience_title": "Marketing Manager & Communications | Antenna Audio / easy 97.2 | Ρυθμος 94.9 | Soundis.gr", "url": "https://www.linkedin.com/in/agapi/", "email": None, "domain": "ant1.gr"}, {"name": "Nikolaos Katsaros", "latest_experience_title": "Digital Performance Product Manager", "url": "https://www.linkedin.com/in/nikolaos-katsaros-52b9a11b/", "email": None, "domain": "ant1.gr"}, {"name": "Ioanna Panagioti", "latest_experience_title": "Social Media Expert", "url": "https://www.linkedin.com/in/ioanna-panagioti-25bba91a9/", "email": None, "domain": "ant1.gr"}, {"name": "Nancy Fafouti", "latest_experience_title": "Social Media Manager", "url": "https://www.linkedin.com/in/nancy-fafouti/", "email": None, "domain": "ant1.gr"}, {"name": "Olympia Tsamasfyra", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/olympia-tsamasfyra-40033069/", "email": None, "domain": "ant1.gr"}, {"name": "Iraklis Ioannidis", "latest_experience_title": "Music Marketing Lead - Marketing Manager Heaven/Warner & Antenna Intelligence", "url": "https://www.linkedin.com/in/iraklis-ioannidis-0a536622/", "email": None, "domain": "ant1.gr"}, {"name": "Konstantinos Tzouros", "latest_experience_title": "Digital technical manager", "url": "https://www.linkedin.com/in/konstantinos-tzouros-5b912025/", "email": None, "domain": "ant1.gr"}, {"name": "Angelos Chatzidakis", "latest_experience_title": "Marketing Executive", "url": "https://www.linkedin.com/in/angelos-chatzidakis/", "email": None, "domain": "ant1.gr"}, {"name": "Marianna Alexiou", "latest_experience_title": "Social Media Expert", "url": "https://www.linkedin.com/in/marianna-alexiou-5287bb13a/", "email": None, "domain": "ant1.gr"}, {"name": "ALIKI ROULIA", "latest_experience_title": "Marketing Supervisor", "url": "https://www.linkedin.com/in/aliki-roulia-821a8097/", "email": None, "domain": "ant1.gr"}, {"name": "Matina Fountzoula", "latest_experience_title": "Digital Content Editor", "url": "https://www.linkedin.com/in/matina-fountzoula-6b215725/", "email": None, "domain": "ant1.gr"}, {"name": "Stavros Papazoglou", "latest_experience_title": "Social Media Coordinator", "url": "https://www.linkedin.com/in/stavpapazoglou/", "email": None, "domain": "ant1.gr"}, {"name": "Diacopoulos Fragiskos", "latest_experience_title": "marketing and tv production manager", "url": "https://www.linkedin.com/in/diacopoulos-fragiskos-7856b618/", "email": None, "domain": "ant1.gr"}], "ant1.gr"),
    "Box Now": ([{"name": "Chris Papandropoulos", "latest_experience_title": "Group CMO", "url": "https://www.linkedin.com/in/chrispapandropoulos/", "email": None, "domain": "boxnow.gr"}, {"name": "Nikolaos Katsadramis", "latest_experience_title": "Marketing Executive", "url": "https://www.linkedin.com/in/nikolaos-katsadramis-185881129/", "email": None, "domain": "boxnow.gr"}, {"name": "Anastasia Kalliaropoulou", "latest_experience_title": "Marketing Supervisor", "url": "https://www.linkedin.com/in/anastasia-kalliaropoulou-6572b2194/", "email": "anastasia.kalliaropoulou@boxnow.gr", "domain": "boxnow.gr"}], "boxnow.gr"),
    "MUA Makeup Academy": ([
        {"name": "Panagiotis Velonas", "latest_experience_title": "Founder | BVELO", "url": "https://www.linkedin.com/in/panagiotis-velonas-886254353/", "email": "pvelonas@bvelo.gr", "domain": "muagreece.com"},
        {"name": "Julia Velona", "latest_experience_title": "Chief Operating Officer", "url": "https://www.linkedin.com/in/julia-velona-a586b9291/", "email": "julia@bvelo.gr", "domain": "muagreece.com"},
        {"name": "Maria Oikonomou", "latest_experience_title": "Digital Marketing & Performance Strategist", "url": "https://www.linkedin.com/in/maria-oikonomou-28938a1aa/", "email": "maria.oikonomou@bravescale.gr", "domain": "muagreece.com"},
        {"name": "Angeliki Gidiotou", "latest_experience_title": "Influencer Marketing and UGC Specialist", "url": "https://www.linkedin.com/in/angeliki-gidiotou/", "email": None, "domain": "muagreece.com"},
        {"name": "Efi Konstantinidou", "latest_experience_title": "Senior Graphic Designer", "url": "https://www.linkedin.com/in/efi-konstantinidou-375b92135/", "email": None, "domain": "muagreece.com"},
        {"name": "Evdokia Gousiou", "latest_experience_title": "B2B Specialist", "url": "https://www.linkedin.com/in/evdokia-gousiou-636614299/", "email": "egousiou@bvelo.gr", "domain": "muagreece.com"},
        {"name": "Anastasis Tsimpouktsoglou", "latest_experience_title": "Shopify Tech & Performance Specialist", "url": "https://www.linkedin.com/in/anastasis-tsimpouktsoglou-9b327a2b9/", "email": None, "domain": "muagreece.com"},
    ], "muagreece.com"),
    "Myikona": ([{"name": "Garyfallos Ntalampiras", "latest_experience_title": "Director of E-commerce, Marketing & Growth", "url": "https://www.linkedin.com/in/garyfallos-ntalampiras-/", "email": "g.ntalampiras@myikona.gr", "domain": "myikona.gr"}], "myikona.gr"),
    "Bvelo": ([
        {"name": "Panagiotis Velonas", "latest_experience_title": "Founder | BVELO", "url": "https://www.linkedin.com/in/panagiotis-velonas-886254353/", "email": "pvelonas@bvelo.gr", "domain": "bvelo.gr"},
        {"name": "Julia Velona", "latest_experience_title": "Chief Operating Officer", "url": "https://www.linkedin.com/in/julia-velona-a586b9291/", "email": "julia@bvelo.gr", "domain": "bvelo.gr"},
        {"name": "Maria Oikonomou", "latest_experience_title": "Digital Marketing & Performance Strategist", "url": "https://www.linkedin.com/in/maria-oikonomou-28938a1aa/", "email": "maria.oikonomou@bravescale.gr", "domain": "bvelo.gr"},
        {"name": "Angeliki Gidiotou", "latest_experience_title": "Influencer Marketing and UGC Specialist", "url": "https://www.linkedin.com/in/angeliki-gidiotou/", "email": None, "domain": "bvelo.gr"},
        {"name": "Efi Konstantinidou", "latest_experience_title": "Senior Graphic Designer", "url": "https://www.linkedin.com/in/efi-konstantinidou-375b92135/", "email": None, "domain": "bvelo.gr"},
        {"name": "Evdokia Gousiou", "latest_experience_title": "B2B Specialist", "url": "https://www.linkedin.com/in/evdokia-gousiou-636614299/", "email": "egousiou@bvelo.gr", "domain": "bvelo.gr"},
        {"name": "Anastasis Tsimpouktsoglou", "latest_experience_title": "Shopify Tech & Performance Specialist", "url": "https://www.linkedin.com/in/anastasis-tsimpouktsoglou-9b327a2b9/", "email": None, "domain": "bvelo.gr"},
    ], "bvelo.gr"),
}
BATCH_ARCHIVED_2 = {
    "Muagreece": ([
        {"name": "Angeliki Gidiotou", "latest_experience_title": "Influencer Marketing and UGC Specialist", "url": "https://www.linkedin.com/in/angeliki-gidiotou/", "email": "", "domain": "muagreece.com"},
        {"name": "Maria Oikonomou", "latest_experience_title": "Digital Marketing & Performance Strategist", "url": "https://www.linkedin.com/in/maria-oikonomou-28938a1aa/", "email": "", "domain": "muagreece.com"},
        {"name": "Efi Konstantinidou", "latest_experience_title": "Senior Graphic Designer", "url": "https://www.linkedin.com/in/efi-konstantinidou-375b92135/", "email": "", "domain": "muagreece.com"},
        {"name": "Julia Velona", "latest_experience_title": "Chief Operating Officer", "url": "https://www.linkedin.com/in/julia-velona-a586b9291/", "email": "", "domain": "muagreece.com"},
        {"name": "Panagiotis Velonas", "latest_experience_title": "Founder | BVELO", "url": "https://www.linkedin.com/in/panagiotis-velonas-886254353/", "email": "", "domain": "muagreece.com"},
        {"name": "Anna Dovliatidou", "latest_experience_title": "Professional Makeup Artist", "url": "https://www.linkedin.com/in/anna-dovliatidou-94b290126/", "email": "", "domain": "muagreece.com"},
        {"name": "Evdokia Gousiou", "latest_experience_title": "B2B Specialist", "url": "https://www.linkedin.com/in/evdokia-gousiou-636614299/", "email": "", "domain": "muagreece.com"},
        {"name": "Anastasis Tsimpouktsoglou", "latest_experience_title": "Shopify Tech & Performance Specialist", "url": "https://www.linkedin.com/in/anastasis-tsimpouktsoglou-9b327a2b9/", "email": "", "domain": "muagreece.com"},
    ], "muagreece.com"),
}
BATCH_ARCHIVED = {
    "adidas": ([
        {"name": "Fay Petroulaki", "latest_experience_title": "Director Brand Comms, Newsroom & Publishing Southeast Europe", "url": "https://www.linkedin.com/in/fay-petroulaki-2871b592/", "email": "", "domain": "adidas.com"},
        {"name": "Theodore Michopoulos", "latest_experience_title": "Senior Manager, Human Resources Operations, South East Europe and Italy", "url": "https://www.linkedin.com/in/theodore-michopoulos-27489616/", "email": "", "domain": "adidas.com"},
        {"name": "Georgia Koutsoukou", "latest_experience_title": "Wholesale Activation Manager, SEE", "url": "https://www.linkedin.com/in/georgia-koutsoukou-87179393/", "email": "", "domain": "adidas.com"},
        {"name": "Christina Panagiotou", "latest_experience_title": "Senior Manager, Internal Communications - South Europe", "url": "https://www.linkedin.com/in/christinapanagiotou/", "email": "", "domain": "adidas.com"},
        {"name": "George Sifnios", "latest_experience_title": "Country Manager South East Europe", "url": "https://www.linkedin.com/in/george-sifnios-369b331/", "email": "", "domain": "adidas.com"},
        {"name": "Anna Kechaidou", "latest_experience_title": "Brand Communications & PR Country Lead", "url": "https://www.linkedin.com/in/anna-kechaidou-0a478b41/", "email": "", "domain": "adidas.com"},
        {"name": "Konstantinos Koutroumpis", "latest_experience_title": "Digital Performance Manager JR - SEE", "url": "https://www.linkedin.com/in/kostas-koutroumpis/", "email": "", "domain": "adidas.com"},
        {"name": "Konstantinos Konstantinidis", "latest_experience_title": "Senior Manager Sports Marketing SEE", "url": "https://www.linkedin.com/in/konstantinos-konstantinidis-1120125b/", "email": "", "domain": "adidas.com"},
        {"name": "Angeliki Prezani, MBA", "latest_experience_title": "Jr. Manager, Wholesale Activation, South East Europe", "url": "https://www.linkedin.com/in/angeliki-prezani/", "email": "", "domain": "adidas.com"},
        {"name": "Theodoros Gkolfinopoulos", "latest_experience_title": "Senior Manager Digital Sales, SEE", "url": "https://www.linkedin.com/in/theodoros-gkolfinopoulos-95970633/", "email": "", "domain": "adidas.com"},
        {"name": "Eirini Gazidelli", "latest_experience_title": "Manager HR Operations & Payroll South East Europe", "url": "https://www.linkedin.com/in/eirini-gazidelli-19b36241/", "email": "", "domain": "adidas.com"},
        {"name": "Alex Zourelidis", "latest_experience_title": "Senior Manager, Visual Merchandising, Retail Space Management & Product Learning Southeast Europe", "url": "https://www.linkedin.com/in/alex-zourelidis/", "email": "", "domain": "adidas.com"},
        {"name": "Konstantinos Christodoulakis", "latest_experience_title": "Omnichannel Director, South East Europe", "url": "https://www.linkedin.com/in/konstantinos-christodoulakis/", "email": "", "domain": "adidas.com"},
        {"name": "Mina Gkasti", "latest_experience_title": "HR Director, South East Europe", "url": "https://www.linkedin.com/in/mina-gkasti-22274b33/", "email": "", "domain": "adidas.com"},
        {"name": "Alexandros E. Bontikoulis", "latest_experience_title": "Director GTM", "url": "https://www.linkedin.com/in/alexandros-e-bontikoulis-5730981b/", "email": "", "domain": "adidas.com"},
        {"name": "Efstathios Spinos", "latest_experience_title": "Director Marketplace Development", "url": "https://www.linkedin.com/in/efstathios-spinos-a60a0962/", "email": "", "domain": "adidas.com"},
        {"name": "Aggelos Kampanakis", "latest_experience_title": "Sr Manager Customer Fulfillment SEE", "url": "https://www.linkedin.com/in/kampanakisang/", "email": "", "domain": "adidas.com"},
        {"name": "Agis Skarvelis", "latest_experience_title": "Senior Go To Market Manager South East Europe", "url": "https://www.linkedin.com/in/agamemnonskarvelis/", "email": "", "domain": "adidas.com"},
        {"name": "George Generalis", "latest_experience_title": "Senior Director Customer Service, South Europe", "url": "https://www.linkedin.com/in/george-generalis-92637517/", "email": "", "domain": "adidas.com"},
    ], "adidas.com"),
    "Sephora": ([
        {"name": "Marilia Tompra", "latest_experience_title": "Country Director & Board Member SEPHORA Greece, LVMH group", "url": "https://www.linkedin.com/in/marilia-tompra-05342517/", "email": "", "domain": "sephora.com"},
        {"name": "Afroditi Kiousi", "latest_experience_title": "Retail HR Manager", "url": "https://www.linkedin.com/in/afroditi-kiousi-947703144/", "email": "", "domain": "sephora.com"},
        {"name": "Maria Delagrammatika", "latest_experience_title": "Retail Operations Director Greece & Balkans (Croatia, Serbia, Romania, Bulgaria)", "url": "https://www.linkedin.com/in/maroulio-delagrammatika/", "email": "", "domain": "sephora.com"},
        {"name": "Eirini Iοannou", "latest_experience_title": "Head of Ecommerce & Growth", "url": "https://www.linkedin.com/in/eirini-i%ce%bfannou-77840998/", "email": "", "domain": "sephora.com"},
        {"name": "Maria Simnianaki", "latest_experience_title": "PR & Social Media Specialist", "url": "https://www.linkedin.com/in/maria-simnianaki-29288b181/", "email": "", "domain": "sephora.com"},
        {"name": "Maria Sotiriou", "latest_experience_title": "Head of Commercial Strategy & Selective Market", "url": "https://www.linkedin.com/in/maria-sotiriou-81797b30/", "email": "", "domain": "sephora.com"},
        {"name": "Evangelos Bakeas, MBA, CFA, Dipl", "latest_experience_title": "CFO & Supply Director Greece", "url": "https://www.linkedin.com/in/evangelos-bakeas-mba-cfa-dipl-b611754a/", "email": "", "domain": "sephora.com"},
        {"name": "Maria Sideri", "latest_experience_title": "Head of IT Store Efficiency Sephora EME", "url": "https://www.linkedin.com/in/maria-sideri-a8a89228/", "email": "", "domain": "sephora.com"},
        {"name": "Vasilios Exarchos", "latest_experience_title": "Finance Manager", "url": "https://www.linkedin.com/in/vasilios-exarchos-157a7721/", "email": "", "domain": "sephora.com"},
        {"name": "Marianna Papazisi", "latest_experience_title": "Loyalty CRM Specialist", "url": "https://www.linkedin.com/in/marianna-papazisi-a6445b100/", "email": "", "domain": "sephora.com"},
        {"name": "Dimitris Karampelas", "latest_experience_title": "Store Manager", "url": "https://www.linkedin.com/in/dimitris-karampelas-426a12111/", "email": "", "domain": "sephora.com"},
        {"name": "Karmiris Panagiotis", "latest_experience_title": "Store Manager", "url": "https://www.linkedin.com/in/karmiris-panagiotis-89756274/", "email": "", "domain": "sephora.com"},
        {"name": "Antonia Sofianidou", "latest_experience_title": "Media & Advertising Specialist", "url": "https://www.linkedin.com/in/antonia-sofianidou-83157b17a/", "email": "", "domain": "sephora.com"},
        {"name": "Kassandra Megarioti", "latest_experience_title": "Training Specialist", "url": "https://www.linkedin.com/in/kassandra-megarioti-88603629a/", "email": "", "domain": "sephora.com"},
        {"name": "Frosso Stratakou", "latest_experience_title": "Accounting Specialist", "url": "https://www.linkedin.com/in/frosso-stratakou-1492a2228/", "email": "", "domain": "sephora.com"},
        {"name": "Elena Retzou", "latest_experience_title": "Category Specialist", "url": "https://www.linkedin.com/in/elena-retzou-1ab414b1/", "email": "", "domain": "sephora.com"},
        {"name": "Dimitris Karelos", "latest_experience_title": "Marketing Trainee", "url": "https://www.linkedin.com/in/dimitris-karelos-296b68225/", "email": "", "domain": "sephora.com"},
        {"name": "Thodoris Lalaounis", "latest_experience_title": "Estore Catalog & Commercial Activation Specialist", "url": "https://www.linkedin.com/in/thodoris-lalaounis-7476491b4/", "email": "", "domain": "sephora.com"},
        {"name": "Thelia Chrisomalli", "latest_experience_title": "Store Manager", "url": "https://www.linkedin.com/in/thelia-chrisomalli-365309196/", "email": "", "domain": "sephora.com"},
    ], "sephora.com"),
}


if __name__ == "__main__":
    if "git_sync" in sys.argv:
        git_sync()
    elif "save" in sys.argv:
        if not BATCH:
            print("BATCH is empty — nothing to save.")
        else:
            for brand, (contacts, domain) in BATCH.items():
                save_contacts(brand, contacts, domain)
            print("Done.")
    elif "enrich_emails" in sys.argv:
        idx = sys.argv.index("enrich_emails")
        brands = sys.argv[idx + 1:]
        if not brands:
            print("Usage: python discover_contacts.py enrich_emails \"Brand1\" [\"Brand2\" ...]")
        else:
            list_for_enrichment(brands)
    elif "purge" in sys.argv:
        purge_banned()
    elif "export" in sys.argv:
        export_excel()
    elif "sync_sheets" in sys.argv:
        sync_to_sheets()
    elif "mark_empty" in sys.argv:
        idx = sys.argv.index("mark_empty")
        brands = sys.argv[idx + 1:]
        if not brands:
            print("Usage: python discover_contacts.py mark_empty \"Brand1\" \"Brand2\" ...")
        else:
            for b in brands:
                zz = STORE_DIR / f"ZZ-{b}.json"
                zz.write_text("[]", encoding="utf-8")
                old = STORE_DIR / f"00-{b}.json"
                if old.exists():
                    old.unlink()
                _sync_actionable(b)
                print(f"  '{b}' → ZZ-{b}.json (marked empty)")
    elif "unmark_empty" in sys.argv:
        idx = sys.argv.index("unmark_empty")
        brands = sys.argv[idx + 1:]
        if not brands:
            print("Usage: python discover_contacts.py unmark_empty \"Brand1\" ...")
        else:
            for b in brands:
                zz = STORE_DIR / f"ZZ-{b}.json"
                if zz.exists():
                    zz.unlink()
                    _sync_actionable(b)
                    print(f"  '{b}' removed from empty — will appear in next batch")
                else:
                    print(f"  '{b}' was not marked empty")
    else:
        status_report()

