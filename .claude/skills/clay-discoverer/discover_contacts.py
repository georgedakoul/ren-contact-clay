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
    "Muagreece":         "muagreece.com",   # parent co BVELO operates MUA Greece brand
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
BATCH = {}
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

