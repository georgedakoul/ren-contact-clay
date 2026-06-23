"""discover_contacts.py — Clay contact discovery driven by actionable_contacts.json.

Usage:
  python discover_contacts.py          → print status report
  python discover_contacts.py save     → run save logic for the BATCH dict below

Claude fills the BATCH dict after each round of Clay MCP calls, then runs:
  python discover_contacts.py save
"""
import json, sys, unicodedata, re
from datetime import datetime, timezone
from pathlib import Path

ROOT        = Path(__file__).resolve().parents[3]
STORE_DIR   = ROOT / "AI Sales Agent System" / "output" / "employees"
CONTACTS_FILE = ROOT / "AI Sales Agent System" / "actionable_contacts.json"
CLAY_STATE_FILE = ROOT / "AI Sales Agent System" / "output" / "clay_search_state.json"
TODAY       = datetime.now(timezone.utc).strftime("%Y-%m-%d")
STORE_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Identifier overrides: brands where linkedin_slug resolves to the wrong company
# or where domain gives better results than the LinkedIn company page.
# ---------------------------------------------------------------------------
IDENTIFIER_OVERRIDES = {
    "COSMOTE":           "cosmote.gr",          # slug → wrong company
    "Village Cinemas":   "villagecinemas.gr",    # slug → Australian entity
    "Cosmos Sport":      "cosmossport.gr",       # slug had issues
    "Coca-Cola":         "coca-colahellenic.com",
    "Psichogios Books":  "psichogios.gr",
    "Pame Stoixima":     "opap.gr",             # brand lives under OPAP
    "ION":               "ion.gr",              # slug returns wrong/0 results
    "instacar":          "instacar.gr",         # slug returns US company
    "Apivita":           "apivita.com",         # slug returns 0 results
    "more.com":          "wind.com.gr",         # slug → wrong entity; Nova/WIND domain
    "Alterlife":         "alterlife.gr",        # slug returned 0 emails; domain works
    "SKY express":       "skyexpress.gr",       # slug returned 0 emails; domain works
    "Fresh Line":        "freshline.gr",        # slug returned 0 emails; domain works
    "Alumil":            "alumil.com",          # slug returned 0 emails; domain works
    "BSB Fashion":       "bsbfashion.com",      # slug returned 0 emails; domain works
    "La Vie en Rose":    "lavieenrose.com",     # slug → Swiss NGO; use domain + Greece filter
    "Mind Your Style":   "mindyourstyle.gr",    # slug returned 0 emails; domain works
    "Protergia":         "protergia.gr",        # slug returned 0 emails; domain works
}

# Global brands that need locations=["Greece"] filter to avoid non-GR employees
GLOBAL_BRANDS = {
    "Samsung", "Apple", "Herbalife", "IKEA", "LEGO", "Huawei",
    "Motorola", "Starbucks", "Red Bull", "Wolt", "FREENOW",
    "Hertz", "Sony Music Entertainment", "Chanel", "Visa",
    "Nespresso", "Monster Energy", "La Vie en Rose",
}


def _load_state():
    if CLAY_STATE_FILE.exists():
        return json.loads(CLAY_STATE_FILE.read_text(encoding="utf-8"))
    return {"empty": {}}

def _save_state(state):
    CLAY_STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")

def _mark_brand_empty(brand_name):
    state = _load_state()
    entry = state["empty"].get(brand_name, {"attempts": 0})
    entry["last_searched"] = TODAY
    entry["attempts"] = entry.get("attempts", 0) + 1
    state["empty"][brand_name] = entry
    _save_state(state)

def _unmark_brand_empty(brand_name):
    state = _load_state()
    if brand_name in state.get("empty", {}):
        del state["empty"][brand_name]
        _save_state(state)


def _strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def normalize(s):
    return re.sub(r"\s+", " ", _strip_accents(s).lower().strip())


def get_identifier(brand_name, linkedin_slug):
    if brand_name in IDENTIFIER_OVERRIDES:
        return IDENTIFIER_OVERRIDES[brand_name]
    if linkedin_slug:
        return f"https://www.linkedin.com/company/{linkedin_slug}"
    return None


def _employee_path(brand_name):
    """Return existing employee file; checks 00- prefix variant too."""
    p = STORE_DIR / f"{brand_name}.json"
    if p.exists():
        return p
    p0 = STORE_DIR / f"00-{brand_name}.json"
    if p0.exists():
        return p0
    return p  # default for new files


def save_contacts(brand_name, contacts, domain=None):
    if not contacts:
        _mark_brand_empty(brand_name)
        state = _load_state()
        attempts = state["empty"][brand_name]["attempts"]
        print(f"  {brand_name}: Clay returned 0 contacts → marked EMPTY (attempt #{attempts}, skipped in future batches)")
        return
    _unmark_brand_empty(brand_name)
    path = _employee_path(brand_name)
    existing = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    by_name = {normalize(e["name"]): e for e in existing}
    added = email_added = 0
    for c in contacts:
        name = (c.get("name") or "").strip()
        if not name:
            continue
        key = normalize(name)
        title  = (c.get("latest_experience_title") or "").strip()
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
    print(f"  {brand_name}: {added} new contacts, {email_added} emails added → {len(merged)} total")


def _scan_store():
    """Scan ALL employee files → {brand_name: {total, emails, last_seen}}."""
    store = {}
    for f in sorted(STORE_DIR.glob("*.json")):
        if f.name == "desktop.ini":
            continue
        brand_name = re.sub(r"^00-", "", f.stem)
        contacts = json.loads(f.read_text(encoding="utf-8"))
        email_count = sum(1 for c in contacts if c.get("email"))
        last_seen = max(
            (c.get("last_seen") or "1970-01-01") for c in contacts
        ) if contacts else None
        store[brand_name] = {"total": len(contacts), "emails": email_count, "last_seen": last_seen}
    return store


def status_report():
    # Pass 1: inventory the full employee store
    store = _scan_store()
    total_contacts = sum(v["total"] for v in store.values())
    total_emails   = sum(v["emails"] for v in store.values())
    brands_with_emails = sum(1 for v in store.values() if v["emails"] > 0)

    print(f"\n=== EMPLOYEE STORE OVERVIEW ===")
    print(f"  Files on disk : {len(store):>4}  ({brands_with_emails} with emails)")
    print(f"  Total contacts: {total_contacts:>4}  ({total_emails} with emails)")

    # Pass 2: load actionable_contacts.json and find searchable brands
    data      = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
    ac_brands = {b["brand"]: b for b in data["brands"]}

    searchable = {
        name: b for name, b in ac_brands.items()
        if b.get("linkedin_slug") or name in IDENTIFIER_OVERRIDES
    }

    # Categorise searchable brands against the store + state file
    empty_state = _load_state().get("empty", {})

    covered      = []   # have ≥1 email
    linkedin_only = []  # have contacts, 0 emails (00- prefix file)
    empty        = []   # Clay returned 0 contacts (tracked in state file)
    missing      = []   # never searched at all

    for name in sorted(searchable):
        slug       = searchable[name].get("linkedin_slug", "")
        identifier = get_identifier(name, slug)
        needs_loc  = name in GLOBAL_BRANDS
        info       = store.get(name)
        if info is not None and info["emails"] > 0:
            covered.append((name, identifier, info, needs_loc))
        elif info is not None:
            linkedin_only.append((name, identifier, info, needs_loc))
        elif name in empty_state:
            empty.append((name, identifier, needs_loc, empty_state[name]))
        else:
            missing.append((name, identifier, needs_loc))

    print(f"\n=== CLAY-SEARCHABLE BRANDS ({len(searchable)} total) ===")
    print(f"  ✓ Covered       (emails found)          : {len(covered)}")
    print(f"  ∅ LinkedIn-only (contacts, no emails)   : {len(linkedin_only)}")
    print(f"  ✗ Empty         (Clay returned 0, skip) : {len(empty)}")
    print(f"  ✗ Missing       (never searched)        : {len(missing)}")

    if covered:
        print(f"\n--- COVERED (have emails) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'#C':>4} {'#E':>4} {'Last seen':<12}")
        print("  " + "─" * 100)
        for name, identifier, info, needs_loc in covered:
            loc = "  [+GR]" if needs_loc else ""
            print(f"  {name:<32} {identifier:<45} {info['total']:>4} {info['emails']:>4} {info['last_seen'] or '—':<12}{loc}")

    if linkedin_only:
        print(f"\n--- LINKEDIN-ONLY (contacts found, 0 emails) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'#C':>4}  {'Last seen':<12}")
        print("  " + "─" * 96)
        for name, identifier, info, needs_loc in linkedin_only:
            loc = "  [+GR]" if needs_loc else ""
            zero_hint = "  ← 0 contacts, consider mark_empty" if info["total"] == 0 else ""
            print(f"  {name:<32} {identifier:<45} {info['total']:>4}  {info['last_seen'] or '—':<12}{loc}{zero_hint}")

    if empty:
        print(f"\n--- EMPTY (Clay returned 0 contacts — excluded from next batch) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'Tries':>5} {'Last searched':<12}")
        print("  " + "─" * 100)
        for name, identifier, needs_loc, einfo in empty:
            loc = "  [+GR]" if needs_loc else ""
            print(f"  {name:<32} {identifier:<45} {einfo.get('attempts',1):>5}  {einfo.get('last_searched','—'):<12}{loc}")

    if missing:
        print(f"\n--- MISSING (never searched) ---")
        print(f"  {'Brand':<32} {'Identifier':<45} {'Flags'}")
        print("  " + "─" * 88)
        for name, identifier, needs_loc in missing:
            loc = "[+GR]" if needs_loc else ""
            print(f"  {name:<32} {identifier:<45} {loc}")

    # Next batch: missing first, then linkedin_only — empty brands are NEVER included
    next_up = list(missing) + [(n, i, nl) for n, i, info, nl in linkedin_only]
    if next_up:
        print(f"\n=== NEXT BATCH RECOMMENDATION (top 10) ===")
        print(f"  (Empty brands are excluded — re-run mark_empty to re-enable)")
        print(f"  {'#':<4} {'Tier':<12} {'Brand':<32} {'Identifier':<45} {'Flags'}")
        print("  " + "─" * 100)
        for idx, (name, identifier, needs_loc) in enumerate(next_up[:10], 1):
            tier = "missing" if idx <= len(missing) else "linkedin-only"
            loc  = "[+GR]" if needs_loc else ""
            print(f"  [{idx:>2}] {tier:<12}  {name:<32} {identifier:<45} {loc}")
    else:
        print("\nAll searchable brands are covered or pending LinkedIn-only retry.")



# ---------------------------------------------------------------------------
# BATCH — populated by Claude after each round of Clay MCP calls.
# Format: brand_name → (contacts_list, domain_hint)
# contacts_list items: {"name", "latest_experience_title", "url", "email", "domain"}
# ---------------------------------------------------------------------------
BATCH = {}
_LAST_BATCH_2026_06_22 = {
    "Douleutaras": ([
        {"name": "Danai Dimitriou", "latest_experience_title": "Performance Marketing Specialist", "url": "https://www.linkedin.com/in/danai-dimitriou/", "email": None, "domain": "douleutaras.gr"},
        {"name": "Apostolos Chatziathanasiou", "latest_experience_title": "Performance Marketing Manager", "url": "https://www.linkedin.com/in/apostolos-chatziathanasiou-63a347180/", "email": None, "domain": "douleutaras.gr"},
        {"name": "Ermis Anastasopoulos", "latest_experience_title": "Communications Designer", "url": "https://www.linkedin.com/in/ermis-anastasopoulos/", "email": None, "domain": "douleutaras.gr"},
        {"name": "Maria Gkouti", "latest_experience_title": "Social Media Specialist", "url": "https://www.linkedin.com/in/maria-gkouti-8285b91ab/", "email": None, "domain": "douleutaras.gr"},
        {"name": "Maria Panagopoulou", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/maria-panagopoulou-344315131/", "email": None, "domain": "douleutaras.gr"},
        {"name": "Christina Nicolaou", "latest_experience_title": "Social Media Specialist", "url": "https://www.linkedin.com/in/christina-nicolaou-6565981ba/", "email": None, "domain": "douleutaras.gr"},
    ], "douleutaras.gr"),

    "Box Now": ([
        {"name": "Chris Papandropoulos", "latest_experience_title": "Group CMO", "url": "https://www.linkedin.com/in/chrispapandropoulos/", "email": None, "domain": "boxnow.gr"},
        {"name": "Nikolaos Katsadramis", "latest_experience_title": "Marketing Executive", "url": "https://www.linkedin.com/in/nikolaos-katsadramis-185881129/", "email": None, "domain": "boxnow.gr"},
        {"name": "Anastasia Kalliaropoulou", "latest_experience_title": "Marketing Supervisor", "url": "https://www.linkedin.com/in/anastasia-kalliaropoulou-6572b2194/", "email": None, "domain": "boxnow.gr"},
        {"name": "Antonis Mpalamos", "latest_experience_title": "Director of Partners Operations", "url": "https://www.linkedin.com/in/antonis-mpalamos-2195825b/", "email": None, "domain": "boxnow.gr"},
        {"name": "ALEXANDROS VAGIAS", "latest_experience_title": "Sales Director", "url": "https://www.linkedin.com/in/alexandros-vagias-4811b27b/", "email": None, "domain": "boxnow.gr"},
    ], "boxnow.gr"),

    "Germanos": ([
        {"name": "Thanasis Panagopoulos", "latest_experience_title": "Digital and Social Media Senior Specialist", "url": "https://www.linkedin.com/in/thanasispanagopoulos/", "email": None, "domain": "germanos.gr"},
        {"name": "Maria Pitsiou", "latest_experience_title": "Corporate Communications COSMOTE eValue & Germanos", "url": "https://www.linkedin.com/in/maria-pitsiou-97310ba8/", "email": None, "domain": "germanos.gr"},
        {"name": "Georgia Konstantopoulou", "latest_experience_title": "Section Manager Trade Marketing GERMANOS Shops / BU RETAIL OTE GROUP", "url": "https://www.linkedin.com/in/georgia-konstantopoulou-b15768161/", "email": None, "domain": "germanos.gr"},
        {"name": "George Evangelinos", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/george-evangelinos-b771a9152/", "email": None, "domain": "germanos.gr"},
        {"name": "Christos Liaros", "latest_experience_title": "Trade Marketing", "url": "https://www.linkedin.com/in/christos-liaros-78752435/", "email": None, "domain": "germanos.gr"},
        {"name": "Chrisanthos Mavromatis", "latest_experience_title": "Sales Marketing Manager", "url": "https://www.linkedin.com/in/chrisanthos-mavromatis-721220162/", "email": None, "domain": "germanos.gr"},
        {"name": "lefteris bamos", "latest_experience_title": "Mobile Marketing Manager", "url": "https://www.linkedin.com/in/lefteris-bamos-0816922bb/", "email": None, "domain": "germanos.gr"},
        {"name": "George Grammatikopoulos", "latest_experience_title": "Marketing Team Member", "url": "https://www.linkedin.com/in/george-grammatikopoulos-467b42289/", "email": None, "domain": "germanos.gr"},
        {"name": "Spyros Kafalis", "latest_experience_title": "Trade Marketing Manager", "url": "https://www.linkedin.com/in/spyros-kafalis-6059bb13/", "email": None, "domain": "germanos.gr"},
    ], "germanos.gr"),

    "MEVGAL": ([
        {"name": "GIOTA BETZOUNI", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/giota-betzouni-79355550/", "email": None, "domain": "mevgal.gr"},
        {"name": "Joanne Jouli", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/joanne-jouli-17738418/", "email": None, "domain": "mevgal.gr"},
        {"name": "Stavroula Koutsou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/stavroula-koutsou-05283211/", "email": None, "domain": "mevgal.gr"},
        {"name": "Stathis Manakos", "latest_experience_title": "Export Marketing Manager", "url": "https://www.linkedin.com/in/stathis-manakos-9a8a159a/", "email": None, "domain": "mevgal.gr"},
        {"name": "Stelios Malakopoulos", "latest_experience_title": "Brand Ambassador", "url": "https://www.linkedin.com/in/stelios-malakopoulos-3b8b1bb1/", "email": None, "domain": "mevgal.gr"},
        {"name": "George Vanidis", "latest_experience_title": "Exports Director", "url": "https://www.linkedin.com/in/george-vanidis-a974719/", "email": None, "domain": "mevgal.gr"},
        {"name": "Dimitrios Katsaras", "latest_experience_title": "Commercial director", "url": "https://www.linkedin.com/in/dimitrios-katsaras-95071698/", "email": None, "domain": "mevgal.gr"},
        {"name": "Konstantina Kokkinopoulou", "latest_experience_title": "Marketing Department", "url": "https://www.linkedin.com/in/konstantina-kokkinopoulou-aa693313b/", "email": None, "domain": "mevgal.gr"},
    ], "mevgal.gr"),

    "Vitex": ([
        {"name": "Eleni Souladaki", "latest_experience_title": "Trade Marketing Manager", "url": "https://www.linkedin.com/in/eleni-souladaki-15954b19/", "email": None, "domain": "vitex.gr"},
        {"name": "Vana Paraskevopoulou", "latest_experience_title": "Trade Marketing Specialist", "url": "https://www.linkedin.com/in/vana-paraskevopoulou-098a28204/", "email": None, "domain": "vitex.gr"},
        {"name": "Anastasis Nikolouzos", "latest_experience_title": "Trade Marketing Coordinator", "url": "https://www.linkedin.com/in/anastasis-nikolouzos-2943a52b3/", "email": None, "domain": "vitex.gr"},
        {"name": "Efthymios Koletsis", "latest_experience_title": "International Business Director", "url": "https://www.linkedin.com/in/efthymios-koletsis-527a805a/", "email": None, "domain": "vitex.gr"},
    ], "vitex.gr"),

    "Three Cents": ([
        {"name": "Alexander Sourmpatis", "latest_experience_title": "Global Brand Ambassador", "url": "https://www.linkedin.com/in/alexander-sourmpatis-4694811b/", "email": None, "domain": "threecents.com"},
        {"name": "George Bagos", "latest_experience_title": "Business Data Analyst", "url": "https://www.linkedin.com/in/george-bagos-2962216a/", "email": None, "domain": "threecents.com"},
        {"name": "Sophia Terzopoulou", "latest_experience_title": "Head of Marketing & Communications", "url": "https://www.linkedin.com/in/sophia-terzopoulou-a25a9340/", "email": None, "domain": "threecents.com"},
    ], "threecents.com"),

    "Snappi": ([
        {"name": "Dimitris Ganetsos", "latest_experience_title": "Head of Marketing | Snappi", "url": "https://www.linkedin.com/in/dimitris-ganetsos-43ab8226/", "email": None, "domain": "snappibank.com"},
        {"name": "Sofia Nefeli Spilioti", "latest_experience_title": "Senior Marketing Manager", "url": "https://www.linkedin.com/in/sofianefeli/", "email": None, "domain": "snappibank.com"},
        {"name": "Evita Vitsentzatou", "latest_experience_title": "Senior Marketing Manager", "url": "https://www.linkedin.com/in/evita-vitsentzatou-aaa95059/", "email": None, "domain": "snappibank.com"},
        {"name": "Piret Reinson", "latest_experience_title": "Head of Global Brand & Communications", "url": "https://www.linkedin.com/in/piret-reinson-03b12012/", "email": None, "domain": "snappibank.com"},
        {"name": "Marios Koumpas", "latest_experience_title": "Public Relations Manager", "url": "https://www.linkedin.com/in/marios-koumpas-07a8205a/", "email": None, "domain": "snappibank.com"},
        {"name": "Maria Mouzakiti", "latest_experience_title": "Head of Culture & Internal Communications", "url": "https://www.linkedin.com/in/mouzakiti/", "email": None, "domain": "snappibank.com"},
        {"name": "George Todoris", "latest_experience_title": "Digital Marketing Lead", "url": "https://www.linkedin.com/in/george-todoris-b2612029/", "email": None, "domain": "snappibank.com"},
        {"name": "Dimitra Tsigkri", "latest_experience_title": "Social Media & Content Coordinator", "url": "https://www.linkedin.com/in/dimitra-tsigkri/", "email": None, "domain": "snappibank.com"},
    ], "snappibank.com"),

    "LG": ([
        {"name": "John Mantas", "latest_experience_title": "Director, AirConditioning, Energy & B2B", "url": "https://www.linkedin.com/in/john-mantas-88b8091/", "email": None, "domain": "lg.com"},
        {"name": "Krisa Ylli", "latest_experience_title": "Marketing Supervisor (TV&Audio)", "url": "https://www.linkedin.com/in/krisa-ylli-b7b72a185/", "email": None, "domain": "lg.com"},
        {"name": "CHARALAMPOS NIKOLAOU", "latest_experience_title": "Digital Display & Project Implementation Specialist – B2B | LG Electronics (Greece)", "url": "https://www.linkedin.com/in/charalampos-nikolaou-9a248b191/", "email": None, "domain": "lg.com"},
    ], "lg.com"),
}

_LAST_BATCH = {
    "JYSK": ([
        {"name": "Kostas Dimopoulos DipWSET", "latest_experience_title": "Sales & Marketing Manager", "url": "https://www.linkedin.com/in/kostas-dimopoulos-dipwset-a2897320/", "email": None, "domain": "jysk.com"},
        {"name": "Eleni Kolotsiou", "latest_experience_title": "Commercial Sales and Marketing Assistant, Greece", "url": "https://www.linkedin.com/in/eleni-kolotsiou-6999682a/", "email": None, "domain": "jysk.com"},
        {"name": "Tonia Liadi", "latest_experience_title": "Jysk Influencer", "url": "https://www.linkedin.com/in/tonia-liadi-ba977b33b/", "email": None, "domain": "jysk.com"},
    ], "jysk.com"),

    "Kinder": ([
        {"name": "Elena Kosmatou", "latest_experience_title": "Brand Manager Kinder | GR/CY/MT", "url": "https://www.linkedin.com/in/elena-kosmatou-15556a11b/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Anna Kokkali", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/anna-kokkali-34497990/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Alexandros Tzagkarakis", "latest_experience_title": "Trade Marketing Manager", "url": "https://www.linkedin.com/in/alextzagkarakis/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Maria Pappa", "latest_experience_title": "Brand Manager | GCM", "url": "https://www.linkedin.com/in/maria-pappa-829085130/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Anthoula Flegka", "latest_experience_title": "Business Insights and Trade Marketing Specialist", "url": "https://www.linkedin.com/in/anthoula-flegka/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Dora Kontorouda", "latest_experience_title": "Country Marketing Manager - Nutella, Premium Chocolate, Biscuits, Candies", "url": "https://www.linkedin.com/in/dora-kontorouda-113b5b20/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Eftychia Dania", "latest_experience_title": "Brand Manager Kinder Greece, Cyprus, Malta", "url": "https://www.linkedin.com/in/eftychia-dania-9280448b/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "Marianna Masoura", "latest_experience_title": "Country Marketing Manager", "url": "https://www.linkedin.com/in/marianna-masoura-900b5220b/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "marina andronikou", "latest_experience_title": "marketing research manager", "url": "https://www.linkedin.com/in/marina-andronikou-42a4633a/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "SAKIS BOURATZIS", "latest_experience_title": "Marketing", "url": "https://www.linkedin.com/in/sakis-bouratzis-380568100/", "email": None, "domain": "ferrerocareers.com"},
        {"name": "simone spada", "latest_experience_title": "MARKETING MANAGER", "url": "https://www.linkedin.com/in/simone-spada-3154778/", "email": None, "domain": "ferrerocareers.com"},
    ], "ferrerocareers.com"),

    "NIVEA": ([
        {"name": "Eleni Krietsepi", "latest_experience_title": "Senior Brand Manager", "url": "https://www.linkedin.com/in/elenikrietsepi/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Arianna Gkiolia", "latest_experience_title": "Shopper and Customer Marketing Manager", "url": "https://www.linkedin.com/in/ariannagkiolia/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Despoina Katsira", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/despoina-katsira-60622615b/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Petros Chanis", "latest_experience_title": "Head of Shopper & Customer Marketing", "url": "https://www.linkedin.com/in/petros-chanis-a9796314/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Marlene Tabet", "latest_experience_title": "Head of Marketing", "url": "https://www.linkedin.com/in/marlene-tabet-57853725/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Lydia Karamaria", "latest_experience_title": "Junior Brand Manager", "url": "https://www.linkedin.com/in/lydiakaramaria/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Marilena Ramoutsaki", "latest_experience_title": "Digital Marketing Specialist", "url": "https://www.linkedin.com/in/marilena-ramoutsaki/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Alexandra Famelou", "latest_experience_title": "Media & PR Activation Manager", "url": "https://www.linkedin.com/in/alexandra-famelou-0a1734283/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Efi Karanasiou", "latest_experience_title": "Medical Manager Europe & Marketing Manager Greece", "url": "https://www.linkedin.com/in/efi-karanasiou-b0a2665/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Maria Ilektra Akarepi", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/maria-ilektra-akarepi-b8758112/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Evi Christodoulou", "latest_experience_title": "Brand Manager", "url": "https://www.linkedin.com/in/evi-christodoulou-3b2044113/", "email": None, "domain": "beiersdorf.com"},
        {"name": "Elisavet Terzivassian", "latest_experience_title": "Jr Brand Manager HealthCare", "url": "https://www.linkedin.com/in/elisavet-terzivassian-a0b560257/", "email": None, "domain": "beiersdorf.com"},
    ], "beiersdorf.com"),

    "Dove": ([
        {"name": "Georgios Karavas", "latest_experience_title": "Senior Brand Manager Deodorants(AXE, Dove & Dove Men+Care)", "url": "https://www.linkedin.com/in/-george-karavas10/", "email": None, "domain": "unilever.com"},
        {"name": "Konstantina Dimoka", "latest_experience_title": "Senior Brand Manager Laundry Skip", "url": "https://www.linkedin.com/in/konstantina-dimoka-ab23b659/", "email": None, "domain": "unilever.com"},
        {"name": "Michaela Karpouzi", "latest_experience_title": "Head of Marketing GR & CY at UFS", "url": "https://www.linkedin.com/in/michaela-karpouzi/", "email": None, "domain": "unilever.com"},
        {"name": "Katerina Ntalli", "latest_experience_title": "Marketing Manager Skin Cleansing & Oral Care", "url": "https://www.linkedin.com/in/katerinantalli/", "email": None, "domain": "unilever.com"},
        {"name": "Kelly Karli", "latest_experience_title": "Marketing Manager Deodorants, Hair & Skin Care", "url": "https://www.linkedin.com/in/kelly-karli-9bb731bb/", "email": None, "domain": "unilever.com"},
        {"name": "Asylena Demerouti", "latest_experience_title": "Marketing Manager Skin & Oral Care", "url": "https://www.linkedin.com/in/asylena-demerouti-2904481a/", "email": None, "domain": "unilever.com"},
        {"name": "Valia Sakkou", "latest_experience_title": "Digital-Media-Commerce Strategy Lead (VET DMC) 1U Europe & Country Media Lead for Greece, Spain, CBB", "url": "https://www.linkedin.com/in/valia-sakkou-a1756174/", "email": None, "domain": "unilever.com"},
        {"name": "Aspa Diki", "latest_experience_title": "Marketing Manager Foods", "url": "https://www.linkedin.com/in/aspadiki/", "email": None, "domain": "unilever.com"},
        {"name": "Vicky Statha", "latest_experience_title": "Brand & Portfolio Manager", "url": "https://www.linkedin.com/in/vicky-statha-b371031b7/", "email": None, "domain": "unilever.com"},
        {"name": "Ευσταθία Παπαθανασίου", "latest_experience_title": "Senior Brand Manager Klinex Greece", "url": "https://www.linkedin.com/in/efstathia-papathanasiou/", "email": None, "domain": "unilever.com"},
        {"name": "KATERINA ALIFRAGKI", "latest_experience_title": "Trade Marketing Lead Homecare", "url": "https://www.linkedin.com/in/katerina-alifragki-08a34115/", "email": None, "domain": "unilever.com"},
        {"name": "Ioanna Lioliou", "latest_experience_title": "Head of Marketing", "url": "https://www.linkedin.com/in/ioannalioliou/", "email": None, "domain": "unilever.com"},
        {"name": "Evangelia Vossou", "latest_experience_title": "Marketing Team Lead Personal Care 1U countries Europe", "url": "https://www.linkedin.com/in/evangelia-vossou-b1271714/", "email": None, "domain": "unilever.com"},
        {"name": "Ilias Faratzis", "latest_experience_title": "Head of Shopper Marketing", "url": "https://www.linkedin.com/in/ilias-faratzis-22881774/", "email": None, "domain": "unilever.com"},
        {"name": "Silia Patsou", "latest_experience_title": "Home Care Business Unit Marketing Lead", "url": "https://www.linkedin.com/in/silia-patsou-89074b59/", "email": None, "domain": "unilever.com"},
        {"name": "Nikoletta Isari", "latest_experience_title": "Brand Manager Ice Cream", "url": "https://www.linkedin.com/in/nikoletta-isari/", "email": None, "domain": "unilever.com"},
        {"name": "Eleni Stella", "latest_experience_title": "Brand Manager Skin Cleansing", "url": "https://www.linkedin.com/in/eleni-stella/", "email": None, "domain": "unilever.com"},
        {"name": "Afrodite Foniadaki", "latest_experience_title": "Head of Marketing", "url": "https://www.linkedin.com/in/afrodite-foniadaki-577767a3/", "email": None, "domain": "unilever.com"},
        {"name": "Athanasios Vourdas", "latest_experience_title": "Business Development Manager", "url": "https://www.linkedin.com/in/athanasios-vourdas-75620685/", "email": None, "domain": "unilever.com"},
        {"name": "Eleni Papathanasiou", "latest_experience_title": "Assistant Brand Marketing Skip", "url": "https://www.linkedin.com/in/eleni-papathanasiou-b34baa196/", "email": None, "domain": "unilever.com"},
    ], "unilever.com"),

    "Converse": ([
        {"name": "Artemis Schistou", "latest_experience_title": "Senior Digital Merchandiser", "url": "https://www.linkedin.com/in/artemis-schistou-99827275/", "email": None, "domain": "converse.com"},
    ], "converse.com"),

    "BMW": ([
        {"name": "Dimitra Bikou", "latest_experience_title": "BMW Marketing Manager", "url": "https://www.linkedin.com/in/dimitra-bikou-7ab3a129/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Alexandra Karvouni", "latest_experience_title": "Marketing & Project Specialist", "url": "https://www.linkedin.com/in/alexandra-karvouni-192030130/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Michael Gaganis", "latest_experience_title": "MINI Marketing Manager", "url": "https://www.linkedin.com/in/michael-gaganis-9a439b7/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Irene-Penelope Angeletou", "latest_experience_title": "Digital & CRM Specialist", "url": "https://www.linkedin.com/in/ireneangeletou/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Katerina Ninou", "latest_experience_title": "Customer Support Marketing & Projects Specialist", "url": "https://www.linkedin.com/in/κατερίνα-νίνου/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Elli Stasinou", "latest_experience_title": "Events Expert & Dealer Marketing Coordinator", "url": "https://www.linkedin.com/in/elli-stasinou-b7b43559/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Dimitris Kokalias", "latest_experience_title": "BMW Motorrad Sales and Marketing Specialist", "url": "https://www.linkedin.com/in/dimitris-kokalias-860a38271/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Elpida Sella", "latest_experience_title": "Motorrad Marketing and Sales Specialist", "url": "https://www.linkedin.com/in/elpida-sella-73bb45279/", "email": None, "domain": "bmwgroup.com"},
        {"name": "Lisa Ceruvija", "latest_experience_title": "Senior Marketing Specialist", "url": "https://www.linkedin.com/in/lisa-ceruvija-859b09246/", "email": None, "domain": "bmwgroup.com"},
    ], "bmwgroup.com"),

    "Ferryhopper": ([
        {"name": "Kaj van Zweeden", "latest_experience_title": "Chief Marketing Officer", "url": "https://www.linkedin.com/in/kaj-van-zweeden-43a552b6/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Sofia Drogoudi", "latest_experience_title": "Design Lead [Marketing]", "url": "https://www.linkedin.com/in/sofiadrogoudi/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Athina Sofou", "latest_experience_title": "Performance Marketing Manager", "url": "https://www.linkedin.com/in/athinasofou/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Christina Gkini", "latest_experience_title": "Head of Performance Marketing", "url": "https://www.linkedin.com/in/christina-gkini/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Michalis Kakotaritis", "latest_experience_title": "Brand Lead", "url": "https://www.linkedin.com/in/michalis-kakotaritis-88844493/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Aristeidis Remoundos", "latest_experience_title": "Affiliate Marketing Specialist", "url": "https://www.linkedin.com/in/aristeidis-remoundos-49a2701ba/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Nicole Adonopoulos", "latest_experience_title": "Brand Specialist", "url": "https://www.linkedin.com/in/nicole-adonopoulos-5683b462/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Andromachi Ferro", "latest_experience_title": "Marketing Expert", "url": "https://www.linkedin.com/in/andromachi-ferro-55409016b/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Chris Michalopoulos", "latest_experience_title": "Creative Marketing Specialist", "url": "https://www.linkedin.com/in/chris-michalopoulos/", "email": None, "domain": "ferryhopper.com"},
        {"name": "Ioannis Floros", "latest_experience_title": "Social Media Marketing Specialist", "url": "https://www.linkedin.com/in/ioannis-floros-b9b5681a9/", "email": None, "domain": "ferryhopper.com"},
    ], "ferryhopper.com"),

    "Vans": ([
        {"name": "Labrini Kranioti", "latest_experience_title": "Marketing Executive", "url": "https://www.linkedin.com/in/labrinikranioti/", "email": None, "domain": "vans.com"},
    ], "vans.com"),
}

_LAST_BATCH = {
    "ANT1": ([
        {"name": "Konstantinos Bourounis", "latest_experience_title": "Chief Marketing Officer - Antenna Group Greece", "url": "https://www.linkedin.com/in/konstantinos-bourounis-8675a0/", "email": None, "domain": "ant1.gr"},
        {"name": "Tasos Katsikadakos", "latest_experience_title": "Commercial, Digital & Research Director", "url": "https://www.linkedin.com/in/tasos-katsikadakos-4566856/", "email": None, "domain": "ant1.gr"},
        {"name": "Eva Vazaka", "latest_experience_title": "Commercial Manager | ANT1 Events & Sponsorships", "url": "https://www.linkedin.com/in/evazaka/", "email": None, "domain": "ant1.gr"},
        {"name": "Olympia Tsamasfyra", "latest_experience_title": "Marketing Manager", "url": "https://www.linkedin.com/in/olympia-tsamasfyra-40033069/", "email": None, "domain": "ant1.gr"},
        {"name": "Nancy Fafouti", "latest_experience_title": "Social Media Manager", "url": "https://www.linkedin.com/in/nancy-fafouti/", "email": None, "domain": "ant1.gr"},
        {"name": "Agapi Kantartzi", "latest_experience_title": "Marketing Manager & Communications | Antenna Audio / easy 97.2 | Ρυθμος 94.9 | Soundis.gr", "url": "https://www.linkedin.com/in/agapi/", "email": None, "domain": "ant1.gr"},
        {"name": "Thetis Gouliou", "latest_experience_title": "Chief Strategy & Business Development Officer", "url": "https://www.linkedin.com/in/thetis-gouliou-63680934/", "email": None, "domain": "ant1.gr"},
        {"name": "Tina Magaraki", "latest_experience_title": "Head of Business Development", "url": "https://www.linkedin.com/in/tina-magaraki-46b8086/", "email": None, "domain": "ant1.gr"},
        {"name": "Savvas Vitalis", "latest_experience_title": "Head of Digital Sales", "url": "https://www.linkedin.com/in/savvas-vitalis-00082b27/", "email": None, "domain": "ant1.gr"},
        {"name": "Kiki Malerdou", "latest_experience_title": "Commercial Manager Radio & Events", "url": "https://www.linkedin.com/in/kiki-malerdou-8803742a/", "email": None, "domain": "ant1.gr"},
        {"name": "Iraklis Ioannidis", "latest_experience_title": "Music Marketing Lead - Marketing Manager Heaven/Warner & Antenna Intelligence", "url": "https://www.linkedin.com/in/iraklis-ioannidis-0a536622/", "email": None, "domain": "ant1.gr"},
        {"name": "Ioanna Panagioti", "latest_experience_title": "Social Media Expert", "url": "https://www.linkedin.com/in/ioanna-panagioti-25bba91a9/", "email": None, "domain": "ant1.gr"},
        {"name": "Vlasia Manteska", "latest_experience_title": "Digital Sales Executive", "url": "https://www.linkedin.com/in/vlasia-manteska-848032140/", "email": None, "domain": "ant1.gr"},
        {"name": "Marco Struecker", "latest_experience_title": "Interim General Manager - ANT1+", "url": "https://www.linkedin.com/in/marcostruecker/", "email": None, "domain": "ant1.gr"},
    ], "ant1.gr"),

    "AEK FC": ([
        {"name": "Antonis Apostolopoulos", "latest_experience_title": "Event Management & Sports Marketing Assistant", "url": "https://www.linkedin.com/in/antonis-apostolopoulos-7b753b187/", "email": None, "domain": "aekfc.gr"},
        {"name": "Christina Koromila", "latest_experience_title": "Social Media", "url": "https://www.linkedin.com/in/christina-koromila-768326b5/", "email": None, "domain": "aekfc.gr"},
    ], "aekfc.gr"),

    "Jumbo": ([
        {"name": "Andreas Xenofontos", "latest_experience_title": "Head of Digital Marketing & E-commerce", "url": "https://www.linkedin.com/in/andreas-xenofontos/", "email": None, "domain": "e-jumbo.gr"},
        {"name": "Maria Gkara", "latest_experience_title": "Media and Communications Consultant", "url": "https://www.linkedin.com/in/maria-gkara-7940535/", "email": None, "domain": "e-jumbo.gr"},
        {"name": "Andra Bilici", "latest_experience_title": "Digital Marketing & E-commerce Manager", "url": "https://www.linkedin.com/in/andra-bilici-731278144/", "email": None, "domain": "e-jumbo.gr"},
        {"name": "Panagiotis Mylonidis", "latest_experience_title": "E-Commerce & Social Media Manager", "url": "https://www.linkedin.com/in/pmylonidis/", "email": None, "domain": "e-jumbo.gr"},
    ], "e-jumbo.gr"),

    "FAGE": ([
        {"name": "Katerina Gkouvaki", "latest_experience_title": "Senior Brand Manager GR", "url": "https://www.linkedin.com/in/katerina-gkouvaki-743462183/", "email": None, "domain": "home.fage"},
        {"name": "Alexis Alexopoulos", "latest_experience_title": "Marketing & Communications Director", "url": "https://www.linkedin.com/in/alexis-alexopoulos-8a539b1/", "email": None, "domain": "home.fage"},
    ], "home.fage"),

    "New Balance": ([
        {"name": "Eleana Ravani", "latest_experience_title": "Trade & Marketing Communication", "url": "https://www.linkedin.com/in/eleana-ravani-917520193/", "email": None, "domain": "newbalance.com"},
        {"name": "Panagiotis Kyvrikosaios", "latest_experience_title": "Retail Marketing and Running Category Manager", "url": "https://www.linkedin.com/in/panoskyvrikosaios/", "email": None, "domain": "newbalance.com"},
    ], "newbalance.com"),

    "Zara": ([
        {"name": "Mary Geroukali", "latest_experience_title": "Communications - PR & Marketing Director (Greece)", "url": "https://www.linkedin.com/in/mary-geroukali-05250731/", "email": None, "domain": "inditexpeople.com"},
    ], "inditexpeople.com"),

    "Puma": ([
        {"name": "Esteve Planas", "latest_experience_title": "General Manager PUMA Southern Europe", "url": "https://www.linkedin.com/in/esteveplanas/", "email": None, "domain": "puma.com"},
        {"name": "Pedro Moscoso", "latest_experience_title": "Marketing & Communications Director Southern Europe", "url": "https://www.linkedin.com/in/pedromoscoso/", "email": None, "domain": "puma.com"},
        {"name": "Carmen Ponce", "latest_experience_title": "Lead Brand Consumer Southern Europe", "url": "https://www.linkedin.com/in/carmenponce6/", "email": None, "domain": "puma.com"},
    ], "puma.com"),

    "CarVertical": ([
        {"name": "Mantas Ribelis", "latest_experience_title": "Chief Marketing Officer", "url": "https://www.linkedin.com/in/mantasribelis/", "email": None, "domain": "carvertical.com"},
        {"name": "Ovidijus S.", "latest_experience_title": "Influencer Marketing Performance Manager", "url": "https://www.linkedin.com/in/ovidijussaldauskas/", "email": None, "domain": "carvertical.com"},
        {"name": "Tautvydas Sirevicius", "latest_experience_title": "Influencer Marketing Manager", "url": "https://www.linkedin.com/in/tautvydas-sireviäius/", "email": None, "domain": "carvertical.com"},
        {"name": "Deivydas Simas", "latest_experience_title": "Influencer Marketing Strategist", "url": "https://www.linkedin.com/in/deivydas-å¡imas/", "email": None, "domain": "carvertical.com"},
        {"name": "Gerda Servaite", "latest_experience_title": "Influencer Marketing Manager", "url": "https://www.linkedin.com/in/gerda-servait%c4%97-647774151/", "email": None, "domain": "carvertical.com"},
        {"name": "Ieva Pauliukoniene", "latest_experience_title": "Manager Strategic Partnerships", "url": "https://www.linkedin.com/in/ievapaul/", "email": None, "domain": "carvertical.com"},
        {"name": "Daniel Artisiuk", "latest_experience_title": "Senior Country Manager | B2B Sales & Partnerships, EU Markets", "url": "https://www.linkedin.com/in/danielartisiuk/", "email": None, "domain": "carvertical.com"},
    ], "carvertical.com"),
}


if __name__ == "__main__":
    if "save" in sys.argv:
        if not BATCH:
            print("BATCH is empty — nothing to save.")
        else:
            for brand, (contacts, domain) in BATCH.items():
                save_contacts(brand, contacts, domain)
            print("Done.")
    elif "mark_empty" in sys.argv:
        idx = sys.argv.index("mark_empty")
        brands = sys.argv[idx + 1:]
        if not brands:
            print("Usage: python discover_contacts.py mark_empty \"Brand1\" \"Brand2\" ...")
        else:
            for b in brands:
                _mark_brand_empty(b)
                state = _load_state()
                attempts = state["empty"][b]["attempts"]
                print(f"  Marked '{b}' as EMPTY (attempt #{attempts})")
    elif "unmark_empty" in sys.argv:
        idx = sys.argv.index("unmark_empty")
        brands = sys.argv[idx + 1:]
        if not brands:
            print("Usage: python discover_contacts.py unmark_empty \"Brand1\" ...")
        else:
            for b in brands:
                _unmark_brand_empty(b)
                print(f"  '{b}' removed from empty list — will appear in next batch again")
    else:
        status_report()
