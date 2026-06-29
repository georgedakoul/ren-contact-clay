"""google-domain-discoverer.py — Find and store official website domains for all brands in
actionable_contacts.json, so Clay can use domain (more reliable than LinkedIn slug) for
every contact discovery call.

Commands:
  python google-domain-discoverer.py                       → status report
  python google-domain-discoverer.py backfill              → fill from employee files + SEED_DOMAINS
  python google-domain-discoverer.py discover [--limit N] [--dry-run]
                                                           → HTTP-probe candidate domains (no API key)
  python google-domain-discoverer.py export-missing        → write CSV with Google search URLs for manual lookup
  python google-domain-discoverer.py import domains.csv    → apply brand_name,domain rows from CSV
  python google-domain-discoverer.py set "Brand Name" domain.com
                                                           → manually set one brand

Domain priority (highest → lowest):
  1. SEED_DOMAINS / `set` command (trusted manual, never auto-overwritten)
  2. Employee file domain (most common domain from Clay contacts already saved)
  3. HTTP-probed candidate (discover command)
  4. Manual CSV import (export-missing → fill → import)

discover strategy:
  Generates candidate domains from the brand name slug (.com/.gr/.eu/.net/.org),
  sends a HEAD request to each, accepts the first that responds.
  ~2-5s per brand. Covers most well-known brands without any API key.
  Brands that don't resolve → left as null for the CSV workflow.
"""
import csv, json, re, sys, time, urllib.request, urllib.parse, urllib.error
from collections import Counter
from pathlib import Path

ROOT          = Path(__file__).resolve().parents[3]
CONTACTS_FILE = ROOT / "AI Sales Agent System" / "actionable_contacts.json"
STORE_DIR     = ROOT / "AI Sales Agent System" / "output" / "employees"
MISSING_CSV   = ROOT / "AI Sales Agent System" / "output" / "domains_missing.csv"

# Trusted manual seeds — ambiguous or wrong LinkedIn slugs.
# Never auto-overwritten. Add entries here when confirmed.
SEED_DOMAINS = {
    "COSMOTE":          "cosmote.gr",
    "Village Cinemas":  "villagecinemas.gr",
    "Cosmos Sport":     "cosmossport.gr",
    "Coca-Cola":        "coca-colahellenic.com",
    "Psichogios Books": "psichogios.gr",
    "Pame Stoixima":    "opap.gr",
    "ION":              "ion.gr",
    "instacar":         "instacar.gr",
    "Apivita":          "apivita.com",
    "more.com":         "more.com",
    "Wind":             "wind.gr",
    "Alterlife":        "alterlife.gr",
    "SKY express":      "skyexpress.gr",
    "Carroten":         "carroten.gr",
    "Fresh Line":       "freshline.gr",
    "Alumil":           "alumil.com",
    "BSB Fashion":      "bsbfashion.com",
    "La Vie en Rose":   "lavieenrose.com",
    "Mind Your Style":  "mindyourstyle.gr",
    "Protergia":        "protergia.gr",
    "Douleutaras":      "douleutaras.gr",
    "Box Now":          "boxnow.gr",
    "Germanos":         "germanos.gr",
    "MEVGAL":           "mevgal.gr",
    "Vitex":            "vitex.gr",
    "Three Cents":      "threecents.com",
    "Snappi":           "snappibank.com",
    "LG":               "lg.com",
    "JYSK":             "jysk.com",
    "Kinder":           "ferrerocareers.com",
    "NIVEA":            "beiersdorf.com",
    "Dove":             "unilever.com",
    "Converse":         "converse.com",
    "BMW":              "bmwgroup.com",
    "Ferryhopper":      "ferryhopper.com",
    "Vans":             "vans.com",
    "ANT1":             "ant1.gr",
    "AEK FC":           "aekfc.gr",
    "Jumbo":            "e-jumbo.gr",
    "FAGE":             "home.fage",
    "New Balance":      "newbalance.com",
    "Zara":             "inditexpeople.com",
    "Puma":             "puma.com",
    "CarVertical":      "carvertical.com",
}

_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"


# ---------------------------------------------------------------------------
# I/O helpers
# ---------------------------------------------------------------------------

def _load():
    data = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
    brands = data if isinstance(data, list) else data.get("brands", [])
    return data, brands


def _save(data):
    CONTACTS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def _emp_domains():
    """Return {brand_name: most_common_domain} extracted from saved employee files."""
    result = {}
    for f in STORE_DIR.iterdir():
        if f.suffix != ".json":
            continue
        try:
            contacts = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(contacts, list):
            continue
        domains = [c.get("domain", "") for c in contacts if c.get("domain")]
        if domains:
            top = Counter(domains).most_common(1)[0][0]
            brand = f.stem.lstrip("Z").lstrip("0").lstrip("-").strip()
            result[brand] = top
    return result


# ---------------------------------------------------------------------------
# HTTP domain probe
# ---------------------------------------------------------------------------

def _slug(name: str) -> str:
    """Brand name → lowercase alphanumeric slug, spaces removed."""
    s = name.lower()
    s = re.sub(r"[^\w\s-]", "", s)   # keep word chars, spaces, hyphens
    s = re.sub(r"\s+", "", s)         # collapse spaces
    s = re.sub(r"-+", "-", s).strip("-")
    return s


def _probe(domain: str, timeout: int = 4) -> bool:
    """Return True if domain responds to HTTP(S) HEAD request."""
    for scheme in ("https", "http"):
        try:
            req = urllib.request.Request(
                f"{scheme}://{domain}",
                method="HEAD",
                headers={"User-Agent": _UA},
            )
            with urllib.request.urlopen(req, timeout=timeout) as _:
                return True
        except urllib.error.HTTPError:
            # Server responded (even 4xx) — domain exists
            return True
        except Exception:
            continue
    return False


def _guess_domain(brand_name: str) -> str | None:
    """Try common slug+TLD combos. Return first that responds, or None."""
    full = _slug(brand_name)
    # also try first-word slug for multi-word brands
    first = _slug(brand_name.split()[0]) if " " in brand_name else None

    candidates: list[str] = []
    for slug in filter(None, [full, first]):
        if len(slug) < 2:
            continue
        candidates += [
            f"{slug}.com",
            f"{slug}.gr",
            f"{slug}.eu",
            f"{slug}.net",
            f"{slug}.org",
        ]

    # dedupe while preserving order
    seen: set[str] = set()
    for c in candidates:
        if c not in seen:
            seen.add(c)
            if _probe(c):
                return c
    return None


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_status():
    _, brands = _load()
    emp = _emp_domains()
    total = len(brands)
    with_domain = sum(1 for b in brands if b.get("domain"))
    seeds_not_set = [k for k in SEED_DOMAINS if not any(
        b.get("brand") == k and b.get("domain") for b in brands
    )]
    emp_not_set = [br for br in emp if not any(
        b.get("brand") == br and b.get("domain") for b in brands
    )]

    print(f"actionable_contacts.json : {total} brands")
    print(f"  with domain            : {with_domain}")
    print(f"  missing domain         : {total - with_domain}")
    print()
    print(f"Quick wins (run `backfill`):")
    print(f"  SEED_DOMAINS not yet applied : {len(seeds_not_set)}")
    print(f"  employee-file domains unused : {len(emp_not_set)}")
    print()
    after = total - with_domain - len(seeds_not_set) - len(emp_not_set)
    print(f"Remaining after backfill : ~{max(after, 0)} brands")
    print(f"  → run `discover` for HTTP-probe, then `export-missing` for the rest")


def cmd_backfill():
    data, brands = _load()
    emp = _emp_domains()
    filled = 0

    for b in brands:
        name = b.get("brand", "")
        b.setdefault("domain", None)

        if b["domain"]:
            continue  # already set — never auto-overwrite

        if name in SEED_DOMAINS:
            b["domain"] = SEED_DOMAINS[name]
            filled += 1
            continue

        if name in emp:
            b["domain"] = emp[name]
            filled += 1

    _save(data)
    with_domain = sum(1 for b in brands if b.get("domain"))
    print(f"Backfill complete. Filled {filled} brands.")
    print(f"Total with domain: {with_domain}/{len(brands)}")
    print(f"Still missing    : {len(brands) - with_domain}")


def cmd_discover(limit: int = 100, dry_run: bool = False):
    data, brands = _load()
    missing = [b for b in brands if not b.get("domain")]

    if not missing:
        print("All brands have a domain.")
        return

    batch = missing[:limit]
    print(f"{len(missing)} missing. Probing {len(batch)} (--limit {limit}).")
    if dry_run:
        print("[dry-run] No writes.\n")

    found = 0
    for i, b in enumerate(batch, 1):
        name = b["brand"]
        print(f"[{i}/{len(batch)}] {name!r} ...", end=" ", flush=True)
        domain = _guess_domain(name)
        if domain:
            print(domain)
            found += 1
            if not dry_run:
                b["domain"] = domain
        else:
            print("—")

    if not dry_run:
        _save(data)
        print(f"\nSaved {found}/{len(batch)}.")

    remaining = len(missing) - len(batch)
    if remaining > 0:
        print(f"{remaining} still missing — run again or use export-missing.")


def cmd_export_missing():
    _, brands = _load()
    missing = [b for b in brands if not b.get("domain")]
    if not missing:
        print("No brands missing domain.")
        return

    MISSING_CSV.parent.mkdir(parents=True, exist_ok=True)
    with MISSING_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["brand_name", "industry", "times_advertised", "linkedin_slug", "domain", "google_url"])
        for b in missing:
            name = b["brand"]
            gurl = "https://www.google.com/search?q=" + urllib.parse.quote(f"{name} official website Greece")
            w.writerow([name, b.get("industry", ""), b.get("times_advertised", ""), b.get("linkedin_slug", ""), "", gurl])

    print(f"Wrote {len(missing)} brands to:")
    print(f"  {MISSING_CSV}")
    print()
    print("Fill the `domain` column, then run:")
    print("  python google-domain-discoverer.py import domains_missing.csv")


def cmd_import(csv_path: str):
    path = Path(csv_path)
    if not path.exists():
        # try relative to MISSING_CSV dir
        path = MISSING_CSV.parent / csv_path
    if not path.exists():
        print(f"File not found: {csv_path}")
        sys.exit(1)

    data, brands = _load()
    brand_index = {b["brand"]: b for b in brands}

    applied = skipped = not_found = 0
    with path.open(encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("brand_name") or row.get("brand") or "").strip()
            domain = (row.get("domain") or "").strip()
            if not name or not domain:
                skipped += 1
                continue
            if name not in brand_index:
                print(f"  not found in JSON: {name!r}")
                not_found += 1
                continue
            brand_index[name]["domain"] = domain
            applied += 1

    _save(data)
    print(f"Import complete: {applied} set, {skipped} skipped (empty), {not_found} brand names not found.")


def cmd_set(brand_name: str, domain: str):
    data, brands = _load()
    b = next((x for x in brands if x.get("brand") == brand_name), None)
    if b is None:
        print(f"Brand not found: {brand_name!r}")
        sys.exit(1)
    old = b.get("domain")
    b["domain"] = domain
    _save(data)
    print(f"Set {brand_name!r}: {old!r} → {domain!r}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]

    if not args:
        cmd_status()
        return

    cmd = args[0]

    if cmd == "backfill":
        cmd_backfill()

    elif cmd == "discover":
        limit, dry_run = 100, False
        i = 1
        while i < len(args):
            if args[i] == "--limit" and i + 1 < len(args):
                limit = int(args[i + 1]); i += 2
            elif args[i] == "--dry-run":
                dry_run = True; i += 1
            else:
                i += 1
        cmd_discover(limit=limit, dry_run=dry_run)

    elif cmd == "export-missing":
        cmd_export_missing()

    elif cmd == "import":
        if len(args) < 2:
            print("Usage: python google-domain-discoverer.py import <path.csv>")
            sys.exit(1)
        cmd_import(args[1])

    elif cmd == "set":
        if len(args) < 3:
            print('Usage: python google-domain-discoverer.py set "Brand Name" domain.com')
            sys.exit(1)
        cmd_set(args[1], args[2])

    else:
        print(f"Unknown command: {cmd!r}")
        print("Commands: backfill | discover [--limit N] [--dry-run] | export-missing | import <file.csv> | set \"Brand\" domain.com")
        sys.exit(1)


if __name__ == "__main__":
    main()
