---
name: clay-discoverer
description: >
  Clay MCP contact enrichment batch processor. Use when the user asks to save, process, or
  store contacts retrieved via the Clay MCP tool (find-and-enrich-contacts-at-company).
  Handles merging raw Clay task-context results into the per-brand employee store at
  output/employees/<brand_name>.json, with dedup by normalized name, domain filtering,
  email extraction from enrichments[], and a status report of which brands have employee
  files and how many emails are known.
  Also contains DOMAIN_MAP for per-brand companyIdentifier overrides (domain preferred over
  LinkedIn slug), and GLOBAL_BRANDS for brands that need a Greece location filter.
  Invoke when the user says "save clay contacts", "save the batch", "run discover_contacts",
  "store the clay results", "show me the contact status report", or "enrich emails".
---

# clay-discoverer

## Purpose

Companion script to the Clay MCP workflow. After Claude fires Clay MCP calls and collects
task contexts via `get-task-context`, this script saves the raw contact dicts into the
persistent employee store at `AI Sales Agent System/output/employees/<brand_name>.json`.

It is **not** a stand-alone discoverer — it is the save/merge/filter layer.

---

## Usage

```bash
# Print status report: full employee store overview + next batch recommendation
python ".claude/skills/clay-discoverer/discover_contacts.py"

# Phase 1: Claude populates BATCH dict in discover_contacts.py, then:
python ".claude/skills/clay-discoverer/discover_contacts.py" save

# ALWAYS run after save — commits + pushes employee files + actionable_contacts.json to GitHub
python ".claude/skills/clay-discoverer/discover_contacts.py" git_sync

# Phase 3: print contacts without emails, ready for find-and-enrich-list-of-contacts
python ".claude/skills/clay-discoverer/discover_contacts.py" enrich_emails "Brand1" "Brand2"

# Export to Excel
python ".claude/skills/clay-discoverer/discover_contacts.py" export
```

---

## 3-Phase Clay Workflow

### Phase 1 — Employee discovery (domain → employees)

Use `find-and-enrich-contacts-at-company` with the brand's **domain** (not LinkedIn slug):

```
companyIdentifier: "stoiximan.gr"          ← from DOMAIN_MAP or status report
numberOfContacts: 11
contactFilters: { locations: ["Greece"] }
```

Clay resolves the correct LinkedIn entity from domain internally. The `DOMAIN_MAP` dict
in the script has known good domains for ~44 brands. For brands not yet in DOMAIN_MAP,
the status report shows the LinkedIn URL as fallback — add the domain to DOMAIN_MAP once
confirmed.

After getting task context from `get-task-context`, populate `BATCH` and run:
```bash
python discover_contacts.py save
python discover_contacts.py git_sync
```
Writes `00-BrandName.json`. `BANNED_TITLES` applied automatically on save.
`git_sync` commits and pushes the results to GitHub — **always run it**, otherwise
the data stays on the agent's machine only and is lost on reset.

### Phase 2 — Review / title filter

`save` applies `BANNED_TITLES` automatically. To retroactively clean existing files:
```bash
python discover_contacts.py purge
```

### Phase 3 — Email enrichment (filtered contacts → emails)

```bash
python discover_contacts.py enrich_emails "Brand1" "Brand2"
```

Outputs a JSON array to stdout:
```json
[
  {"contactName": "Maria Papadopoulos", "companyIdentifier": "brand.gr"},
  ...
]
```

Pass this to `find-and-enrich-list-of-contacts` with:
```
dataPoints: { contactDataPoints: [{ type: "Email" }] }
```

After getting the enriched results via `get-task-context`, populate `BATCH` as usual and
run `save` again — the merge layer fills in emails for existing contacts.

---

## Status report — what it shows

1. **Employee store overview** — scans ALL files in `output/employees/`.
2. **Clay-searchable brands** — coverage per brand:
   - ✓ Covered: have ≥1 email
   - ∅ LinkedIn-only: have contacts but 0 emails
   - ✗ Empty: Clay returned 0 contacts (ZZ- file)
   - ✗ Missing: no employee file
3. **Next batch recommendation** — ranked list of brands to search next.

The `Identifier` column shows the domain when in `DOMAIN_MAP`, LinkedIn URL otherwise.

---

## DOMAIN_MAP

Maps brand → preferred `companyIdentifier` (domain). Domains are preferred over LinkedIn
slugs because Clay resolves the correct entity more reliably from domain, and avoids slug
ambiguity (wrong region, wrong company entity, 0-result slugs).

Add new entries when a brand's domain is confirmed. Brands not in DOMAIN_MAP fall back to
`linkedin_slug` from `actionable_contacts.json`.


---

## Output

Writes directly to `AI Sales Agent System/output/employees/<brand_name>.json`.

Contact schema per entry:
```json
{
  "name": "...",
  "job_title": "...",
  "email": "...",
  "linkedin_url": "...",
  "verified": false,
  "active": true,
  "source": "clay",
  "domain": "...",
  "first_seen": "YYYY-MM-DD",
  "last_seen": "YYYY-MM-DD"
}
```

