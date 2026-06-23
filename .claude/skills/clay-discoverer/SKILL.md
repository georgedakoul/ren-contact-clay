---
name: clay-discoverer
description: >
  Clay MCP contact enrichment batch processor. Use when the user asks to save, process, or
  store contacts retrieved via the Clay MCP tool (find-and-enrich-contacts-at-company).
  Handles merging raw Clay task-context results into the per-brand employee store at
  output/employees/<brand_name>.json, with dedup by normalized name, domain filtering,
  email extraction from enrichments[], and a status report of which brands have employee
  files and how many emails are known.
  Also contains IDENTIFIER_OVERRIDES for brands where the LinkedIn slug resolves to the
  wrong company entity, and GLOBAL_BRANDS for brands that need a Greece location filter.
  Invoke when the user says "save clay contacts", "save the batch", "run discover_contacts",
  "store the clay results", or "show me the contact status report".
---

# clay-discoverer

## Purpose

Companion script to the Clay MCP workflow. After Claude fires
`find-and-enrich-contacts-at-company` calls and collects task contexts via
`get-task-context`, this script saves the raw contact dicts into the persistent
employee store at `AI Sales Agent System/output/employees/<brand_name>.json`.

It is **not** a stand-alone discoverer — it is the save/merge layer that persists Clay
results across sessions so they don't have to be re-fetched.

---

## Usage

```bash
# Print status report: full employee store overview + next batch recommendation
python ".claude/skills/clay-discoverer/discover_contacts.py"

# Save a batch (Claude fills the BATCH dict in the script first)
python ".claude/skills/clay-discoverer/discover_contacts.py" save
```

---

## Status report — what it shows

The report has three passes:

1. **Employee store overview** — scans ALL files in `output/employees/` first.
   Shows total files, total contacts, total emails across the whole store.

2. **Clay-searchable brands** — of the 80 brands with a Clay identifier (linkedin_slug
   or IDENTIFIER_OVERRIDES entry), shows:
   - ✓ Covered: have ≥1 email
   - ∅ Partial: have contacts but 0 emails (typically global brands needing `+GR` filter)
   - ✗ Missing: no employee file at all

3. **Next batch recommendation** — explicit ranked list of which brands to search next
   (missing first, then partial), ready to copy into a Clay MCP call sequence.

---

## Workflow with Clay MCP

1. Run the status report to see the "Next batch recommendation"
2. Claude calls `find-and-enrich-contacts-at-company` for each brand → gets `taskId`
3. Claude calls `get-task-context(taskId)` → gets raw contact list with enrichments
4. Claude edits the `BATCH` dict in this script with the results
5. Claude runs `python discover_contacts.py save` to merge into employee store

---

## IDENTIFIER_OVERRIDES

Brands where the LinkedIn company slug resolves to the wrong entity or where the domain
gives better Clay results. Edit `IDENTIFIER_OVERRIDES` at the top of the script to add
new overrides when Clay returns results from the wrong company.

## GLOBAL_BRANDS

Brands that need `contactFilters: {locations: ["Greece"]}` in the Clay call to avoid
returning non-Greek employees from global headquarters. Check this set before firing Clay
for a multinational brand.

---

## Output

Writes directly to `AI Sales Agent System/output/employees/<brand_name>.json`.
No separate output file — the employee store IS the output.

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

---

## Related skills

- `linkedin-discoverer` — Playwright LinkedIn scraper (higher recall, requires browser session)
- `apollo-discoverer` — Apollo.io free-tier name+title lookup (no email on free tier)
- `contact-discoverer` — DNS/Hunter.io/scraping pipeline (lowest detection risk)
