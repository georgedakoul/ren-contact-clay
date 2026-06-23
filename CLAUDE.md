# Renfluence — project context for Claude

## Critical: the user is on the Renfluence MCP team

The user works at the company that builds and operates the **Renfluence MCP server**
(`mcp__d12b4540-d612-4aa0-af3b-5ad8803da768__*` tools). They can request modifications
to the MCP itself.

**Implication for every future answer**: when a workflow hits an MCP limitation, do
NOT silently route around it. Surface it as a candidate MCP-side change *first*, then
offer the workaround as fallback. Examples of valid suggestions:

- "Field X is missing from response Y — recommend adding it server-side."
- "Tool Z requires a UUID we don't have — recommend a new lookup tool / parameter."
- "This filter is too coarse — recommend a new option."

The user's mental model: improving the MCP is cheaper at scale than building
client-side workarounds in every consumer skill.

## Shipped MCP modifications (as of 2026-06-05)

| Change | Status | Why |
| --- | --- | --- |
| Add `profile_uuid` to each item returned by `search_media` | **Shipped** | Removes the post→creator dead-end. Lets brand-search → group by creator in one paginated query instead of forcing `search_profiles` as a separate entry point. |

## Requested MCP modifications (not yet shipped)

These are server-side changes flagged during /grill-me sessions. They have client-side
workarounds in place (or are deferred), but shipping them simplifies consumer skills
and reduces call volume. Sorted by priority.

### High priority — `save_brand_sentiment_report` and siblings (for Sentiment_Analyser)

The per-comment sentiment classification is already persisted server-side (via
`create_sentiment_analysis`). But the **derived aggregate layer** (overall_score,
brand_label, themes, per_creator_sentiment, per_post_sentiment, trend) is currently
written to a local JSON file. That blocks cross-machine access for Message_Crafter,
the future React/Next.js dashboard, and any teammate not on the machine that ran the
skill. Proposed endpoints:

```
save_brand_sentiment_report(
    brand_name:                  str,
    analysis_date:               str (ISO date),
    mode:                        "flagged_only" | "all_brand_posts",
    lookback_days:               int,
    source_perf_analysis_uuid:   str | null,    # link to upstream Performance_Analyser run
    overall_score:               float,         # 0–10, engagement-weighted
    brand_label:                 "Negative" | "Neutral" | "Positive",
    positive_pct:                int,
    neutral_pct:                 int,
    negative_pct:                int,
    by_platform:                 [{ platform, score, positive_pct, neutral_pct, negative_pct, comments }],
    per_creator_sentiment:       [{ profile_uuid, name, handle, score, positive_pct, neutral_pct, negative_pct, post_count, comment_count }],
    per_post_sentiment:          [{ media_uuid, profile_uuid, score, positive_pct, neutral_pct, negative_pct, comment_count, perf_priority, perf_score }],
    themes:                      [{ theme, mention_count, sentiment, sentiment_score, sample_comments, source_breakdown }],
    trend:                       "improving" | "stable" | "declining" | "insufficient_data",
    trend_delta:                 float,
    total_comments_analyzed:     int,
    credits_spent_estimated:     int,
    notes:                       [str]
) -> brand_sentiment_report_uuid

get_brand_sentiment_report(report_uuid: str) -> <full report>

get_latest_brand_sentiment_report(brand_name: str) -> <full report> | null

list_brand_sentiment_reports(
    brand_name: str | null,
    date_from:  str | null,
    date_to:    str | null,
    mode:       str | null,
    page:       int = 1,
    page_size:  int = 50
) -> { items: [{ report_uuid, brand_name, analysis_date, mode, overall_score, brand_label }], total, page, page_size }
```

When this ships, the Sentiment_Analyser activates Step 7 (push to server) via the
`cfg.feature_flags.save_to_server` config flag. The local JSON file stays as a
fallback and audit trail.

### Lower priority — smaller server-side wins

| Change | Skill that benefits | Workaround in place today |
| --- | --- | --- |
| ~~`search_media(paid_partnership_only=true)` filter~~ — **WITHDRAWN 2026-06-10** | — | Greek-market reality: creators rarely use the platform "Paid partnership" disclosure, so a server-side `paid_partnership_only=true` filter would zero out most real campaigns. Not actionable. Performance_Analyser uses `min_likelihood_score + max_brands + is_target_brand` heuristic instead and labels results as "AI-detected likely partnerships." |
| Per-post AI signal that distinguishes brand-central content from brand-incidental content (e.g. `brand_role: "primary"` / `"prominent"` / `"background"`) | Performance_Analyser, Sentiment_Analyser | Today the only signal is `likelihood_score` which combines presence confidence with importance. A role/centrality field would let us drop life-hack videos and aesthetic placements without losing real partnerships. Higher value for Greek market where platform disclosure is rare. |
| `search_media` response items include a canonical `media_url` / `permalink` field | Orchestrator queue, Sentiment_Analyser | Currently the client has to guess URL shape from `platform` + `platform_id`, which is ambiguous (19-digit numeric IDs match both Instagram stories and TikTok videos). Without this field, the Orchestrator surfaces `video_uuid` + null URL when ambiguous. Confirmed necessary 2026-06-10. |
| `create_sentiment_analysis(media_uuid=...)` accepts UUID directly (alternative to `media_url`) | Sentiment_Analyser Step 2a | Construct URL from `platform_id` + handle; brittle. |
| `list_sentiment_analyses(media_uuids=[...])` — filter by one or more media UUIDs server-side | Sentiment_Analyser Step 1 | Paginate the **full account history** client-side and build an in-memory dict. See scale note below. |
| `get_sentiment_analysis` and `get_media_comments` comment objects include `like_count` per comment | Sentiment_Analyser scoring | Confirmed absent 2026-06-10. Blocks true engagement-weighted scoring (`score = Σ(label_value × (1 + likes)) / Σ(1 + likes)`). Without it all comments score equally regardless of resonance — highest-impact enhancement for scoring accuracy. |
| `get_media_comments` includes `platform` per comment row | Sentiment_Analyser (informational) | **Solved**: `get_sentiment_analysis` `media.platform` is used directly (confirmed 2026-06-10). No longer blocks the workflow. Shipping makes `get_media_comments` self-contained but is low priority. |
| `list_sentiment_analyses(analysis_uuids=[...])` for batch polling | Sentiment_Analyser Step 2c poll loop | N parallel calls, one per pending analysis. |
| `save_perf_analysis_report(...)` (parallel to brand sentiment) | Performance_Analyser | Local JSON file only. Same cross-machine problem. Lower priority because flagged campaigns are rederivable from raw data. |

#### Scale note — why `list_sentiment_analyses(media_uuids=[...])` matters

Before triggering any new classifications, Sentiment_Analyser checks which of the current brand's
flagged posts already have a completed analysis (skip them — free hit) vs. are pending (wait) vs.
are absent (trigger). That lookup needs to happen for ~8–15 UUIDs per brand run.

Without the filter, the only option is to paginate the **entire account's analysis history** and
build a local dict. Today the account is empty — 1 call, instant. But this skill will run on 84
brands repeatedly:

| Stage | Analyses in account | Pages (100/page) | Calls just to build the index |
| --- | --- | --- | --- |
| First runs | 0–50 | 1 | 1 |
| After 84 brands × 1 run | ~700 | 7 | 7 |
| After 84 brands × 5 runs | ~3,500 | 35 | 35 |
| After 1 year of weekly runs | 20,000+ | 200+ | 200+ |

Every brand run pays the full cost of reading every analysis ever created, just to look up the 8–15
UUIDs it cares about. With `media_uuids=[...]` on the server, Step 1 becomes **1 call regardless of
account age** and the pagination loop disappears.

The Performance_Analyser now runs **Flow A (post-first)** exclusively:
1. `search_media(brand_queries=[…])` — paginated, returns posts WITH `profile_uuid`
2. Group by `profile_uuid`, apply ≥3 gate
3. Batch call `get_profile_overviews(profile_uuids=[…])` — baseline + enrichment in one call, up to 20 at a time
4. Score locally
5. `get_profile_overviews` response cached — no separate enrichment step needed

Result: ~20 calls/brand for ~10 creators, no `search_profiles` step needed, no
caption-search scope concerns. Flow B (profile-first) is retained in the skill as dead
code until formally removed.

## Project layout

- `AI Sales Agent System/` — the AI sales pipeline (this is the work)
  - `Orchestrator Agent.md` — design doc for the brain. The runnable version lives in `.claude/skills/orchestrator/`.
  - `SKILL Descriptions/` — design docs for each skill (`Performance_Analyser.md`, `Sentiment_Analyser.md`, `Message_Crafter.md`)
  - `actionable_contacts.json` — 84 brands with decision-maker contacts. Input to the pipeline.
  - `orchestrator_config.json` — central config (the spec's shared config.json). Cooldowns, freshness, tier matrix, feature flags, stub message templates.
  - `perf_analyzer_config.json` — runtime config for the `performance-analyser` skill.
  - `sentiment_analyzer_config.json` — runtime config for the `sentiment-analyser` skill.
  - `contact_discoverer_config.json` — runtime config for the `contact-discoverer` skill.
  - `message_crafter_config.json` — runtime config for the `message-crafter` skill (language, persona rubric, tier→theme map, Greek intro/CTA/subject banks, writing rules, PDF spec).
  - `output/` — per-run JSON reports from each skill (plus `orchestrator_run_*.json`, `message_*.json` artifacts, and `report_*.pdf` Brand Health Reports).
  - `state/` — Orchestrator-owned state: `cooldowns.json`, `approval_queue.json`, `outreach_log.json`, `runs.json`. Append-only where relevant; never edit historical entries.
- `.claude/skills/` — actual Claude Code skills: `performance-analyser`, `sentiment-analyser`, `contact-discoverer`, `orchestrator`, `message-crafter` (ships with `report_generator.py` for the PDF).

## Skills in this project

- **performance-analyser** — analyzes a brand's recent paid-partnership posts to flag
  underperforming campaigns. Single-brand input. Mode 2 (no brand named) picks a
  random brand from `actionable_contacts.json` for testing. Composite score = weighted
  5-metric formula (views, likes, comments, shares, saves). Output: JSON file in
  `AI Sales Agent System/output/`. See SKILL.md for details.
- **sentiment-analyser** — consumes the latest `perf_analysis_*.json` for one brand,
  classifies comments via MCP, extracts themes via Claude API, outputs a brand-level
  sentiment profile. Strategy 1 (flagged-only) implemented; Strategy 2 (all brand posts)
  feature-flagged off.
- **contact-discoverer** — DNS + scraping pipeline for decision-maker names and likely
  emails. Re-runnable weekly/monthly; per-brand cache. Feeds `actionable_contacts.json`.
- **message-crafter** — drafts the outreach copy for one lead: a <300-char Greek LinkedIn
  note + a 4–6 sentence Greek email (observation→implication→mechanism→CTA), plus an
  optional 5-page Brand Health Report PDF (`report_generator.py`, ReportLab). Copy is
  generated **in-context by Claude** — no Anthropic API key, no LLM Python call; Python
  only loads data, assembles merge values, and renders the PDF. Persona (P1/P2/P3) is
  classified from `job_title`; theme is driven by the Orchestrator's tier (v1 = T1
  comparison for Tier 1/2, neutral-curiosity for Tier 3; T3/T4/T6 are roadmap, blocked on
  video/competitor/profile analysis data the pipeline doesn't yet produce). Mode 1
  (production): Orchestrator calls `Skill(skill="message-crafter", args="<queue_id>")`,
  copy is written back into `approval_queue.json` (`_stub:false`). Mode 2 (testing): brand
  name → finds latest perf+sentiment, writes an audit artifact only, never touches the
  queue. **Drafts only — never sends.** Hard compliance rule: flagged campaigns are
  AI-detected *likely* partnerships, so copy says "πρόσφατη συνεργασία", never "paid
  partnership". The `(Sender_name)` identity in the config is a `TODO_*` placeholder until
  the CEO's details are filled in.
- **orchestrator** — sweeps `actionable_contacts.json`, filters by cooldown, decides
  per-brand stage, invokes perf+sentiment for stale data, assigns tiers, generates stub
  messages, pushes to `state/approval_queue.json`. State machine: NEEDS_ANALYSIS →
  ANALYSIS_READY → IN_QUEUE → IN_COOLDOWN. **Message_Crafter now exists** but stays
  gated: `feature_flags.message_crafter_enabled = false` until the CEO reviews sample
  output and fills the sender identity — while off, the queue gets bland stub copy; when
  flipped on, the Orchestrator enqueues the stub then calls `message-crafter` per
  `queue_id` to overwrite it with real copy. **Scheduling is deferred** — designed runnable, cadence
  decided after first dry-run. Cooldowns advance only on `sent` events; the Orchestrator
  reads cooldowns, a future sender process writes them.

## Design heuristics for this project (locked via /grill-me sessions)

- **Baseline scope**: creator's *paid-partnership* posts only (not all content). The
  per-brand 3-video gate is separate from baseline computation.
- **Division-by-zero in score ratios** → ratio = 1.0 (neutral). Never causes false-
  positive flags.
- **Flagging window**: `0.30 ≤ score < 0.70`. Posts outside this range are silently
  dropped — below 0.30 are anomalous outliers, ≥ 0.70 are adequately performing.
  Priority bands within the window: CRITICAL `0.30–0.40`, HIGH `0.40–0.60`, MEDIUM `0.60–0.70`.
  Config keys: `min_score_to_flag: 0.30`, `underperformance_threshold: 0.70`,
  `priority_bands: { critical_max: 0.40, high_max: 0.60, medium_max: 0.70 }`.
- **Output granularity**: one row per underperforming video (Orchestrator dedupes).
- **Determinism matters**: downstream skills depend on these outputs, so always sort
  inputs/outputs by explicit keys, round floats to 4 decimals.

## Known MCP filter notes

- `brand_queries` requires the original name PLUS at least one realistic variation.
  Single-element arrays are rejected server-side. **Variations must be conservative**:
  capitalization, spacing, punctuation only. NEVER split a multi-word brand name into
  individual word fragments (e.g. don't add `"Monster"` for `"Monster Energy"`) — that
  catches unrelated brands sharing a token (Gentle Monster, Monster Kart, etc.) and
  inflates the candidate set with false positives. Confirmed via the quarantined
  `perf_analysis_monster_energy_2026-06-05.json` incident (2026-06-10).
- `is_ad=true` filter has **OR semantics** — returns posts that are either
  platform-marked as paid partnerships OR have an AI-detected brand at
  `min_likelihood_score`. **Greek-market caveat (confirmed by user 2026-06-10): most
  Greek creators do NOT use the Instagram/TikTok "Paid partnership" disclosure
  feature.** Filtering strictly to `is_ad == true` in the response zeros out most
  real partnerships and produces empty flagged lists. So `is_ad` is unreliable as a
  partnership signal here. Performance_Analyser instead relies on `min_likelihood_score
  >= 8` + `max_brands = 1` + a defensive `is_target_brand` check that the highest-
  likelihood detected brand actually matches the requested one. The trade-off:
  flagged campaigns are "AI-detected likely partnerships," NOT verified. Downstream
  copy should hedge ("we noticed your recent collab" not "your paid partnership").
  Run notes include a `platform_marked / ai_only` split so the human reviewer can see
  which signal each flag came from.
- `is_sponsored=false` is the MCP team's documented recommendation for any
  performance-measurement workflow. Always pass it.
- `min_likelihood_score` default is 8. Strict (9) often returns zero results — use 8
  unless precision over recall is required.
- `has_ai_analysis=false` is dangerous to set: it can exclude all of a brand's
  posts if Renfluence has analyzed them. Default to `null` (don't filter).
- `media_types` MUST exclude `"story"` for any performance-scoring workflow.
  Instagram stories return `view_count=0` (24h expiry, no permanent metrics);
  scoring them against a reel baseline produces a score of 0.0 → CRITICAL flag
  on every story → systemic false-positive class. Performance_Analyser defaults
  to `["video", "reel", "short"]`. Confirmed via the Stoiximan 2026-06-10
  partial run where 5 of the first 10 items were zero-engagement stories.
- `search_media` response `profile_uuid` is nested at `profile_info.profile_uuid`
  (not top-level). Same for `platform`, `nickname`, `influencer.firstname/lastname`.
  Confirmed 2026-06-10. Documented in `performance-analyser` SKILL.md Step A1.5.
- `brand_queries` distinctness check is **case-insensitive on the server side** —
  for single-word Latin brands like `"Stoiximan"`, `["Stoiximan", "stoiximan",
  "STOIXIMAN"]` is rejected as not having "at least one realistic variation."
  Need a script/language variation (e.g. Greek `"Στοίχημαν"`) or a meaningful
  spelling change. The skill's `brand_variations()` auto-generator covers
  capitalization/spacing/punctuation but does NOT produce script transliterations
  — for single-word brands you'll need a per-brand override (manual variation in
  config or `actionable_contacts.json`). Confirmed 2026-06-10 on the Stoiximan run.

## search_media `platform_id` is NOT a stable URL key

`platform_id` shape varies by platform and post type:

| platform_id shape | Real platform/type | URL constructable? |
| --- | --- | --- |
| 11-char alphanumeric (e.g. `DZX3C0eIAIX`) | Instagram reel/post shortcode | ✅ `https://www.instagram.com/reel/<id>/` |
| 19-digit numeric (e.g. `7647114413094341910`) | TikTok video | ✅ if creator handle known: `https://www.tiktok.com/@<handle>/video/<id>` |
| 19-digit numeric (e.g. `3914110757025767391`) | **Instagram story** | ❌ stories have no permalink |
| 11-char alphanumeric | YouTube video | ✅ `https://www.youtube.com/watch?v=<id>` |

**Critical: 19-digit numeric platform_id is AMBIGUOUS between Instagram (story) and
TikTok (video).** Never infer platform from platform_id shape alone — that produces
wrong URLs (e.g. an Instagram story ID rendered as a tiktok.com link 404s). Orchestrator
constructs URLs only when `platform` is explicitly present in the perf output AND
matches the expected shape; falls back to `null` plus the Renfluence internal
`video_uuid` for the CEO to look up manually. Candidate MCP-side fix: a canonical
`media_url` (or `permalink`) field on every `search_media` item.

## linkedin-discoverer notes (confirmed 2026-06-16)

- **geoUrn `urn:li:geographicregion:105072130` (Greece country-level) returns 0 employees.**
  Greek employees set city-level locations (Athens, Thessaloniki), not country-level "Greece".
  The skill's `greece_geo_urn` config is now set to `""` (empty = no filter). These are all
  Greek companies — their employees are predominantly Greek without needing a geo filter.
- **JS extraction requires an external `.js` file** (`.claude/skills/linkedin-discoverer/extract_employees.js`).
  Python's Edit tool inserts Unicode curly quotes when JS is embedded as a string literal,
  which V8 rejects. Never inline multi-line JavaScript in `.py` files for this skill.
- **LinkedIn people page uses `.scaffold-finite-scroll` as scroll container**, not the window.
  `window.scrollTo()` doesn't trigger lazy-load. The skill scrolls `.scaffold-finite-scroll`
  or `.scaffold-layout__content` specifically. Confirmed working: 182 employees scraped across
  15 pages for Stoiximan (Stoiximan has 201-500 employees per LinkedIn).
- **LinkedIn access cap**: free/non-premium accounts see all employees up to ~200 in the
  finite-scroll container. The skill's default `max_pages_per_company = 15` covers this range.
