---
name: story-view-estimator
description: Estimate the average Instagram story view count for a creator whose stories are private to Renfluence (shared_insights=false / off-platform), using a tier-based ratio model calibrated on platform creators (shared_insights=true) whose real story views ARE visible in the MCP. Use when the user asks to estimate/predict story views for a creator, calibrate or recalibrate the story model, or run a comparison test of predicted vs real story views. Two modes: calibrate (fit on platform creators) and predict (apply to a given creator's public metrics).
---

# Story-View Estimator

## Why this exists

Renfluence sees **real story view counts only for creators who shared insights** via the
official platform API (`search_profiles(shared_insights=true)` — these stories return a
real `view_count` from `search_media(media_types=["story"])`). For everyone else (off-platform
or not connected), stories return `view_count=0`. This skill fits a model on the creators we
*can* see and applies it to the ones we can't.

**Model**: `story_views ≈ median_ratio[follower_tier] × base`, where `base` is follower count,
avg feed/reel views, or avg feed engagement. Median ratio per tier (outlier-robust); IQR of the
tier's ratios gives a low/high band. The math + a self-test live in `story_model.py`.

> ### ⚠️ Accuracy reality (measured 2026-06-29, 26-creator Greek IG set)
> **Story views are only weakly predictable from public metrics.** This is a low-confidence
> estimator — treat its output as an order-of-magnitude range, never a precise number.
> - story-views / followers ratio spans **1.2%–85%** across creators (70× spread).
> - Leave-one-out MAPE: followers **123%**, feed_views **154%**, engagement **178%** (engagement
>   does NOT help). A power-law (log-log OLS) on feed_views is the best variant at **~87%** MAPE.
> - Best correlate: avg feed views, log-log r≈0.71. Still not decision-grade.
> - **Root cause**: story reach is a creator-behaviour signal (how "personal"/story-active the
>   account is — lifestyle/mum/personal creators run 15–85%, aggregator/aesthetic/music accounts
>   ~1–4%), largely decoupled from any public feed count. A numeric-only model can't see it.
>
> **Honest recommendation**: for an off-platform creator, report a wide tier range + "low
> confidence, story reach varies heavily by creator type", not a point estimate. The only exact
> source is Renfluence's own data (shared_insights=true). Don't oversell this number to the CEO.

> **MCP-side note (surface to the team):** the cleanest home for this is server-side. Renfluence
> already has real story views for the whole `shared_insights=true` population — far more than a
> client run can paginate. Recommend an `estimated_story_views` (+ `confidence`, `n`) field on
> `get_profile_overviews` for non-insight profiles, refit server-side on a schedule. This skill is
> the prototype that proves the formula before porting the coefficients server-side.

## Config

Read `AI Sales Agent System/story_estimator_config.json` (defaults + overrides deep-merge, leaf
wins; spoken per-run overrides apply for that run only, not written back). Referenced as `cfg.*`.

---

## Mode A — Calibrate (fit / recalibrate the model)

Run when there are no coeffs yet, they're older than `cfg.recalibrate_after_days`, or the user
says "recalibrate" / "rebuild the story model".

### A1. Gather platform creators
`search_profiles(shared_insights=true, platforms=cfg.calibration.platforms,
countries=cfg.calibration.countries, sort_by="followers", page_size=10)`, paginating until you
have `cfg.calibration.target_sample_size` creators **spread across follower tiers** (don't take
only mega accounts — page through follower bands so every tier in `cfg.tiers` gets samples; use
`min_followers`/`max_followers` to target thin tiers). Keep `follower_count` and, when present,
`normal_stats.avg_views` (= avg feed/reel views) for each.

### A2. Pull real story views per creator
For each creator: `search_media(profile_uuid=<uuid>, media_types=["story"],
order_by="created_time", page_size=cfg.calibration.stories_lookback_page_size)`.
- Collect each story's `view_count`.
- If `cfg.calibration.exclude_zero_view_stories`: drop `view_count==0` (image stories / not-yet-counted).
- Skip the creator if fewer than `cfg.calibration.min_stories_per_creator` usable stories remain
  (too noisy to trust).
- `real_story_views` = **median** of the kept story view counts. If `drop_story_view_outliers_iqr`,
  drop story values outside `[Q1-1.5·IQR, Q3+1.5·IQR]` before taking the median.

### A3. Build the rows file and fit
Write `rows.json`: `{"tiers": cfg.tiers, "rows": [{"followers", "avg_feed_views",
"real_story_views"}, ...]}` (omit `avg_feed_views` when the creator had no `normal_stats`).
Then:
```
python story_model.py calibrate rows.json <cfg.coeffs_file>
```
It writes the coeffs and prints `recommended_base` + leave-one-out MAPE for both bases.

### A4. Report
Tell the user: sample size, per-tier sample counts, the recommended base, and the LOO-MAPE.
If MAPE > ~35%, say so and flag the upgrade path (add engagement_rate factor or log-log OLS —
see the `ponytail:` note in `story_model.py`). Keep `rows.json` in the output dir as the audit trail.

---

## Mode B — Predict (estimate for one creator)

Run when the user gives a creator (by name/handle, by `profile_uuid`, or by raw public metrics)
and wants an estimated avg story view count.

### B1. Get the creator's metrics
- If given raw numbers (followers, recent post views), use them directly.
- If given a handle/uuid that IS in the MCP: `search_profiles(username=...)` or
  `get_profile_overviews([uuid])` → take `follower_count` and `avg_views`. If
  `shared_insights==true`, tell the user the **real** story views are available directly
  (`search_media` stories) — no estimate needed.

### B2. Predict
```
python story_model.py predict <cfg.coeffs_file> --followers <N> [--feed-views <N>]
```
Output: `estimate`, `low`, `high` (band), tier used, and the tier's sample size.

### B3. Report
Give the point estimate with the low–high band and the tier it fell in. State plainly it's an
**estimate** from a follower-tier ratio model, and weaker for tiers with small `tier_sample_n`.
If coeffs are missing → run Mode A first.

---

## Notes
- `story_model.py selftest` runs a synthetic-data check (asserts ratio recovery + base selection).
- Determinism: medians/quartiles are order-independent; floats rounded in the coeffs file.
- Off-platform creators the user supplies by hand may not be in the MCP at all — that's fine, the
  model only needs follower count (+ optional feed views) as input, from any source.
