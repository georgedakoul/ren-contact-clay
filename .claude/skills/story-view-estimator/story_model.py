#!/usr/bin/env python3
"""Story-view estimator: fit a per-follower-tier ratio model on platform creators
(shared_insights=true, real story view_count known) and predict avg story views for
creators where stories return 0 (shared_insights=false / off-platform).

Model: story_views ~= median_ratio[tier] * base_value, where base is followers OR
avg feed/reel views. Median (not mean) for outlier robustness; IQR of ratios gives a
low/high band. Two bases are fitted and compared by leave-one-out MAPE; the lower one
is recommended.

ponytail: median-ratio-per-tier. Upgrade to log-log OLS or add engagement_rate as a
second factor only if holdout MAPE stays bad (> ~35%).

CLI:
  python story_model.py calibrate <rows.json> <coeffs_out.json>
  python story_model.py predict   <coeffs.json> --followers N [--feed-views N] [--base auto|followers|feed_views]
  python story_model.py selftest
"""
import argparse
import json
import statistics as st
import sys
from pathlib import Path

# (label, lo_inclusive, hi_exclusive) — edit in config, mirrored here as the default.
DEFAULT_TIERS = [
    ["nano", 0, 10_000],
    ["micro", 10_000, 50_000],
    ["mid", 50_000, 200_000],
    ["macro", 200_000, 1_000_000],
    ["mega", 1_000_000, 10**12],
]
# base name -> row key holding that base value.
# "engagement" = avg feed engagement/post (likes+comments+shares+saves); proxy for the
# creator's ACTIVE audience, which story viewers track better than raw follower count.
# (active_audience = followers * engagement_rate is algebraically == avg_engagement, so not separate.)
BASE_KEYS = {"followers": "followers", "feed_views": "avg_feed_views", "engagement": "avg_engagement"}


def tier_of(followers, tiers):
    for label, lo, hi in tiers:
        if lo <= followers < hi:
            return label
    return tiers[-1][0]  # above last band -> top tier


def _quartiles(vals):
    """median, q1, q3. Robust for tiny n (q1=q3=median when n<2)."""
    s = sorted(vals)
    med = st.median(s)
    if len(s) < 2:
        return med, med, med
    # inclusive median method -> q1/q3 from halves
    mid = len(s) // 2
    lower = s[:mid]
    upper = s[mid + 1:] if len(s) % 2 else s[mid:]
    q1 = st.median(lower) if lower else med
    q3 = st.median(upper) if upper else med
    return med, q1, q3


def _ratios_by_tier(rows, tiers, base_key):
    """{tier: [ratio,...]} for rows with a positive base value and positive story views."""
    out = {t[0]: [] for t in tiers}
    for r in rows:
        base = r.get(base_key)
        sv = r.get("real_story_views")
        if not base or base <= 0 or sv is None or sv < 0:
            continue
        out[tier_of(r["followers"], tiers)].append(sv / base)
    return out


def fit_one_base(rows, tiers, base):
    base_key = BASE_KEYS[base]
    by_tier = _ratios_by_tier(rows, tiers, base_key)
    all_ratios = [x for v in by_tier.values() for x in v]
    global_med = round(st.median(all_ratios), 6) if all_ratios else None
    coeffs = {}
    for label, rs in by_tier.items():
        if rs:
            med, q1, q3 = _quartiles(rs)
            coeffs[label] = {"n": len(rs), "ratio": round(med, 6),
                             "ratio_low": round(q1, 6), "ratio_high": round(q3, 6)}
        else:
            # empty tier -> borrow global median so prediction never crashes
            coeffs[label] = {"n": 0, "ratio": global_med,
                             "ratio_low": global_med, "ratio_high": global_med,
                             "borrowed_global": True}
    return {"base": base, "global_ratio": global_med, "tiers": coeffs}


def loo_mape(rows, tiers, base):
    """Leave-one-out MAPE: predict each row from the median ratio of OTHER rows in its
    tier. Rows whose tier has <2 usable samples are skipped (can't hold out)."""
    base_key = BASE_KEYS[base]
    usable = [r for r in rows if r.get(base_key, 0) > 0 and (r.get("real_story_views") or 0) > 0]
    errs = []
    for i, r in enumerate(usable):
        t = tier_of(r["followers"], tiers)
        others = [o[base_key] and (o["real_story_views"] / o[base_key])
                  for j, o in enumerate(usable)
                  if j != i and tier_of(o["followers"], tiers) == t]
        others = [x for x in others if x]
        if not others:
            continue
        pred = st.median(others) * r[base_key]
        errs.append(abs(pred - r["real_story_views"]) / r["real_story_views"])
    return round(100 * sum(errs) / len(errs), 2) if errs else None, len(errs)


def calibrate(rows, tiers):
    result = {"n_rows": len(rows), "tiers": tiers, "models": {}, "mape": {}}
    for base in BASE_KEYS:
        result["models"][base] = fit_one_base(rows, tiers, base)
        m, n = loo_mape(rows, tiers, base)
        result["mape"][base] = {"loo_mape_pct": m, "n_scored": n}
    # recommend the base with the lower (non-null) MAPE
    scored = {b: v["loo_mape_pct"] for b, v in result["mape"].items()
              if v["loo_mape_pct"] is not None}
    result["recommended_base"] = min(scored, key=scored.get) if scored else "followers"
    return result


def predict(cal, followers, features=None, base="auto"):
    """features: optional {"feed_views": N, "engagement": N}. followers always available
    and used for tiering. Falls back to the followers base if the chosen base value is missing."""
    bands = cal["tiers"]
    features = features or {}
    if base == "auto":
        base = cal.get("recommended_base", "followers")
    base_val = followers if base == "followers" else features.get(base)
    if not base_val:
        base = "followers"  # fall back if the alt base value wasn't supplied
        base_val = followers
    t = tier_of(followers, bands)
    m = cal["models"][base]["tiers"][t]
    if m["ratio"] is None or not base_val:
        return {"error": "no ratio available for this tier/base", "tier": t}
    return {
        "tier": t, "base": base, "base_value": base_val,
        "estimate": round(m["ratio"] * base_val),
        "low": round(m["ratio_low"] * base_val),
        "high": round(m["ratio_high"] * base_val),
        "ratio_used": m["ratio"], "tier_sample_n": m["n"],
    }


def selftest():
    import random
    random.seed(7)
    true_ratio = {"nano": 0.25, "micro": 0.15, "mid": 0.09, "macro": 0.04, "mega": 0.02}
    rows = []
    for label, lo, hi in DEFAULT_TIERS:
        hi = min(hi, lo * 8 + 5000)
        for _ in range(15):
            f = random.randint(lo + 1, hi)
            sv = true_ratio[label] * f * random.uniform(0.7, 1.3)  # +-30% noise
            rows.append({"followers": f, "avg_feed_views": f * random.uniform(2, 6),
                         "real_story_views": round(sv)})
    cal = calibrate(rows, DEFAULT_TIERS)
    # recovered follower-ratio per tier within 20% of truth
    for label in true_ratio:
        got = cal["models"]["followers"]["tiers"][label]["ratio"]
        assert abs(got - true_ratio[label]) / true_ratio[label] < 0.20, (label, got)
    # followers must win (story views were generated from followers, not feed views)
    assert cal["recommended_base"] == "followers", cal["recommended_base"]
    # holdout error in the noise ballpark
    assert cal["mape"]["followers"]["loo_mape_pct"] < 25, cal["mape"]
    # a prediction lands near truth
    p = predict(cal, followers=120_000, base="followers")
    exp = true_ratio["mid"] * 120_000
    assert abs(p["estimate"] - exp) / exp < 0.30, p
    print("selftest OK", {"mape_followers": cal["mape"]["followers"]["loo_mape_pct"],
                          "mape_feed": cal["mape"]["feed_views"]["loo_mape_pct"]})


def _load_tiers(rows_obj):
    return [list(t) for t in rows_obj.get("tiers", DEFAULT_TIERS)]


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("calibrate"); c.add_argument("rows"); c.add_argument("out")
    p = sub.add_parser("predict"); p.add_argument("coeffs")
    p.add_argument("--followers", type=int, required=True)
    p.add_argument("--feed-views", type=int, default=None)
    p.add_argument("--avg-engagement", type=int, default=None)
    p.add_argument("--base", default="auto", choices=["auto", "followers", "feed_views", "engagement"])
    sub.add_parser("selftest")
    a = ap.parse_args()

    if a.cmd == "selftest":
        selftest()
    elif a.cmd == "calibrate":
        obj = json.loads(Path(a.rows).read_text())
        rows = obj["rows"] if isinstance(obj, dict) else obj
        tiers = _load_tiers(obj) if isinstance(obj, dict) else DEFAULT_TIERS
        cal = calibrate(rows, tiers)
        Path(a.out).write_text(json.dumps(cal, indent=2))
        print(json.dumps({"recommended_base": cal["recommended_base"],
                          "mape": cal["mape"], "n_rows": cal["n_rows"]}, indent=2))
    elif a.cmd == "predict":
        cal = json.loads(Path(a.coeffs).read_text())
        feats = {"feed_views": a.feed_views, "engagement": a.avg_engagement}
        print(json.dumps(predict(cal, a.followers, feats, a.base), indent=2))


if __name__ == "__main__":
    main()
