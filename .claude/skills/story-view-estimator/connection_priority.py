#!/usr/bin/env python3
"""Connection-priority ranking (step B): for off-platform creators we CAN'T measure story
views on, rank who to push hardest to connect / share insights — because connecting them is
the only path to real story numbers (estimation caps at ~74% MAPE, see SKILL.md).

Priority = estimated story reach (what we're currently blind to) × a strategic weight, with
the estimate's low-confidence flagged. Big estimated reach + we have no real data = chase first.

Input JSON: [{"handle","followers","niche"(1 personal/0 promo),"feed_views"?,"weight"?}]
  weight: optional strategic multiplier (e.g. brand fit / deal size), default 1.0.

Usage:
  python connection_priority.py <coeffs.json> <creators.json> [out.json]
  python connection_priority.py selftest
"""
import json
import sys
from pathlib import Path

import story_model as sm


def rank(coeffs, creators):
    out = []
    for c in creators:
        feats = {"feed_views": c.get("feed_views"), "niche": c.get("niche")}
        base = "powerlaw_niche" if coeffs.get("powerlaw_niche") and c.get("niche") is not None else "auto"
        p = sm.predict(coeffs, c["followers"], feats, base)
        est = p.get("estimate")
        if est is None:
            continue
        weight = float(c.get("weight", 1.0))
        out.append({
            "handle": c.get("handle"),
            "followers": c["followers"],
            "niche": c.get("niche"),
            "est_story_views": est,
            "est_low": p.get("low"), "est_high": p.get("high"),
            "model": p.get("base"),
            "confidence": "low",  # estimation ceiling; real number needs connection
            "weight": weight,
            # priority: estimated reach we currently can't see, scaled by strategic weight.
            "priority_score": round(est * weight),
        })
    out.sort(key=lambda r: r["priority_score"], reverse=True)
    for i, r in enumerate(out, 1):
        r["rank"] = i
    return out


def selftest():
    coeffs = json.loads((Path(__file__).parent.parent.parent.parent
                         / "AI Sales Agent System/output/story-view-estimator/story_coeffs.json").read_text())
    creators = [
        {"handle": "big_personal", "followers": 300000, "niche": 1},
        {"handle": "small_promo", "followers": 20000, "niche": 0},
        {"handle": "mid_personal_weighted", "followers": 120000, "niche": 1, "weight": 2.0},
    ]
    r = rank(coeffs, creators)
    assert [x["handle"] for x in r][0] == "big_personal", r  # biggest reach ranks first
    assert r[-1]["handle"] == "small_promo", r
    assert all(x["confidence"] == "low" for x in r)
    assert r[0]["rank"] == 1
    print("selftest OK:", [(x["rank"], x["handle"], x["priority_score"]) for x in r])


def main():
    if len(sys.argv) == 2 and sys.argv[1] == "selftest":
        return selftest()
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    coeffs = json.loads(Path(sys.argv[1]).read_text())
    creators = json.loads(Path(sys.argv[2]).read_text())
    ranked = rank(coeffs, creators)
    out_path = sys.argv[3] if len(sys.argv) > 3 else None
    if out_path:
        Path(out_path).write_text(json.dumps(ranked, indent=2, ensure_ascii=False))
    # compact table to stdout
    print(f"{'#':>2}  {'handle':22}  {'followers':>9}  {'niche':5}  {'est_views':>9}  {'band':>17}")
    for r in ranked:
        band = f"{r['est_low']}-{r['est_high']}"
        print(f"{r['rank']:>2}  {(r['handle'] or '')[:22]:22}  {r['followers']:>9}  "
              f"{'pers' if r['niche'] else 'promo':5}  {r['est_story_views']:>9}  {band:>17}")


if __name__ == "__main__":
    main()
