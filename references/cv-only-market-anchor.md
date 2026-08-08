# CV-Only Market Anchor — Workflow Notes

## When to Use
Use this mode when you have **only a candidate's CV** — no offer amount, no internal
equity cohort file, no existing compensation data. The output is a market-anchored
recommended range, not a full adjustment request.

## What You Need
- Candidate name
- Specialty (must match a specialty in the survey file)
- Years of experience (YOE) — from CV
- Survey Combined file (Excel)

## What You Get
1. `market_anchor_exhibit.png` — horizontal percentile band chart
2. `market_anchor_report.docx` — 2-page market justification
3. `market_anchor_report.pdf` — PDF export (if LibreOffice available)
4. `market_anchor_summary.json` — machine-readable summary

## Recommended Range Logic
- **Low end:** p25 (round to nearest $1,000)
- **Midpoint:** Directional estimate at ~p35 (interpolated between p25 and p50)
- **High end:** ~p40-p50 (60% of the way from p25 to p50)

The rationale: for a new hire at this YOE, targeting the p25-p50 band is typical.
The midpoint is a **directional estimate**, not an exact survey percentile — survey
distributions are not linear between points.

## Critical Caveats
1. **No internal equity:** This analysis does not consider where current incumbents
   sit. The recommended range must be validated against internal equity before use.
2. **No cost impact:** Without an offer amount, no cost impact can be computed.
3. **Interpolation is directional:** Piecewise linear interpolation between survey
   percentile points is an estimate. Actual survey distributions may be non-linear.
4. **Blended benchmark:** Uses weighted-average across available sources
   (Survey 1, Survey 2, Survey 3) by n-count. Sources with suppressed data (n=0)
   are excluded from the blend.

## When NOT to Use
- You have an offer amount → use the full adjustment report instead
- You have an internal equity file → use the committee template generator
- The specialty is not in the survey file → no market anchor is possible