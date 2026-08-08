# Real-World Walkthrough: A Week in the Life of a Provider Comp Analyst

This walks through how a compensation analyst at a hospital-employed or academic
multi-specialty medical group actually uses this skill, start to finish, using the
fictional data in `examples/`. Every command shown here is real and runnable after
cloning.

---

## The scenario

It's Monday. Three things land in your inbox:

1. **A division chief email:** "My three pediatric critical care attendings are
   getting recruited hard. I need a market adjustment proposal for the comp
   committee by Thursday."
2. **A recruiter forward:** a CV for a potential new PICU hire — no offer drafted
   yet, the department wants to know "what range should we even be talking about?"
3. **A committee coordinator reminder:** the committee's Excel request form for
   one of the adjustments is due, filled out in *their* template, with benchmarks.

You have three source files on your drive (the fictional stand-ins live in `examples/`):

| Your real file | Example stand-in |
|---|---|
| Annual Salary Increases workbook (post-increase salaries, TCC/wRVU percentiles) | `examples/example_salary_file.xlsx` |
| Survey Combined workbook (Survey 1 + Survey 3 + Survey 2 by specialty) | `examples/example_survey_combined.xlsx` |
| Committee request form template (two sheets: Request Form + Tracker) | `examples/example_committee_template.xlsx` |

---

## Task 1 — The cohort adjustment packet (Thursday's committee meeting)

You build a cohort file of the three attendings (name, YOE, current base, planned
base after the annual increase, proposed base). That's `examples/example_cohort.tsv`:

```
Name	YOE	CurrentBase	PlannedBase	ProposedBase
Provider 1	8	275000	287000	325000
Provider 2	12	295000	307000	345000
Provider 3	5	260000	270000	298000
```

Market anchors come from your survey file's blended percentiles → `examples/example_anchors.json`.

One command produces the full packet — DOCX, PDF, market-band exhibit chart, and a CSV audit trail:

```bash
python3 scripts/build_adjustment_report.py --config examples/example_config.json
```

You get an 8-section consulting-style report: data sources, executive summary
(totals, incremental cost vs current *and* vs planned — committees always ask for
both), cohort detail table, the market-band exhibit, empirical rationale,
cost impact with your benefits rate applied, governance guardrails, and footnotes.

**The chief also wants something he can forward.** Generate the email-ready HTML
version — same numbers, same exhibit, KPI cards up top, renders in any client:

```bash
python3 scripts/build_html_email.py --config examples/example_config.json \
  --output examples/adjustment_email.html
```

Open the `.html` in a browser and copy-paste into an email, or attach the PDF
(exported automatically with 0.6" margins if WeasyPrint is installed).

---

## Task 2 — The CV-only range question (recruiter, no offer yet)

You don't have an offer amount or internal equity file — just a CV showing
8 years post-fellowship. You need a defensible range, not a full packet:

```bash
python3 scripts/build_cv_only_market_anchor.py \
  --name "Provider 1" \
  --specialty "Pediatric Critical Care" \
  --yoe 8 \
  --survey-file examples/example_survey_combined.xlsx \
  --output-dir ./anchor_out
```

Output: a 2-page DOCX/PDF with the blended benchmark table, a recommended range
(low ≈ p25, midpoint ≈ p35 directional, high ≈ p40–p50), an exhibit chart, and a
machine-readable `summary.json`. The report leads with a red-flagged caveat that
no internal equity was considered — so nobody mistakes it for a final offer rec.

---

## Task 3 — The committee's Excel form (their template, not yours)

The committee wants *their* two-sheet workbook filled out. This is normally 30-45
minutes of manual lookups: provider row in the salary file, benchmark rows in the
survey file per specialty, weighted-average math, percentile interpolation,
tracker cross-references. One command instead:

```bash
python3 scripts/committee_template_generator.py \
  --name "Provider 1" \
  --salary-file examples/example_salary_file.xlsx \
  --survey-file examples/example_survey_combined.xlsx \
  --template examples/example_committee_template.xlsx \
  --proposed-base 325000 --stipend 42000 --wrvu 4228 --track-num 100026 \
  --no-academic-rank --output committee_request.xlsx
```

What it does automatically:

- Finds the provider in the salary file (YOE, FTE, current/new salary, percentiles, division, specialty)
- Looks up all three survey sources for that specialty — handling the ASCII-hyphen
  vs en-dash mismatch between salary and survey files that silently breaks naive lookups
- Writes n-counts and percentiles into the source rows; the blended weighted-average
  row computes via live formulas (so it updates if a committee member edits a number)
- Writes interpolated-percentile *formulas* (not hardcoded values) for current and
  projected TCC positioning
- Fills the Tracker sheet with cross-sheet references so it always mirrors Sheet 1
- Preserves every template formula (benefits, totals) — verified by re-opening the
  saved file and checking 6 structural assertions before reporting success
- Formats everything for reading: `$316,200` not `316200`, `47.0` not `46.976415094`

By Thursday you've delivered: a styled PDF packet for the committee, an HTML email
for the chief, a defensible range memo for the recruiter, and the committee's own
form — all numerically consistent because they draw from the same source files.

---

## Adapting to your organization

The example files mirror a common layout, but every org's files differ. Adjust:

- **Sheet names / column letters** — constants at the top of
  `committee_template_generator.py` (`SALARY_SHEET_NAME`, `SURVEY_SHEET_NAME`,
  column mappings documented in `references/`)
- **Benefits rate** — read from your template's formula, or set `BENEFITS_RATE`
- **Cell positions in your committee template** — mappings in `fill_template()`,
  documented cell-by-cell in `references/committee-excel-template.md`

Run `python3 -m pytest tests/ -v` after any adaptation — 87+ tests cover the
lookup, blending, interpolation, parsing, and fill logic.
