# Example Data

Fictional data for trying the skill's scripts. All names are invented.

## Files

| File | Description |
|------|-------------|
| `example_cohort.tsv` | 3 fictional providers (Avery, Morgan, Patel) with YOE, current/planned/proposed base |
| `example_anchors.json` | p25/p50/p75/p90 TCC market anchors |
| `example_config.json` | Config file pointing at the example data (for `build_adjustment_report.py --config`) |
| `example_salary_file.xlsx` | Fake "Annual Salary Increases" workbook (3 providers) -- for the committee template generator |
| `example_survey_combined.xlsx` | Fake "Survey Combined" workbook (2 specialties, blended TCC/wRVU by Survey 1/Survey 3/Survey 2) |
| `example_fmv_template.xlsx` | Fake FMV review workbook with formulas intact |

Regenerate the xlsx fixtures anytime with `python3 scripts/generate_example_xlsx_fixtures.py`.

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Build a DOCX adjustment report from the example cohort
python scripts/build_adjustment_report.py \
  --name "Pediatric Critical Care Faculty" \
  --providers-file examples/example_cohort.tsv \
  --market-anchors-file examples/example_anchors.json \
  --output examples/example_report.docx \
  --no-pdf

# Build a CV-only market anchor
# NOTE: this script auto-looks-up benchmarks from a Survey Combined workbook by
# specialty (--survey-file) or a pasted benchmark text file (--benchmarks) -- it
# does not take --current-base/--anchors-file/--output (those flags don't exist).
python scripts/build_cv_only_market_anchor.py \
  --name "Provider 1" \
  --specialty "Pediatric Critical Care" \
  --yoe 8 \
  --survey-file examples/example_survey_combined.xlsx \
  --output-dir ./examples/cv_anchor_out

# Committee Excel template fill (fills both sheets, all formulas preserved)
python scripts/fmv_workbook_generator.py \
  --name "Provider 1" \
  --salary-file examples/example_salary_file.xlsx \
  --survey-file examples/example_survey_combined.xlsx \
  --template examples/example_fmv_template.xlsx \
  --proposed-base 325000 --stipend 42000 --wrvu 4228 --track-num 100026 \
  --no-academic-rank --output examples/committee_request_output.xlsx
```

## FMV Review Workbook

The committee template generator requires a salary workbook and survey workbook in a specific format.
`example_salary_file.xlsx` and `example_survey_combined.xlsx` above follow that format and can be used
directly (see the command above). See `references/fmv-workbook-cell-map.md` and
`references/survey-combined-file-structure.md` for the expected column layouts if you need to adapt this
to your organization's actual files.