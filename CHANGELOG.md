# Changelog

All notable changes to this project will be documented in this file.

## [1.0.0] - 2026-08-08

### Added
- **Committee Excel template generator** (`committee_template_generator.py`) — fills committee-request Excel templates with provider data, benchmarks, and blended formulas
- **Adjustment report builder** (`build_adjustment_report.py`) — generates DOCX/PDF reports with 8 sections: data sources, executive summary, cohort detail, market-band exhibit, rationale, cost impact, governance, footnotes
- **CV-only market anchor** (`build_cv_only_market_anchor.py`) — builds market-anchored TCC range from a CV alone when survey data is unavailable
- **Benchmark blending** — weighted-average TCC/wRVU across Survey 1, Survey 3, and Survey 2 by sample size
- **Percentile interpolation** — linear interpolation between survey percentile brackets
- **Configurable benefits rate** — defaults to 4.8% but configurable via config JSON
- **Reference documentation** — committee Excel template structure, survey combined file structure, CV-only market anchor methodology, example report
- **Config template** (`templates/request_config.json`) — all inputs in one JSON file
- **87 pytest tests** with mock Excel/TSV/JSON fixtures — full coverage of core functions
- **GitHub Actions CI** — auto-runs tests on every push/PR
- **Example data** (`examples/`) — fictional providers, anchors, and config so users can try it immediately
- **MIT License**