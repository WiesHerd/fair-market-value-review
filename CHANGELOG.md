# Changelog

## 1.0.0

Initial public release.

### Workflow
- Guided intake that collects survey data, provider details, compensation components,
  productivity, and the proposal one step at a time
- Setup helper that inspects an organization's own workbooks and proposes column
  mappings rather than assuming a layout
- Request register for tracking reviews from intake through committee decision
- Session logging for reproducible reviews

### Analysis
- Sample-size-weighted blending across survey sources, with suppressed sources excluded
- Piecewise-linear percentile interpolation, labeled directional
- Compensation-per-wRVU against derived market ratios
- Compensation-to-productivity alignment check with flags
- Caveat set covering blend overlap, differing survey definitions, publication lag,
  national vs regional data, small samples, and APP billing under a physician identifier

### Output
- Review workbook fill that preserves template formulas and verifies the saved file
- DOCX and PDF review packet with market-band exhibit
- Self-contained HTML email report
- New-hire market range pack
- CSV and JSON audit extracts

### Project
- Synthetic example data so every documented command runs after cloning
- 129 tests
- MIT License
