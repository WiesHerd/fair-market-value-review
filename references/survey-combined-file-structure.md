# Survey Combined File Structure

## Overview
The Survey Combined file is an Excel workbook that aggregates benchmark data from
multiple survey sources (Survey 1, Survey 2, Survey 3) into a single sheet,
organized by medical specialty.

## Sheet Name
Common sheet names: `Survey Benchmarks 2025`, `Aggregate`, `Combined Survey`.
Check `wb.sheetnames` if the default doesn't match.

## Structure

### Row 1-3: Headers
- Row 1: Section labels (TCC, Base Salary, wRVUs, TCC per wRVU)
- Row 2: Sub-headers (n, 25th, 50th, 75th, 90th per source)
- Row 3: Source names (Survey 1, Survey 2, Survey 3)

### Rows 4-54: Specialty Data
- Column B: Specialty description (e.g., "Pediatric Critical Care",
  "Emergency Medicine – General", "Hospitalist – Adult")

Each specialty row contains 5 columns per source per metric:
`n, 25th, 50th, 75th, 90th`

### Column Groups (verify against your file)
| Metric | Survey 3 | Survey 1 | Survey 2 |
|--------|------|----------------|-----------|
| TCC | 13-17 | 33-37 | 63-67 |
| Base | 18-22 | 38-42 | 68-72 |
| TCC/wRVU | 23-27 | 53-57 | 78-82 |
| wRVU | 28-32 | 58-62 | 83-87 |

## Specialty Matching

### Dash Normalization (CRITICAL PITFALL)
Salary files often use ASCII hyphen `-` (e.g., "Pediatrics - Gynecology") while
survey files may use Unicode en-dash `–` (U+2013) (e.g., "Pediatrics – Gynecology").

**Fix:** Normalize via `unicodedata.normalize('NFKD')` + ASCII encoding before comparison. Note: NFKD does *not* turn an en-dash into a hyphen — it strips it entirely (leaving a double space), so a direct `norm(a) == norm(b)` check still fails across dash variants. The actual fix is the fuzzy split-and-substring match below, not normalization alone.

### Fuzzy Matching
Exact match may fail due to:
- Extra whitespace
- Different ordering ("Pediatric Cardiology" vs "Cardiology, Pediatric")
- Dash variants

**Fuzzy match strategy:** Split on dash, check that ALL parts are present in the
target specialty string. Do NOT match on the first part alone — "Pediatrics" alone
matches "Pediatrics – General" (wrong specialty row).

```python
def norm(s):
    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().strip().lower()

parts = [p.strip() for p in norm(specialty).split('-') if p.strip()]
# Match only if ALL parts are found in the target
all(p in norm(target) for p in parts)
```

## Suppressed Data
Survey values may be suppressed (too few respondents):
- `*` → suppressed (n too small)
- `None` / empty → no data
- `0` → no data (not a real zero)

**Filter automatically:** treat `*`, `None`, `0` as "no data" (n=0, percentiles=None).

## Verified Example
```
Specialty: "Pediatric Critical Care"
Sheet: "Survey Benchmarks 2025"

Survey 3 TCC:     n=142, p25=$280K, p50=$320K, p75=$365K, p90=$410K
SC TCC:       n=89,  p25=$285K, p50=$325K, p75=$370K, p90=$420K
Survey 2:    n=34,  p25=$275K, p50=$315K, p75=$360K, p90=$405K

Blended TCC (weighted by n):
  p25 = (280*142 + 285*89 + 275*34) / (142+89+34) = $281,038
  p50 = (320*142 + 325*89 + 315*34) / 265 = $321,038
  p75 = (365*142 + 370*89 + 360*34) / 265 = $366,038
  p90 = (410*142 + 420*89 + 405*34) / 265 = $412,717
```

These figures match the values produced by `scripts/generate_example_xlsx_fixtures.py`
and the workbook filler, checked with LibreOffice formula recalculation.