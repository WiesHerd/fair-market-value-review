# Committee Excel Template — Cell Mapping & Formula Reference

## Overview
This document describes the generic structure of a two-sheet committee compensation
request Excel template. **Cell positions vary by organization** — use this as a guide,
not gospel. Always verify against your actual template.

## Sheet 1: "FMV Review"

### Header Section (rows 4-10)
| Cell | Content | Notes |
|------|---------|-------|
| B4   | Provider name | "First Last, MD" |
| F4   | Specialty | From salary file |
| B5   | Academic rank (optional) | May be blank |
| F5   | YOE | Years post-fellowship/credentialing |
| B6/C6| Request type | "X" in B6=Existing, C6=Incremental New |
| B8   | FTE | Full-time equivalent (0.0–1.0) |
| C10-F10 | CART distribution | Clinical / Admin / Research / Teaching (must sum to 1.0) |

### Current Compensation (row 13)
| Cell | Content | Formula? |
|------|---------|---------|
| C13  | Current base salary | Value (from salary file, post-annual-increase) |
| D13  | Current stipend | Value |
| E13  | Current production incentive | Value |
| F13  | Other income | Value |
| G13  | Total | `=SUM(C13:F13)` (preserved) |
| H13  | Benefits | `=G13 * benefits_rate` (preserved — do NOT change) |
| I13  | Total + Benefits | `=G13+H13` (preserved) |

### Current Metrics (row 15-16)
| Cell | Content |
|------|---------|
| F15  | total annual wRVUs (sum of work RVUs produced) |
| B16  | Current TCC percentile (interpolated formula — see below) |
| F16  | Current wRVU percentile (interpolated formula — see below) |

### Request Summary (row 20)
| Cell | Content |
|------|---------|
| A20  | Request summary narrative |

### Proposed Compensation (row 23)
| Cell | Content |
|------|---------|
| C23  | Proposed base salary |
| D23  | Proposed stipend |
| E23  | Proposed production incentive |
| F23  | Other income |
| G23  | Total | `=SUM(C23:F23)` (preserved) |
| H23  | Benefits | `=G23 * benefits_rate` (preserved) |
| I23  | Total + Benefits | `=G23+H23` (preserved) |
| C26  | Projected TCC percentile (interpolated formula) |

### TCC Benchmarks (rows 31-34)
| Row | Source | Cols B-G |
|-----|--------|----------|
| 31  | Survey 1 | n, blank, p25, p50, p75, p90 |
| 32  | Survey 2      | n, blank, p25, p50, p75, p90 |
| 33  | Survey 3           | n, blank, p25, p50, p75, p90 |
| 34  | **Blended**    | `=SUM(B31:B33)`, IFERROR weighted-avg formulas |

### wRVU Benchmarks (rows 37-40)
Same structure as TCC, rows 37-40.

### Background (rows 43-46)
| Cell | Content |
|------|---------|
| A43  | General background narrative |
| A45  | "Provider Background:" label |
| A46  | Provider-specific background |

## Sheet 2: "Tracker"

Row 6 (or next empty row) contains cross-sheet references to Sheet 1.

Key columns:
- F6: Provider name
- G6: Specialty
- I6: YOE
- K6: FTE
- Q6-S6: Current comp (base, stipend, prod incentive)
- V6: `='FMV Review'!B20` (TCC percentile — cross-sheet ref)
- X6: `='FMV Review'!E21` (wRVU percentile — cross-sheet ref)
- AA6-AC6: Proposed comp
- AK6-AO6: % change formulas (IFERROR-wrapped)
- AP6: `='FMV Review'!C26` (projected percentile)
- AR6-BC6: Cross-sheet refs to blended benchmark rows (34 and 40)

## Interpolated Percentile Formulas

Instead of hardcoding percentile values, use formulas that auto-calculate from
the blended benchmark row. These update automatically if benchmarks change:

### TCC Percentile (cell B16)
```excel
=IFERROR(
  IF(G13<=$E$34,
    25+(G13-$D$34)/($E$34-$D$34)*25,
    IF(G13<=$F$34,
      50+(G13-$E$34)/($F$34-$E$34)*25,
      75+(G13-$F$34)/($G$34-$F$34)*15
    )
  ),
  ""
)
```

### wRVU Percentile (cell F16)
Same structure but referencing row 40 instead of 34, and F15 instead of G13.

## Encrypted (.dec) File Handling

Some organizations encrypt Excel templates with a `.dec` extension (portal encryption).
The generator scripts handle this transparently:

1. Copy `.dec` file to a temporary `.xlsx` file.
2. Load with openpyxl.
3. Save output as `.xlsx` (the `.dec` encryption is not reapplied).

## Benefits Rate

The benefits/fringe rate varies by organization and is typically somewhere in the
4-7% range depending on how the org defines "benefits" for this calculation
(fringe-only vs. fully-loaded). Confirm the exact figure against your own
template's formula rather than assuming a default.

The rate is typically embedded in the template's benefits formula (e.g., `=G13*0.048`).
**Do not change the formula** — just set the base salary and let the formula calculate.

If the template doesn't have a benefits formula, set `BENEFITS_RATE` in the script
or pass it via the config file.