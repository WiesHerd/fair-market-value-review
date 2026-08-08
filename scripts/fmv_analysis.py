#!/usr/bin/env python3
"""
FMV analytics -- the tests a compensation valuation analyst actually runs.

Grounded in published FMV guidance (PYA "Demystifying Fair Market Value
Compensation"; VMG Health "Debunking Provider Compensation Myths"; CMS Stark
commentary). Three ideas drive this module:

1. ALIGNMENT, NOT PERCENTILE LEVEL. There is no percentile that is automatically
   FMV, and none that is automatically not. CMS has never set a percentile
   threshold. What matters is whether compensation aligns with productivity and
   the facts of the arrangement.

2. SURVEYS COMPUTE EACH METRIC INDEPENDENTLY. A provider at the 90th percentile
   of wRVUs is not thereby entitled to the 90th percentile of compensation per
   wRVU. Market comp-per-wRVU stays relatively flat across quartiles, so applying
   a high-quartile ratio to high production double-counts productivity.

3. RATIOS ARE CONTEXT-DEPENDENT. Comp-to-collections above 100% is normal for
   coverage-based and high-Medicaid specialties and is not itself an FMV problem.

Nothing here produces an FMV opinion -- it produces the market-approach analysis
and the flags a reviewer should look at.
"""
from typing import Optional, Dict, List


def interpolate_percentile(value, p25, p50, p75, p90) -> Optional[float]:
    """Piecewise-linear percentile position. Directional only: survey
    distributions are not linear between published points."""
    if value is None or not all(v is not None and v > 0 for v in (p25, p50, p75, p90)):
        return None
    if value <= p25:
        return 25.0 * (value / p25) if p25 else 0.0
    if value <= p50:
        return 25.0 + 25.0 * (value - p25) / (p50 - p25)
    if value <= p75:
        return 50.0 + 25.0 * (value - p50) / (p75 - p50)
    if value <= p90:
        return 75.0 + 15.0 * (value - p75) / (p90 - p75)
    return 90.0 + 10.0 * (value - p90) / p90


def comp_per_wrvu(tcc, wrvus) -> Optional[float]:
    if not tcc or not wrvus:
        return None
    return tcc / wrvus


def market_comp_per_wrvu(tcc_bench: Dict, wrvu_bench: Dict) -> Dict:
    """Derive market comp-per-wRVU at each percentile. This is a *derived* ratio;
    if the survey publishes its own comp-per-wRVU table, prefer that -- surveys
    compute each metric independently and a derived ratio is a different statistic."""
    out = {}
    for p in ('25', '50', '75', '90'):
        t, w = tcc_bench.get(p), wrvu_bench.get(p)
        out[p] = (t / w) if (t and w) else None
    return out


def alignment_check(tcc_pctile, wrvu_pctile, tolerance: float = 15.0) -> Dict:
    """Compare compensation positioning against productivity positioning.

    A gap wider than `tolerance` percentile points is the classic FMV flag. This
    is a prompt to document *why*, not a verdict -- guarantees during ramp-up,
    coverage requirements, administrative duties and specialty supply all
    legitimately explain gaps.
    """
    if tcc_pctile is None or wrvu_pctile is None:
        return {'status': 'insufficient_data', 'gap': None,
                'message': 'Alignment could not be assessed (missing TCC or wRVU percentile).'}
    gap = tcc_pctile - wrvu_pctile
    if abs(gap) <= tolerance:
        return {'status': 'aligned', 'gap': gap,
                'message': (f"Compensation (p{tcc_pctile:.0f}) and productivity "
                            f"(p{wrvu_pctile:.0f}) are aligned within {tolerance:.0f} "
                            f"percentile points.")}
    if gap > 0:
        return {'status': 'comp_above_production', 'gap': gap,
                'message': (f"Compensation (p{tcc_pctile:.0f}) sits {gap:.0f} percentile "
                            f"points above productivity (p{wrvu_pctile:.0f}). Document the "
                            f"basis -- e.g. income guarantee during ramp-up, coverage/call "
                            f"burden, administrative or teaching duties, specialty supply "
                            f"constraints, or a payer mix that suppresses wRVUs.")}
    return {'status': 'production_above_comp', 'gap': gap,
            'message': (f"Productivity (p{wrvu_pctile:.0f}) exceeds compensation positioning "
                        f"(p{tcc_pctile:.0f}) by {abs(gap):.0f} percentile points -- a "
                        f"retention risk and a supportable basis for adjustment.")}


def percentile_caveats(tcc_pctile) -> List[str]:
    notes = [
        "No percentile is inherently fair market value. CMS has not set a percentile "
        "threshold above or below which compensation is or is not FMV; each arrangement "
        "turns on its own facts and circumstances.",
        "Percentile positioning is one input to a market-approach analysis, not a "
        "conclusion of value on its own.",
    ]
    if tcc_pctile is not None and tcc_pctile >= 75:
        notes.append(
            "Positioning at or above the 75th percentile is not disqualifying -- roughly "
            "25% of surveyed providers sit there by definition -- but it warrants explicit "
            "documentation of the productivity, scope, or market factors that support it.")
    if tcc_pctile is not None and tcc_pctile >= 90:
        notes.append(
            "Positioning at or above the 90th percentile can still be supportable (high "
            "production, stacked call/administrative roles), but expect scrutiny and "
            "document each stacked component separately as well as in aggregate.")
    return notes


def survey_use_caveats(n_sources: int, combined_n=None) -> List[str]:
    notes = []
    if n_sources >= 2:
        notes.append(
            "Multiple independently published surveys were referenced, which is the prudent "
            "practice cited in CMS Stark commentary. Note that blending sources can "
            "double-count providers who respond to more than one survey.")
    if n_sources == 1:
        notes.append(
            "Only one survey source carried data for this specialty. Referencing multiple "
            "objective surveys is preferred where sample sizes permit.")
    notes.append(
        "Surveys define total cash compensation differently and are typically published one "
        "to two years in arrears. Confirm definitions match the components being compared, "
        "and consider whether market movement since the survey date is material.")
    notes.append(
        "National data is generally more reliable than survey 'regional' cuts, which mix "
        "dissimilar markets and differ in composition between publishers. Adjust for local "
        "market factors explicitly rather than relying on a regional table.")
    if combined_n is not None and combined_n < 30:
        notes.append(
            f"Combined sample size is small (n={combined_n}). A single respondent can move "
            f"the benchmark materially; treat percentile positioning as indicative only.")
    notes.append(
        "Verify the productivity comparison is apples-to-apples -- wRVUs billed by advanced "
        "practice providers under a physician's identifier will overstate that physician's "
        "production relative to survey benchmarks.")
    return notes


def build_fmv_summary(tcc, wrvus, tcc_bench: Dict, wrvu_bench: Dict = None,
                      n_sources: int = 1, combined_n=None) -> Dict:
    """Full market-approach summary: positioning, ratio, alignment, caveats."""
    tcc_p = interpolate_percentile(tcc, tcc_bench.get('25'), tcc_bench.get('50'),
                                   tcc_bench.get('75'), tcc_bench.get('90'))
    wrvu_p = ratio = ratio_p = None
    market_ratio = {}
    if wrvus and wrvu_bench:
        wrvu_p = interpolate_percentile(wrvus, wrvu_bench.get('25'), wrvu_bench.get('50'),
                                        wrvu_bench.get('75'), wrvu_bench.get('90'))
        ratio = comp_per_wrvu(tcc, wrvus)
        market_ratio = market_comp_per_wrvu(tcc_bench, wrvu_bench)
        if ratio and market_ratio:
            ratio_p = interpolate_percentile(ratio, market_ratio.get('25'),
                                             market_ratio.get('50'),
                                             market_ratio.get('75'), market_ratio.get('90'))
    return {
        'tcc': tcc, 'tcc_percentile': tcc_p,
        'wrvus': wrvus, 'wrvu_percentile': wrvu_p,
        'comp_per_wrvu': ratio, 'market_comp_per_wrvu': market_ratio,
        'comp_per_wrvu_percentile': ratio_p,
        'alignment': alignment_check(tcc_p, wrvu_p),
        'percentile_caveats': percentile_caveats(tcc_p),
        'survey_caveats': survey_use_caveats(n_sources, combined_n),
        'disclaimer': (
            "This is a market-approach analysis prepared to inform a compensation decision. "
            "It is not a certified fair market value opinion. A defensible FMV determination "
            "considers the market, income, and cost approaches together with the specific "
            "facts and circumstances of the arrangement, and where regulatory reliance is "
            "required should be performed by a qualified valuation professional."),
    }
