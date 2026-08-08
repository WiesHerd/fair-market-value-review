"""Tests for the FMV market-approach analytics module."""
import sys
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from fmv_analysis import (interpolate_percentile, comp_per_wrvu, market_comp_per_wrvu,
                          alignment_check, percentile_caveats, survey_use_caveats,
                          build_fmv_summary)

TCC_B = {'25': 281038, '50': 321038, '75': 366038, '90': 412717}
WRVU_B = {'25': 3604, '50': 4204, '75': 4904, '90': 5604}


class TestInterpolation:
    def test_exact_points(self):
        assert interpolate_percentile(281038, 281038, 321038, 366038, 412717) == 25.0
        assert interpolate_percentile(321038, 281038, 321038, 366038, 412717) == 50.0
        assert interpolate_percentile(412717, 281038, 321038, 366038, 412717) == 90.0

    def test_midpoint_between_p50_and_p75(self):
        mid = (321038 + 366038) / 2
        assert abs(interpolate_percentile(mid, 281038, 321038, 366038, 412717) - 62.5) < 0.01

    def test_missing_inputs_return_none(self):
        assert interpolate_percentile(300000, None, 321038, 366038, 412717) is None
        assert interpolate_percentile(None, 281038, 321038, 366038, 412717) is None


class TestCompPerWrvu:
    def test_basic_ratio(self):
        assert comp_per_wrvu(400000, 5000) == 80.0

    def test_zero_wrvus_returns_none(self):
        assert comp_per_wrvu(400000, 0) is None

    def test_market_ratio_is_relatively_flat_across_quartiles(self):
        """The finding that makes applying high-quartile ratios wrong."""
        m = market_comp_per_wrvu(TCC_B, WRVU_B)
        assert abs(m['50'] - (321038 / 4204)) < 0.01
        assert max(m.values()) / min(m.values()) < 1.35


class TestAlignment:
    def test_aligned_when_close(self):
        assert alignment_check(52.0, 50.0)['status'] == 'aligned'

    def test_flags_comp_above_production(self):
        r = alignment_check(75.0, 51.0)
        assert r['status'] == 'comp_above_production'
        assert r['gap'] == pytest.approx(24.0)
        assert "document the basis" in r['message'].lower()

    def test_flags_production_above_comp(self):
        r = alignment_check(40.0, 80.0)
        assert r['status'] == 'production_above_comp'
        assert "retention risk" in r['message'].lower()

    def test_insufficient_data(self):
        assert alignment_check(None, 50.0)['status'] == 'insufficient_data'


class TestCaveats:
    def test_no_percentile_is_automatically_fmv(self):
        assert "no percentile is inherently fair market value" in " ".join(percentile_caveats(50.0)).lower()

    def test_p75_adds_documentation_note(self):
        assert "not disqualifying" in " ".join(percentile_caveats(78.0)).lower()

    def test_p90_adds_stacking_note(self):
        assert "stacked" in " ".join(percentile_caveats(93.0)).lower()

    def test_single_source_flagged(self):
        assert "only one survey source" in " ".join(survey_use_caveats(1)).lower()

    def test_multi_source_notes_overlap_risk(self):
        assert "double-count" in " ".join(survey_use_caveats(3)).lower()

    def test_small_sample_flagged(self):
        assert "small" in " ".join(survey_use_caveats(2, combined_n=12)).lower()

    def test_app_billing_caveat_always_present(self):
        assert "advanced practice" in " ".join(survey_use_caveats(2, combined_n=200)).lower()


class TestSummary:
    def test_full_summary_flags_misalignment(self):
        r = build_fmv_summary(367000, 4228, TCC_B, WRVU_B, n_sources=3, combined_n=265)
        assert r['alignment']['status'] == 'comp_above_production'
        assert r['comp_per_wrvu'] == pytest.approx(367000 / 4228)
        assert r['tcc_percentile'] > r['wrvu_percentile']

    def test_summary_includes_non_opinion_disclaimer(self):
        assert "not a certified fair market value opinion" in build_fmv_summary(300000, 4000, TCC_B, WRVU_B)['disclaimer']

    def test_summary_without_wrvus_still_positions_tcc(self):
        r = build_fmv_summary(321038, None, TCC_B)
        assert abs(r['tcc_percentile'] - 50.0) < 0.1
        assert r['alignment']['status'] == 'insufficient_data'
