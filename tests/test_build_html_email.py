"""Tests for build_html_email.py -- the email-ready HTML report generator."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from build_html_email import build_html, money


@pytest.fixture
def providers():
    return [
        {"name": "Provider 1", "yoe": 8, "current_base": 275000, "planned_base": 287000, "proposed_base": 325000},
        {"name": "Provider 3", "yoe": 5, "current_base": 260000, "planned_base": 270000, "proposed_base": 298000},
    ]


@pytest.fixture
def anchors():
    return {"p25": 280000, "p50": 320000, "p75": 365000, "p90": 410000}


class TestMoney:
    def test_formats_thousands(self):
        assert money(316200) == "$316,200"

    def test_rounds_to_whole_dollars(self):
        assert money(15177.6) == "$15,178"


class TestBuildHtml:
    def test_produces_self_contained_html(self, providers, anchors):
        html = build_html(providers, anchors, None, {"cohort_name": "Test Cohort"})
        assert html.startswith("<!DOCTYPE html>")
        assert "</html>" in html
        # No external resource references -- must be paste-into-email safe
        assert 'src="http' not in html
        assert '<link' not in html

    def test_includes_all_provider_rows(self, providers, anchors):
        html = build_html(providers, anchors, None, {"cohort_name": "Test"})
        for p in providers:
            assert p["name"] in html

    def test_totals_are_correct(self, providers, anchors):
        html = build_html(providers, anchors, None, {"cohort_name": "Test"})
        assert money(275000 + 260000) in html      # total current
        assert money(325000 + 298000) in html      # total proposed
        assert money((325000 + 298000) - (275000 + 260000)) in html  # delta vs current

    def test_escapes_html_in_names(self, anchors):
        providers = [{"name": "O'Brien <script>alert(1)</script>", "yoe": 1,
                      "current_base": 100000, "planned_base": 101000, "proposed_base": 105000}]
        html = build_html(providers, anchors, None, {"cohort_name": "Test"})
        assert "<script>alert(1)</script>" not in html
        assert "&lt;script&gt;" in html

    def test_embeds_exhibit_as_data_uri(self, providers, anchors, tmp_path):
        png = tmp_path / "exhibit.png"
        png.write_bytes(b"\x89PNG\r\n\x1a\n" + b"0" * 50)
        html = build_html(providers, anchors, str(png), {"cohort_name": "Test"})
        assert "data:image/png;base64," in html

    def test_page_margin_is_0_6in(self, providers, anchors):
        html = build_html(providers, anchors, None, {"cohort_name": "Test"})
        assert "margin: 0.6in" in html

    def test_custom_rationale_used_when_provided(self, providers, anchors):
        meta = {"cohort_name": "Test", "rationale": ["My custom rationale bullet."]}
        html = build_html(providers, anchors, None, meta)
        assert "My custom rationale bullet." in html
        assert "Compression guardrails" not in html  # default not used
