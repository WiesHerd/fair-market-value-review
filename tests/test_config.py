"""
Validate templates/request_config.json.

The committed request config drives committee_template_generator.py at runtime,
so its shape must remain stable. Required fields, sensible types, and the
non-trivial invariants that downstream code depends on.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = REPO_ROOT / "templates" / "request_config.json"


@pytest.fixture(scope="module")
def config() -> dict:
    with open(CONFIG_PATH) as f:
        return json.load(f)


class TestRequestConfig:
    def test_file_exists(self) -> None:
        assert CONFIG_PATH.exists(), f"Missing config file: {CONFIG_PATH}"

    def test_is_valid_json(self) -> None:
        with open(CONFIG_PATH) as f:
            json.load(f)  # raises if invalid

    def test_is_dict(self, config) -> None:
        assert isinstance(config, dict)

    def test_required_string_fields(self, config) -> None:
        for field in ("name", "salary_file", "survey_file",
                      "template", "output"):
            assert field in config, f"Missing required field: {field}"
            assert isinstance(config[field], str) and config[field].strip(), (
                f"{field} must be a non-empty string"
            )

    def test_required_numeric_fields(self, config) -> None:
        assert "proposed_base" in config
        assert isinstance(config["proposed_base"], (int, float))
        assert config["proposed_base"] > 0
        assert "wrvu" in config
        assert isinstance(config["wrvu"], (int, float))
        assert config["wrvu"] >= 0

    def test_request_type_is_known(self, config) -> None:
        """The script's logic only handles a small set of request types."""
        allowed = {"Existing", "Incr New", "New"}
        assert config.get("request_type") in allowed, (
            f"request_type={config.get('request_type')!r} not in {allowed}"
        )

    def test_cart_percentages_sum_to_one(self, config) -> None:
        """CART (Clinical/Admin/Research/Teaching) effort must sum to 1.0."""
        cart_keys = ("cart_clinical", "cart_admin",
                     "cart_research", "cart_teaching")
        for k in cart_keys:
            assert k in config, f"Missing CART key: {k}"
            assert isinstance(config[k], (int, float))
        total = sum(config[k] for k in cart_keys)
        assert abs(total - 1.0) < 1e-6, (
            f"CART percentages sum to {total}, expected 1.0"
        )

    def test_no_academic_rank_toggle_and_rank_value_consistent(self, config) -> None:
        """If no_academic_rank is true, rank must be null. Otherwise, set."""
        no_rank = config.get("no_academic_rank")
        rank = config.get("academic_rank")
        if no_rank is True:
            assert rank is None, (
                "no_academic_rank=True but academic_rank is populated"
            )

    def test_no_real_organization_names(self) -> None:
        """No reference to any real organization anywhere in the config."""
        text = CONFIG_PATH.read_text()
        forbidden = [
            "Acme Health",  # generic placeholder — no real org names
            "AMH",
            "acme medical",
        ]
        for token in forbidden:
            assert token.lower() not in text.lower(), (
                f"Config references real org/token: {token!r}"
            )

    def test_no_synthetic_fixtures_in_committed_config(self) -> None:
        """Committed config should not point at synthetic test fixtures."""
        text = CONFIG_PATH.read_text()
        for fake in ("tests/fixtures/", "/tmp/", "mock_"):
            assert fake not in text, (
                f"Committed config still references test fixture path: {fake}"
            )

    def test_salary_survey_template_files_end_in_xlsx(self, config) -> None:
        for field in ("salary_file", "survey_file", "template", "output"):
            assert config[field].endswith(".xlsx"), (
                f"{field}={config[field]!r} should be .xlsx"
            )

    def test_request_summary_is_present(self, config) -> None:
        """A non-empty request_summary is mandatory for a meaningful form."""
        assert config.get("request_summary")
        assert isinstance(config["request_summary"], str)
        assert len(config["request_summary"].strip()) > 10

    def test_general_and_physician_background_present(self, config) -> None:
        for field in ("general_background", "physician_background"):
            assert config.get(field), f"{field} missing"
            assert isinstance(config[field], str)
            assert len(config[field].strip()) > 10

    def test_no_email_addresses_or_tokens_leaked(self) -> None:
        """A repo-boundary safety check."""
        text = CONFIG_PATH.read_text()
        # No "@" email patterns, no obvious token-looking strings.
        assert not re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", text), (
            "Config contains an email-looking string"
        )
        assert "ghp_" not in text and "github_pat_" not in text, (
            "Config appears to contain a GitHub token"
        )
