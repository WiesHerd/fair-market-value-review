"""
Pytest configuration and shared fixtures for the comp-adjustment-request skill.

Provides thin pytest wrappers around the mock Excel builders in
tests/fixtures/builders.py so tests are fully self-contained and don't depend
on real employer data.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = REPO_ROOT / "scripts"
FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"

# Re-export the builders under their original names so existing test code
# (and other fixtures) can call them directly.
from tests.fixtures.builders import (  # noqa: E402,F401
    build_mock_committee_template_xlsx,
    build_mock_salary_xlsx,
    build_mock_survey_xlsx,
)


# ─── Path fixtures ────────────────────────────────────────────────────

@pytest.fixture
def repo_root() -> Path:
    """Root of the comp-request-skill repo."""
    return REPO_ROOT


@pytest.fixture
def scripts_dir() -> Path:
    """Directory containing the executable scripts."""
    return SCRIPTS_DIR


@pytest.fixture
def fixtures_dir() -> Path:
    """Directory of mock data fixtures."""
    return FIXTURES_DIR


# ─── Pre-existing text/json fixtures ──────────────────────────────────

@pytest.fixture
def mock_cohort_tsv(fixtures_dir: Path) -> Path:
    """TSV cohort fixture (3 fictional providers)."""
    return fixtures_dir / "mock_cohort.tsv"


@pytest.fixture
def mock_anchors_json(fixtures_dir: Path) -> Path:
    """Market anchors JSON fixture (p25/p50/p75/p90)."""
    return fixtures_dir / "mock_anchors.json"


@pytest.fixture
def mock_benchmarks_txt(fixtures_dir: Path) -> Path:
    """Benchmark text paste fixture (Survey 1/Survey 2/Survey 3 format)."""
    return fixtures_dir / "mock_benchmarks.txt"


# ─── Per-test Excel fixtures (built fresh each time) ──────────────────

@pytest.fixture
def mock_salary_xlsx(tmp_path: Path) -> Path:
    """Freshly-built mock salary workbook (single Provider 1 row)."""
    return build_mock_salary_xlsx(tmp_path / "mock_salary.xlsx")


@pytest.fixture
def mock_survey_xlsx(tmp_path: Path) -> Path:
    """Freshly-built mock survey workbook (3 specialties incl. dash variants)."""
    return build_mock_survey_xlsx(tmp_path / "mock_survey.xlsx")


@pytest.fixture
def mock_template_xlsx(tmp_path: Path) -> Path:
    """Freshly-built mock committee template workbook."""
    return build_mock_committee_template_xlsx(tmp_path / "mock_template.xlsx")


# ─── Session-scoped checked-in Excel fixtures ─────────────────────────

@pytest.fixture(scope="session")
def checked_in_salary_xlsx(fixtures_dir: Path) -> Path:
    """Versioned mock salary workbook shipped in tests/fixtures/."""
    return fixtures_dir / "mock_salary.xlsx"


@pytest.fixture(scope="session")
def checked_in_survey_xlsx(fixtures_dir: Path) -> Path:
    """Versioned mock survey workbook shipped in tests/fixtures/."""
    return fixtures_dir / "mock_survey.xlsx"


@pytest.fixture(scope="session")
def checked_in_template_xlsx(fixtures_dir: Path) -> Path:
    """Versioned mock committee template shipped in tests/fixtures/."""
    return fixtures_dir / "mock_committee_template.xlsx"


# ─── JSON config fixture ──────────────────────────────────────────────

@pytest.fixture
def sample_request_config(
    tmp_path: Path,
    mock_salary_xlsx: Path,
    mock_survey_xlsx: Path,
    mock_template_xlsx: Path,
) -> Path:
    """A minimal request_config.json pointing at the per-test mock fixtures."""
    cfg = {
        "name": "Provider 1",
        "salary_file": str(mock_salary_xlsx),
        "survey_file": str(mock_survey_xlsx),
        "template": str(mock_template_xlsx),
        "proposed_base": 325000.0,
        "stipend": 42000,
        "current_stipend": 0,
        "wrvu": 4228,
        "track_num": 100026,
        "submit_date": None,
        "request_type": "Existing",
        "no_academic_rank": True,
        "academic_rank": None,
        "cart_clinical": 1.0,
        "cart_admin": 0,
        "cart_research": 0,
        "cart_teaching": 0,
        "request_summary": "Market adjustment to address compression.",
        "general_background": "Blended weighted-average benchmark.",
        "physician_background": "8 years post-fellowship.",
        "output": str(tmp_path / "output.xlsx"),
    }
    out = tmp_path / "request_config.json"
    with open(out, "w") as f:
        json.dump(cfg, f, indent=2)
    return out
