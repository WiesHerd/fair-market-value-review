#!/usr/bin/env python3
"""Build the session-scoped checked-in Excel fixtures for tests/fixtures/.

This is a plain script (no pytest decorators) that uses the same builders
the tests use, so the fixtures committed to git always reflect the latest
schema expected by the test suite.

Run from the repo root:

    python3 tests/fixtures/build_checked_in_fixtures.py
"""
from __future__ import annotations

import sys
from pathlib import Path

# Make the repo root importable so `tests.fixtures.builders` resolves.
REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tests.fixtures.builders import (  # type: ignore[import-not-found]  # noqa: E402
    build_mock_committee_template_xlsx,
    build_mock_salary_xlsx,
    build_mock_survey_xlsx,
)

FX_DIR = Path(__file__).resolve().parents[1] / "tests" / "fixtures"


def main() -> None:
    build_mock_salary_xlsx(FX_DIR / "mock_salary.xlsx")
    build_mock_survey_xlsx(FX_DIR / "mock_survey.xlsx")
    build_mock_committee_template_xlsx(FX_DIR / "mock_fmv_template.xlsx")
    for p in sorted(FX_DIR.glob("mock_*.xlsx")):
        print(f"{p.name}: {p.stat().st_size} bytes")


if __name__ == "__main__":
    main()
