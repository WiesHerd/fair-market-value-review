"""Tests for the FMV request register."""
import csv, sys
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
import fmv_register as reg


@pytest.fixture(autouse=True)
def isolated_register(tmp_path, monkeypatch):
    """Point the register at a temp dir so tests never touch a real register."""
    monkeypatch.setattr(reg, "REGISTER_DIR", tmp_path / "register")
    monkeypatch.setattr(reg, "REGISTER_FILE", tmp_path / "register" / "fmv_requests.csv")
    yield


class _Args(dict):
    def __getattr__(self, k):
        return self.get(k)


def _add(**kw):
    base = {"provider": "Provider 1", "specialty": "PICU", "type": "adjustment"}
    base.update(kw)
    return reg.cmd_add(_Args(base))


class TestAdd:
    def test_creates_register_and_assigns_id(self):
        assert _add() == "FMV-0001"
        assert reg.REGISTER_FILE.exists()

    def test_ids_increment(self):
        _add()
        assert _add(provider="Provider 2") == "FMV-0002"

    def test_new_request_starts_in_intake(self):
        _add()
        assert reg._load()[0]["status"] == "intake"

    def test_records_core_fields(self):
        _add(requester="Division Chief", due="2026-09-01")
        r = reg._load()[0]
        assert r["provider"] == "Provider 1"
        assert r["review_type"] == "adjustment"
        assert r["requester"] == "Division Chief"
        assert r["due_date"] == "2026-09-01"


class TestUpdate:
    def test_updates_analysis_fields(self):
        _add()
        reg.cmd_update(_Args({"request_id": "FMV-0001", "status": "in_review",
                              "proposed_base": "325000", "tcc_percentile": "75.3",
                              "alignment": "comp_above_production"}))
        r = reg._load()[0]
        assert r["status"] == "in_review"
        assert r["alignment_flag"] == "comp_above_production"

    def test_closing_status_stamps_decision_date(self):
        _add()
        reg.cmd_update(_Args({"request_id": "FMV-0001", "status": "approved", "decision": "OK"}))
        assert reg._load()[0]["decision_date"] != ""

    def test_unknown_id_exits(self):
        _add()
        with pytest.raises(SystemExit):
            reg.cmd_update(_Args({"request_id": "FMV-9999", "status": "approved"}))

    def test_case_insensitive_id(self):
        _add()
        reg.cmd_update(_Args({"request_id": "fmv-0001", "status": "in_review"}))
        assert reg._load()[0]["status"] == "in_review"


class TestListAndSummary:
    def test_open_filter_excludes_closed(self, capsys):
        _add(); _add(provider="Provider 2")
        reg.cmd_update(_Args({"request_id": "FMV-0002", "status": "approved"}))
        capsys.readouterr()  # discard setup output
        reg.cmd_list(_Args({"open": True}))
        out = capsys.readouterr().out
        assert "FMV-0001" in out and "FMV-0002" not in out

    def test_summary_reports_flags(self, capsys):
        _add()
        reg.cmd_update(_Args({"request_id": "FMV-0001", "alignment": "comp_above_production"}))
        reg.cmd_summary(_Args({}))
        assert "Alignment flags to document" in capsys.readouterr().out

    def test_summary_reports_overdue(self, capsys):
        _add(due="2000-01-01")
        reg.cmd_summary(_Args({}))
        assert "OVERDUE" in capsys.readouterr().out

    def test_empty_register_is_graceful(self, capsys):
        reg.cmd_list(_Args({}))
        assert "No requests logged yet" in capsys.readouterr().out


class TestCsvShape:
    def test_header_matches_fields(self):
        _add()
        with reg.REGISTER_FILE.open(newline="", encoding="utf-8") as f:
            assert next(csv.reader(f)) == reg.FIELDS
