"""Unit tests for SessionManager (Spec sec 5.2).

Read/write mode is caller-specified here, not statically inferred — tier-2 validation
(Spec sec 5.4, not built yet) will compute that later and hand it in; SessionManager itself
doesn't guess. A session opened with mode="read_write" is treated as "this run will write to
it" and gets staged through ScratchManager (PRD sec 6.3.1); mode="read_only" opens directly
against the real path, no staging.
"""

from pathlib import Path

import openpyxl
import pytest

from excel_runner.core import ActionExecutionError, WorkbookRef, WorkbookSession
from excel_runner.engine import ScratchManager, SessionManager


def _write_workbook(path: Path, cell_value: str = "original") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = cell_value
    workbook.save(path)
    return path


class TestGetOrOpen:
    def test_opens_an_existing_workbook_read_write_by_default(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("manip")

        assert session.name == "manip"
        assert session.mode == "read_write"
        assert session.handle["Sheet"]["A1"].value == "original"

    def test_read_write_stages_through_scratch(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("manip", mode="read_write")

        assert session.scratch_path is not None
        assert str(tmp_path / "scratch") in session.path
        assert session.path != str(real)

    def test_read_only_does_not_stage(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("manip", mode="read_only")

        assert session.scratch_path is None
        assert session.path == str(real)

    def test_second_call_for_the_same_name_returns_the_cached_session(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        first = manager.get_or_open("manip")
        second = manager.get_or_open("manip")

        assert first is second

    def test_unknown_workbook_name_raises_a_clear_error(self, tmp_path: Path) -> None:
        manager = SessionManager({}, ScratchManager(tmp_path / "scratch"))
        with pytest.raises(ActionExecutionError) as exc_info:
            manager.get_or_open("nonexistent")
        assert "nonexistent" in exc_info.value.detail.message

    def test_missing_file_without_create_if_missing_raises_a_clear_error(self, tmp_path: Path) -> None:
        workbooks = {"manip": WorkbookRef(name="manip", file=str(tmp_path / "missing.xlsx"))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        with pytest.raises(ActionExecutionError) as exc_info:
            manager.get_or_open("manip")
        assert "missing.xlsx" in exc_info.value.detail.message


class TestCreateIfMissing:
    def test_read_only_with_create_if_missing_creates_at_the_real_path(self, tmp_path: Path) -> None:
        """Unusual combination (why read a workbook you just created blank?) but not
        forbidden — must still work correctly rather than being an untested code path."""
        real = tmp_path / "real" / "new.xlsx"
        workbooks = {"new": WorkbookRef(name="new", file=str(real), create_if_missing=True)}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("new", mode="read_only")

        assert real.exists()
        assert session.scratch_path is None
        assert session.handle.sheetnames

    def test_creates_a_blank_workbook_at_the_scratch_path(self, tmp_path: Path) -> None:
        real = tmp_path / "real" / "new.xlsx"
        workbooks = {"new": WorkbookRef(name="new", file=str(real), create_if_missing=True)}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("new")

        assert not real.exists()  # not created in place — only in scratch, until commit
        assert session.handle.sheetnames

    def test_creates_from_a_template_workbook(self, tmp_path: Path) -> None:
        template_real = _write_workbook(tmp_path / "real" / "historical.xlsx", "template content")
        new_real = tmp_path / "real" / "results.xlsx"
        workbooks = {
            "historical": WorkbookRef(name="historical", file=str(template_real)),
            "results": WorkbookRef(
                name="results", file=str(new_real), create_if_missing=True, template="historical"
            ),
        }
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("results")

        assert session.handle["Sheet"]["A1"].value == "template content"


class TestCloseAll:
    def test_closes_every_opened_session(self, tmp_path: Path) -> None:
        real_a = _write_workbook(tmp_path / "real" / "a.xlsx")
        real_b = _write_workbook(tmp_path / "real" / "b.xlsx")
        workbooks = {
            "a": WorkbookRef(name="a", file=str(real_a)),
            "b": WorkbookRef(name="b", file=str(real_b)),
        }
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        manager.get_or_open("a")
        manager.get_or_open("b")

        manager.close_all()  # should not raise

    def test_close_all_with_no_sessions_opened_does_not_raise(self, tmp_path: Path) -> None:
        manager = SessionManager({}, ScratchManager(tmp_path / "scratch"))
        manager.close_all()

    def test_one_failing_close_does_not_prevent_others_from_closing(self, tmp_path: Path) -> None:
        """Crash-safety requirement (PRD sec 6.3): every session must get a close attempt,
        even if an earlier one fails. Uses a fake handle whose close() raises — openpyxl's
        own Workbook.close() is a no-op even when called twice, so it can't produce a real
        failure to test against."""

        class _ExplodingHandle:
            def close(self) -> None:
                raise RuntimeError("simulated close failure")

        real_b = _write_workbook(tmp_path / "real" / "b.xlsx")
        workbooks = {"b": WorkbookRef(name="b", file=str(real_b))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        # "a" (which will fail to close) is inserted first, so iteration reaches it before "b" —
        # this is what actually proves close_all() doesn't stop after the first failure.
        manager._sessions["a"] = WorkbookSession(
            name="a", backend="file", handle=_ExplodingHandle(), path="a.xlsx", mode="read_write"
        )
        session_b = manager.get_or_open("b")
        closed_b = False
        real_close_b = session_b.handle.close

        def _tracking_close() -> None:
            nonlocal closed_b
            closed_b = True
            real_close_b()

        session_b.handle.close = _tracking_close

        with pytest.raises(ExceptionGroup):
            manager.close_all()

        assert closed_b is True


class TestCommitAll:
    def test_saves_dirty_staged_sessions_and_commits_them_to_the_real_path(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        session = manager.get_or_open("manip", mode="read_write")
        session.handle["Sheet"]["A1"] = "changed"
        session.dirty = True

        manager.commit_all()

        reopened = openpyxl.load_workbook(real)
        assert reopened["Sheet"]["A1"].value == "changed"

    def test_read_only_sessions_are_not_touched_by_commit(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        manager.get_or_open("manip", mode="read_only")

        manager.commit_all()  # should not raise, nothing to commit
