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

from excel_runner import engine
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


class TestNeededBackend:
    """PRD sec 6.2.2's capability -> backend mapping, as a standalone pure function."""

    def test_file_capability_needs_file_backend(self) -> None:
        assert engine._needed_backend("file") == "file"

    def test_xlw_capability_needs_xlw_backend(self) -> None:
        assert engine._needed_backend("xlw") == "xlw"

    def test_com_capability_needs_xlw_backend_too(self) -> None:
        """com reaches deeper via xlwings' .api on an xlw-backed session — SessionManager
        never needs a distinct backend state for it (PRD sec 6.2.2)."""
        assert engine._needed_backend("com") == "xlw"

    def test_depends_on_param_capability_is_not_resolvable_here(self) -> None:
        """read_metadata's real capability depends on its target: param at runtime — that
        resolution isn't built yet (Spec sec 5.1), so this can't be mapped to a backend yet."""
        with pytest.raises(ActionExecutionError):
            engine._needed_backend("depends_on_param")

    def test_none_capability_is_not_resolvable_here(self) -> None:
        """Control actions (stop) never reach get_or_open at all — no workbook: field — so this
        is defensive, not a real path (PRD sec 6.9)."""
        with pytest.raises(ActionExecutionError):
            engine._needed_backend("none")


class TestCapabilityBackendMismatch:
    """PRD sec 6.2.2: bidirectional backend switching isn't built yet, so a capability that
    doesn't match a session's current backend must raise clearly rather than silently return
    the wrong backend or crash unhelpfully."""

    def test_matching_capability_returns_the_session_normally(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("manip", capability="file")

        assert session.backend == "file"

    def test_default_capability_is_file_unchanged_from_before_this_param_existed(
        self, tmp_path: Path
    ) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        session = manager.get_or_open("manip")  # no capability given

        assert session.backend == "file"

    def test_mismatched_capability_on_a_brand_new_session_raises_clearly(self, tmp_path: Path) -> None:
        """A brand-new session always opens file-backend today (Spec sec 5.2) — an xlw/com
        capability request against it can't be served without switching, which isn't built
        yet (PRD sec 6.2.2)."""
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))

        with pytest.raises(ActionExecutionError) as exc_info:
            manager.get_or_open("manip", capability="xlw")

        assert "manip" in exc_info.value.detail.message
        assert "switch" in exc_info.value.detail.message.lower()

    def test_mismatched_capability_on_an_already_open_session_raises_clearly(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        manager.get_or_open("manip", capability="file")  # opens it as file-backend first

        with pytest.raises(ActionExecutionError):
            manager.get_or_open("manip", capability="xlw")


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


class TestCheckpoint:
    """checkpoint() persists in-progress writes to the scratch file mid-run, so a later crash
    leaves everything that succeeded so far visible in the recovery artifact (PRD sec 6.3.1) —
    found necessary via a failing integration test: without this, the scratch file on disk
    only ever reflected whatever was there at staging time, since openpyxl writes stay in
    memory until an explicit save and nothing else triggers one mid-run."""

    def test_checkpoint_saves_a_dirty_staged_session_to_its_scratch_file(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        session = manager.get_or_open("manip", mode="read_write")
        session.handle["Sheet"]["A1"] = "in progress"
        session.dirty = True

        manager.checkpoint()

        assert openpyxl.load_workbook(session.path)["Sheet"]["A1"].value == "in progress"

    def test_checkpoint_does_not_touch_the_real_path(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        session = manager.get_or_open("manip", mode="read_write")
        session.handle["Sheet"]["A1"] = "in progress"
        session.dirty = True

        manager.checkpoint()

        assert openpyxl.load_workbook(real)["Sheet"]["A1"].value == "original"

    def test_checkpoint_clears_the_dirty_flag(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        session = manager.get_or_open("manip", mode="read_write")
        session.dirty = True

        manager.checkpoint()

        assert session.dirty is False

    def test_checkpoint_skips_a_non_dirty_session(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        manager.get_or_open("manip", mode="read_write")

        manager.checkpoint()  # should not raise

    def test_checkpoint_skips_read_only_sessions(self, tmp_path: Path) -> None:
        real = _write_workbook(tmp_path / "real" / "manip.xlsx")
        workbooks = {"manip": WorkbookRef(name="manip", file=str(real))}
        manager = SessionManager(workbooks, ScratchManager(tmp_path / "scratch"))
        manager.get_or_open("manip", mode="read_only")

        manager.checkpoint()  # should not raise


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
