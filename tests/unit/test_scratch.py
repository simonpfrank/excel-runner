"""Unit tests for the scratch-copy execution model in excel_runner.engine (Spec sec 5.3,
PRD sec 6.3.1/6.3.3/6.3.4; docs/recalc_and_link_refresh_plan.md sec 1/3).

Scratch now has two subfolders instead of one flat dir: `scratch/working/` (copies named with
the workbook's *original real basename*, not the logical/YAML name — needed so same-folder
(R1) links between two staged workbooks resolve unchanged) and `scratch/originals/` (a
pre-edit backup, made only for write-intent workbooks). Commit is copy-based, not rename-based
(plan doc sec 3): the existing real file (if any) is copied — never moved — to a `.bak`
sibling first, then the working copy is copied onto the real path. Rollback likewise copies a
`.bak` back over `real_path` rather than renaming it.
"""

import logging
from pathlib import Path

import pytest

from excel_runner.core import ActionExecutionError
from excel_runner.engine import ScratchManager


def _write_file(path: Path, content: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content)
    return path


class TestStage:
    def test_copies_the_real_file_into_scratch_working(self, tmp_path: Path) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        working_path = manager.stage("manip", real)

        assert working_path.exists()
        assert working_path.read_text() == "original"
        assert working_path != real
        assert working_path.parent == tmp_path / "working" / "scratch" / "working"

    def test_working_copy_is_named_with_the_real_basename_not_the_logical_name(
        self, tmp_path: Path
    ) -> None:
        real = _write_file(tmp_path / "real" / "Backtesting Manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        working_path = manager.stage("backtesting_manip", real)

        assert working_path.name == "Backtesting Manip.xlsx"

    def test_handles_a_real_file_that_does_not_exist_yet(self, tmp_path: Path) -> None:
        """create_if_missing case — no real file to copy from yet, caller creates the
        workbook directly at the returned scratch path."""
        real = tmp_path / "real" / "new.xlsx"
        manager = ScratchManager(tmp_path / "working")

        working_path = manager.stage("new", real)

        assert not working_path.exists()
        assert working_path.parent == tmp_path / "working" / "scratch" / "working"

    def test_write_intent_workbook_gets_an_originals_backup(
        self, tmp_path: Path
    ) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        manager.stage("manip", real, writes=True)

        original_backup = tmp_path / "working" / "scratch" / "originals" / "manip.xlsx"
        assert original_backup.exists()
        assert original_backup.read_text() == "original"

    def test_read_only_workbook_gets_no_originals_backup(self, tmp_path: Path) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        manager.stage("manip", real, writes=False)

        original_backup = tmp_path / "working" / "scratch" / "originals" / "manip.xlsx"
        assert not original_backup.exists()


class TestCommit:
    def test_commit_copies_working_content_back_to_the_real_path(
        self, tmp_path: Path
    ) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        working_path = manager.stage("manip", real)
        working_path.write_text("changed")

        manager.commit("manip")

        assert real.read_text() == "changed"

    def test_commit_creates_the_real_file_when_it_did_not_exist_before(
        self, tmp_path: Path
    ) -> None:
        real = tmp_path / "real" / "new.xlsx"
        manager = ScratchManager(tmp_path / "working")
        working_path = manager.stage("new", real)
        working_path.write_text("brand new content")

        manager.commit("new")

        assert real.read_text() == "brand new content"

    def test_commit_backs_up_the_original_via_copy_not_rename(
        self, tmp_path: Path
    ) -> None:
        """The original is preserved as a <real>.bak sibling during commit — copied, never
        moved/deleted, before the overwrite (plan doc sec 3.2)."""
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("manip", real).write_text("changed")

        manager.commit("manip")

        bak = real.with_name("manip.xlsx.bak")
        assert bak.exists()
        assert bak.read_text() == "original"
        assert real.read_text() == "changed"


class TestCommitAll:
    def test_commits_every_staged_workbook(self, tmp_path: Path) -> None:
        real_a = _write_file(tmp_path / "real" / "a.xlsx", "a-original")
        real_b = _write_file(tmp_path / "real" / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        manager.commit_all()

        assert real_a.read_text() == "a-changed"
        assert real_b.read_text() == "b-changed"

    def test_deletes_every_backup_on_full_success(self, tmp_path: Path) -> None:
        real_a = _write_file(tmp_path / "real" / "a.xlsx", "a-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")

        manager.commit_all()

        assert not real_a.with_name("a.xlsx.bak").exists()

    def test_a_later_failure_rolls_back_earlier_successful_commits(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        real_a = _write_file(tmp_path / "real" / "a.xlsx", "a-original")
        real_b = _write_file(tmp_path / "real" / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        real_commit = manager.commit
        calls = {"count": 0}

        def failing_commit(name: str) -> None:
            calls["count"] += 1
            if calls["count"] == 2:
                raise OSError("simulated: file locked elsewhere")
            real_commit(name)

        monkeypatch.setattr(manager, "commit", failing_commit)

        with pytest.raises(ActionExecutionError):
            manager.commit_all()

        assert real_a.read_text() == "a-original"  # rolled back
        assert real_b.read_text() == "b-original"  # never touched

    def test_rollback_of_a_newly_created_workbook_removes_it_rather_than_restoring(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """A workbook staged with no real_path yet (create_if_missing) has no `.bak` to
        restore from on rollback — it should simply be removed, not restored to nothing.
        """
        new_real = tmp_path / "real" / "new.xlsx"
        real_b = _write_file(tmp_path / "real" / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("new", new_real).write_text("brand new content")
        manager.stage("b", real_b).write_text("b-changed")

        real_commit = manager.commit

        def failing_commit(name: str) -> None:
            if name == "b":
                raise OSError("simulated: file locked elsewhere")
            real_commit(name)

        monkeypatch.setattr(manager, "commit", failing_commit)

        with pytest.raises(ActionExecutionError):
            manager.commit_all()

        assert not new_real.exists()  # rolled back to "never existed"
        assert real_b.read_text() == "b-original"  # never touched

    def test_needs_human_is_reported_when_rollback_itself_fails(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        real_a = _write_file(tmp_path / "real" / "a.xlsx", "a-original")
        real_b = _write_file(tmp_path / "real" / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        real_commit = manager.commit

        def failing_commit(name: str) -> None:
            if name == "b":
                raise OSError("simulated: file locked elsewhere")
            real_commit(name)

        monkeypatch.setattr(manager, "commit", failing_commit)

        import shutil as shutil_module

        original_copy2 = shutil_module.copy2

        def failing_copy2(
            src: object, dst: object, *args: object, **kwargs: object
        ) -> object:
            # Only the rollback step ever copies a .bak sibling back over real_path — this
            # leaves the normal commit path (which copies real_path -> .bak, the other
            # direction) working normally, so "a" genuinely commits first before its rollback
            # is attempted and fails.
            if isinstance(src, Path) and src.name.endswith(".bak"):
                raise OSError("rollback blocked too")
            return original_copy2(src, dst, *args, **kwargs)

        monkeypatch.setattr(shutil_module, "copy2", failing_copy2)

        with pytest.raises(ActionExecutionError) as exc_info:
            manager.commit_all()

        assert "b" in exc_info.value.detail.message  # the workbook whose commit failed
        assert "a" in exc_info.value.detail.message  # needs_human, named explicitly


class TestLogging:
    """Production-visibility logging (AGENTS.md's logging section) — staging/committing a
    workbook is exactly the kind of "meaningful step" that should be visible at INFO without
    reading code, especially since a large workbook's commit can take a while."""

    def test_stage_logs_at_info(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        with caplog.at_level(logging.INFO, logger="excel_runner.engine"):
            manager.stage("manip", real)

        messages = " ".join(r.message for r in caplog.records)
        assert "manip" in messages
        assert "manip.xlsx" in messages

    def test_commit_logs_at_info(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("manip", real).write_text("changed")

        with caplog.at_level(logging.INFO, logger="excel_runner.engine"):
            manager.commit("manip")

        messages = " ".join(r.message for r in caplog.records)
        assert "manip" in messages

    def test_commit_all_logs_start_and_completion(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("manip", real).write_text("changed")

        with caplog.at_level(logging.INFO, logger="excel_runner.engine"):
            manager.commit_all()

        messages = [r.message for r in caplog.records]
        assert any("1" in m for m in messages)  # count of workbooks committed
        assert any("complete" in m.lower() or "committed" in m.lower() for m in messages)
