"""Unit tests for the scratch-copy execution model in excel_runner.engine (Spec sec 5.3,
PRD sec 6.3.1/6.3.3/6.3.4). ScratchManager now takes a working_dir (not a scratch_dir directly
— scratch/ is a subfolder of it), and commit_all() uses a rename-based commit with per-file
rollback on a later failure, instead of the old copy-then-Path.replace with no rollback."""

from pathlib import Path

import pytest

from excel_runner.core import ActionExecutionError
from excel_runner.engine import ScratchManager


def _write_file(path: Path, content: str) -> Path:
    path.write_text(content)
    return path


class TestStage:
    def test_copies_the_real_file_into_the_scratch_dir(self, tmp_path: Path) -> None:
        (tmp_path / "real").mkdir()
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")

        scratch_path = manager.stage("manip", real)

        assert scratch_path.exists()
        assert scratch_path.read_text() == "original"
        assert scratch_path != real
        assert scratch_path.parent == tmp_path / "working" / "scratch"

    def test_handles_a_real_file_that_does_not_exist_yet(self, tmp_path: Path) -> None:
        """create_if_missing case — no real file to copy from yet, caller creates the
        workbook directly at the returned scratch path."""
        real = tmp_path / "real" / "new.xlsx"
        manager = ScratchManager(tmp_path / "working")

        scratch_path = manager.stage("new", real)

        assert not scratch_path.exists()
        assert scratch_path.parent == tmp_path / "working" / "scratch"


class TestCommit:
    def test_commit_moves_scratch_content_back_to_the_real_path(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real = _write_file(real_dir / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        scratch_path = manager.stage("manip", real)
        scratch_path.write_text("changed")

        manager.commit("manip")

        assert real.read_text() == "changed"

    def test_commit_creates_the_real_file_when_it_did_not_exist_before(self, tmp_path: Path) -> None:
        real = tmp_path / "real" / "new.xlsx"
        manager = ScratchManager(tmp_path / "working")
        scratch_path = manager.stage("new", real)
        scratch_path.write_text("brand new content")

        manager.commit("new")

        assert real.read_text() == "brand new content"

    def test_commit_backs_up_the_original_via_rename_not_copy(self, tmp_path: Path) -> None:
        """The original is preserved as a <real>.bak sibling during commit — a zero-copy
        rename, since the original was already sitting there untouched (PRD sec 6.3.3)."""
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real = _write_file(real_dir / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("manip", real).write_text("changed")

        manager.commit("manip")

        bak = real_dir / "manip.xlsx.bak"
        assert bak.exists()
        assert bak.read_text() == "original"


class TestCommitAll:
    def test_commits_every_staged_workbook(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real_a = _write_file(real_dir / "a.xlsx", "a-original")
        real_b = _write_file(real_dir / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        manager.commit_all()

        assert real_a.read_text() == "a-changed"
        assert real_b.read_text() == "b-changed"

    def test_deletes_every_backup_on_full_success(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real_a = _write_file(real_dir / "a.xlsx", "a-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")

        manager.commit_all()

        assert not (real_dir / "a.xlsx.bak").exists()

    def test_a_later_failure_rolls_back_earlier_successful_commits(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real_a = _write_file(real_dir / "a.xlsx", "a-original")
        real_b = _write_file(real_dir / "b.xlsx", "b-original")
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
        restore from on rollback — it should simply be removed, not restored to nothing."""
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        new_real = real_dir / "new.xlsx"
        real_b = _write_file(real_dir / "b.xlsx", "b-original")
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
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real_a = _write_file(real_dir / "a.xlsx", "a-original")
        real_b = _write_file(real_dir / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "working")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        real_commit = manager.commit

        def failing_commit(name: str) -> None:
            if name == "b":
                raise OSError("simulated: file locked elsewhere")
            real_commit(name)

        monkeypatch.setattr(manager, "commit", failing_commit)

        original_rename = Path.rename

        def failing_rename(self: Path, target: Path) -> Path:
            # Only the rollback step ever renames a .bak sibling back into place — this leaves
            # the normal commit path (which never touches a .bak) working normally, so "a"
            # genuinely commits first before its rollback is attempted and fails.
            if self.name.endswith(".bak"):
                raise OSError("rollback blocked too")
            return original_rename(self, target)

        monkeypatch.setattr(Path, "rename", failing_rename)

        with pytest.raises(ActionExecutionError) as exc_info:
            manager.commit_all()

        assert "b" in exc_info.value.detail.message  # the workbook whose commit failed
        assert "a" in exc_info.value.detail.message  # needs_human, named explicitly
