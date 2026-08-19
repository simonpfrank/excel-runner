"""Unit tests for the scratch-copy execution model in excel_runner.engine (Spec sec 5.3,
PRD sec 6.3.1)."""

from pathlib import Path

from excel_runner.engine import ScratchManager


def _write_file(path: Path, content: str) -> Path:
    path.write_text(content)
    return path


class TestStage:
    def test_copies_the_real_file_into_the_scratch_dir(self, tmp_path: Path) -> None:
        (tmp_path / "real").mkdir()
        real = _write_file(tmp_path / "real" / "manip.xlsx", "original")
        scratch_dir = tmp_path / "scratch"
        manager = ScratchManager(scratch_dir)

        scratch_path = manager.stage("manip", real)

        assert scratch_path.exists()
        assert scratch_path.read_text() == "original"
        assert scratch_path != real

    def test_handles_a_real_file_that_does_not_exist_yet(self, tmp_path: Path) -> None:
        """create_if_missing case — no real file to copy from yet, caller creates the
        workbook directly at the returned scratch path."""
        real = tmp_path / "real" / "new.xlsx"
        scratch_dir = tmp_path / "scratch"
        manager = ScratchManager(scratch_dir)

        scratch_path = manager.stage("new", real)

        assert not scratch_path.exists()
        assert scratch_path.parent == scratch_dir


class TestCommit:
    def test_commit_moves_scratch_content_back_to_the_real_path(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real = _write_file(real_dir / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "scratch")
        scratch_path = manager.stage("manip", real)
        scratch_path.write_text("changed")

        manager.commit("manip")

        assert real.read_text() == "changed"

    def test_commit_all_commits_every_staged_workbook(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real_a = _write_file(real_dir / "a.xlsx", "a-original")
        real_b = _write_file(real_dir / "b.xlsx", "b-original")
        manager = ScratchManager(tmp_path / "scratch")
        manager.stage("a", real_a).write_text("a-changed")
        manager.stage("b", real_b).write_text("b-changed")

        manager.commit_all()

        assert real_a.read_text() == "a-changed"
        assert real_b.read_text() == "b-changed"

    def test_commit_creates_the_real_file_when_it_did_not_exist_before(self, tmp_path: Path) -> None:
        real = tmp_path / "real" / "new.xlsx"
        manager = ScratchManager(tmp_path / "scratch")
        scratch_path = manager.stage("new", real)
        scratch_path.write_text("brand new content")

        manager.commit("new")

        assert real.read_text() == "brand new content"


class TestCleanup:
    def test_wipes_the_scratch_dir_after_a_successful_commit_all(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real = _write_file(real_dir / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "scratch")
        manager.stage("manip", real)
        manager.commit_all()

        manager.cleanup(keep_on_failure=False)

        assert not (tmp_path / "scratch").exists()

    def test_keeps_the_scratch_dir_by_default_if_nothing_was_committed(self, tmp_path: Path) -> None:
        real_dir = tmp_path / "real"
        real_dir.mkdir()
        real = _write_file(real_dir / "manip.xlsx", "original")
        manager = ScratchManager(tmp_path / "scratch")
        manager.stage("manip", real)

        manager.cleanup()  # keep_on_failure defaults to True

        assert (tmp_path / "scratch").exists()

    def test_cleanup_on_a_scratch_dir_that_was_never_created_does_not_raise(
        self, tmp_path: Path
    ) -> None:
        manager = ScratchManager(tmp_path / "never_created")
        manager.cleanup(keep_on_failure=False)  # should not raise
