"""Unit tests for the `save` action (Spec sec 7 catalog)."""

from pathlib import Path

from excel_runner import backends
from excel_runner.actions import save as save_action
from excel_runner.actions import write_cell as write_cell_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestSaveAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["save"] == "file"

    def test_saves_pending_changes_to_the_session_path(
        self, file_session: WorkbookSession
    ) -> None:
        write_cell_action(session=file_session, sheet="Summary", cell="A1", value="Changed")
        result = save_action(session=file_session)
        assert result.status == "success"

        file_session.handle.close()
        reopened = backends.open_workbook(file_session.path, mode="read_only")
        assert backends.read_range(reopened, "Summary", "A1") == "Changed"
        reopened.close()

    def test_saves_to_a_different_path_than_the_original_when_session_path_differs(
        self, file_session: WorkbookSession, tmp_path: Path
    ) -> None:
        """Confirms save() writes to session.path (not some other hardcoded path) — this is
        the seam the future scratch-copy execution model (PRD sec 6.3.1) will use: it just
        needs to point session.path at the scratch copy, save() doesn't change."""
        new_path = tmp_path / "moved.xlsx"
        file_session.path = str(new_path)
        save_action(session=file_session)
        assert new_path.exists()
