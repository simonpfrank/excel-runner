"""Unit tests for the `rename_sheet` action."""

from excel_runner.actions import rename_sheet as rename_sheet_action
from excel_runner.core import ACTION_CAPABILITIES, ACTION_WRITES, WorkbookSession


class TestRenameSheetAction:
    def test_registers_as_a_file_action_that_writes(self) -> None:
        assert ACTION_CAPABILITIES["rename_sheet"] == "file"
        assert ACTION_WRITES["rename_sheet"] is True

    def test_renames_the_sheet(self, file_session: WorkbookSession) -> None:
        result = rename_sheet_action(session=file_session, sheet="Summary", new_name="Overview")
        assert result.status == "success"
        assert "Overview" in file_session.handle.sheetnames
        assert file_session.dirty is True
