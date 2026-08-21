"""Unit tests for the `delete_sheet` action."""

from excel_runner.actions import delete_sheet as delete_sheet_action
from excel_runner.core import ACTION_CAPABILITIES, ACTION_WRITES, WorkbookSession


class TestDeleteSheetAction:
    def test_registers_as_a_file_action_that_writes(self) -> None:
        assert ACTION_CAPABILITIES["delete_sheet"] == "file"
        assert ACTION_WRITES["delete_sheet"] is True

    def test_removes_the_sheet(self, file_session: WorkbookSession) -> None:
        file_session.handle.create_sheet("Data")
        result = delete_sheet_action(session=file_session, sheet="Data")
        assert result.status == "success"
        assert "Data" not in file_session.handle.sheetnames
        assert file_session.dirty is True

    def test_deleting_the_only_sheet_returns_a_structured_error(self, file_session: WorkbookSession) -> None:
        result = delete_sheet_action(session=file_session, sheet="Summary")
        assert result.status == "error"
        assert result.error is not None
        assert "only sheet" in result.error.message
