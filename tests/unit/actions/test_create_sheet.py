"""Unit tests for the `create_sheet` action."""

from excel_runner.actions import create_sheet as create_sheet_action
from excel_runner.core import ACTION_CAPABILITIES, ACTION_WRITES, WorkbookSession


class TestCreateSheetAction:
    def test_registers_as_a_file_action_that_writes(self) -> None:
        assert ACTION_CAPABILITIES["create_sheet"] == "file"
        assert ACTION_WRITES["create_sheet"] is True

    def test_adds_a_new_sheet(self, file_session: WorkbookSession) -> None:
        result = create_sheet_action(session=file_session, name="Data")
        assert result.status == "success"
        assert "Data" in file_session.handle.sheetnames
        assert file_session.dirty is True

    def test_duplicate_name_returns_a_structured_error(
        self, file_session: WorkbookSession
    ) -> None:
        result = create_sheet_action(session=file_session, name="Summary")
        assert result.status == "error"
        assert result.error is not None
        assert "Summary" in result.error.message
