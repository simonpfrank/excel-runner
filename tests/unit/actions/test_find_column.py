"""Unit tests for the `find_column` action (PRD sec 7)."""

from excel_runner.actions import find_column as find_column_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestFindColumnAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["find_column"] == "file"

    def test_output_shape_matches_prd_10_4(self, richer_file_session: WorkbookSession) -> None:
        result = find_column_action(session=richer_file_session, sheet="Summary", header_row=2, pattern="Status")
        assert result.status == "success"
        assert result.output == {"column": "C"}

    def test_not_found_returns_a_structured_error(self, richer_file_session: WorkbookSession) -> None:
        result = find_column_action(
            session=richer_file_session, sheet="Summary", header_row=2, pattern="NoSuchColumn"
        )
        assert result.status == "error"
        assert result.error is not None
