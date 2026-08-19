"""Unit tests for the `find_row` action (PRD sec 7)."""

from excel_runner.actions import find_row as find_row_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestFindRowAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["find_row"] == "file"

    def test_output_shape_matches_prd_10_4(self, richer_file_session: WorkbookSession) -> None:
        result = find_row_action(
            session=richer_file_session, sheet="Summary", column="A", search_value="South", header_row=2
        )
        assert result.status == "success"
        assert result.output == {"row": 4}

    def test_not_found_returns_a_structured_error(self, richer_file_session: WorkbookSession) -> None:
        result = find_row_action(
            session=richer_file_session, sheet="Summary", column="A", search_value="West", header_row=2
        )
        assert result.status == "error"
        assert result.error is not None
