"""Unit tests for the `find_headers_row` action (PRD sec 7/sec 11 item 14)."""

from excel_runner.actions import find_headers_row as find_headers_row_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestFindHeadersRowAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["find_headers_row"] == "file"

    def test_output_shape_matches_prd_10_4(self, richer_file_session: WorkbookSession) -> None:
        result = find_headers_row_action(
            session=richer_file_session, sheet="Summary", search_range="A1:C4", patterns=["Region", "Total"]
        )
        assert result.status == "success"
        assert result.output == {"row": 2, "headers": {"Region": "A", "Total": "B"}}

    def test_no_match_returns_a_structured_error(self, richer_file_session: WorkbookSession) -> None:
        result = find_headers_row_action(
            session=richer_file_session, sheet="Summary", search_range="A1:C4", patterns=["NoSuchColumn"]
        )
        assert result.status == "error"
        assert result.error is not None
