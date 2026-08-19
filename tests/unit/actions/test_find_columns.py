"""Unit tests for the `find_columns` action (PRD sec 7/sec 11 item 16)."""

from excel_runner.actions import find_columns as find_columns_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestFindColumnsAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["find_columns"] == "file"

    def test_output_shape_matches_prd_10_4(self, richer_file_session: WorkbookSession) -> None:
        result = find_columns_action(
            session=richer_file_session,
            sheet="Summary",
            header_row=2,
            patterns={"region": "Region", "status": "Status"},
        )
        assert result.status == "success"
        assert result.output == {"region": "A", "status": "C"}

    def test_unmatched_names_are_simply_omitted_not_an_error(
        self, richer_file_session: WorkbookSession
    ) -> None:
        result = find_columns_action(
            session=richer_file_session, sheet="Summary", header_row=2, patterns={"missing": "NoSuchColumn"}
        )
        assert result.status == "success"
        assert result.output == {}
