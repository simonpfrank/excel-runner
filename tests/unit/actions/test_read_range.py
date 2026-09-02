"""Unit tests for the `read_range` action (Spec sec 7 catalog).

`as: formulas` is deliberately left out of this increment's signature — it depends on which
data_only flag the workbook was opened with, a session-level decision that doesn't exist until
static/dry-run validation (Spec sec 5.4) is built.
"""

from excel_runner.actions import read_range as read_range_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestReadRangeAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["read_range"] == "file"

    def test_output_is_keyed_under_values_per_prd_10_4(
        self, file_session: WorkbookSession
    ) -> None:
        result = read_range_action(session=file_session, sheet="Summary", range="A1")
        assert result.output == {"values": "Region"}

    def test_reads_a_multi_cell_range(self, file_session: WorkbookSession) -> None:
        result = read_range_action(session=file_session, sheet="Summary", range="A1:B2")
        assert result.output == {"values": [["Region", "Total"], ["North", 100]]}

    def test_returns_success_status(self, file_session: WorkbookSession) -> None:
        result = read_range_action(session=file_session, sheet="Summary", range="A1")
        assert result.status == "success"

    def test_an_explicit_list_of_sheets_is_keyed_by_sheet_name(
        self, multi_sheet_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=multi_sheet_file_session,
            sheet=["A&H North", "A&H South"],
            range="A1",
        )
        assert result.output == {
            "values": {"A&H North": "north-value", "A&H South": "south-value"}
        }

    def test_all_reads_every_sheet_in_the_workbook(
        self, multi_sheet_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=multi_sheet_file_session, sheet="all", range="A1"
        )
        assert result.output == {
            "values": {
                "A&H North": "north-value",
                "A&H South": "south-value",
                "Other": "other-value",
            }
        }

    def test_matching_reads_every_sheet_whose_name_matches_the_regex(
        self, multi_sheet_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=multi_sheet_file_session, sheet={"matching": "^A&H"}, range="A1"
        )
        assert result.output == {
            "values": {"A&H North": "north-value", "A&H South": "south-value"}
        }
