"""Unit tests for the `read_range` action (Spec sec 7 catalog)."""

import pytest

from excel_runner.actions import read_range as read_range_action
from excel_runner.core import ACTION_CAPABILITIES, ActionExecutionError, WorkbookSession
from tests.unit.conftest import requires_excel, requires_working_xlwings_save


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


class TestReadRangeNamedRange:
    def test_reads_via_a_workbook_defined_name(
        self, named_range_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=named_range_file_session, sheet="Summary", range="SalesTotal"
        )
        assert result.output == {"values": 100}

    def test_the_defined_names_own_sheet_wins_over_the_passed_sheet(
        self, named_range_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=named_range_file_session, sheet="DoesNotMatter", range="SalesTotal"
        )
        assert result.output == {"values": 100}

    def test_neither_valid_a1_nor_a_real_defined_name_raises_a_clear_error(
        self, named_range_file_session: WorkbookSession
    ) -> None:
        with pytest.raises(ActionExecutionError) as exc_info:
            read_range_action(
                session=named_range_file_session, sheet="Summary", range="NotARealRange"
            )
        assert "NotARealRange" in exc_info.value.detail.message


@requires_excel
@requires_working_xlwings_save
class TestReadRangeFormulaParam:
    def test_defaults_to_the_computed_value_for_a_formula_cell(
        self, formula_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=formula_file_session, sheet="Summary", range="B1"
        )
        assert result.output == {"values": 20}

    def test_formula_true_returns_the_formula_text_instead(
        self, formula_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=formula_file_session, sheet="Summary", range="B1", formula=True
        )
        assert result.output == {"values": "=A1*2"}

    def test_formula_false_is_the_same_as_the_default(
        self, formula_file_session: WorkbookSession
    ) -> None:
        result = read_range_action(
            session=formula_file_session, sheet="Summary", range="B1", formula=False
        )
        assert result.output == {"values": 20}
