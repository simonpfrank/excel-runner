"""Unit tests for the `write_cell` action (Spec sec 7 catalog)."""

from excel_runner import backends
from excel_runner.actions import write_cell as write_cell_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession
from openpyxl.workbook.defined_name import DefinedName
import pytest


class TestWriteCellAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["write_cell"] == "file"

    def test_writes_the_value(self, file_session: WorkbookSession) -> None:
        write_cell_action(session=file_session, sheet="Summary", cell="C1", value="Status")
        assert backends.read_range(file_session.handle, "Summary", "C1") == "Status"

    def test_writes_a_formula_string_unchanged(self, file_session: WorkbookSession) -> None:
        write_cell_action(session=file_session, sheet="Summary", cell="C2", value="=SUM(B2:B2)")
        assert backends.read_range(file_session.handle, "Summary", "C2") == "=SUM(B2:B2)"

    @pytest.mark.parametrize("sheet", [None, ""])
    def test_named_cell_target_allows_an_omitted_or_blank_sheet(
        self, file_session: WorkbookSession, sheet: str | None
    ) -> None:
        file_session.handle.defined_names.add(
            DefinedName("StatusCell", attr_text="Summary!$C$1")
        )

        write_cell_action(session=file_session, sheet=sheet, cell="StatusCell", value="Status")

        assert backends.read_range(file_session.handle, "Summary", "C1") == "Status"

    def test_has_no_meaningful_output(self, file_session: WorkbookSession) -> None:
        result = write_cell_action(session=file_session, sheet="Summary", cell="C1", value="x")
        assert result.output == {}
        assert result.status == "success"
