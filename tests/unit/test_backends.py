"""Unit tests for the file-backend (openpyxl) primitives in excel_runner.backends (Spec sec 3).

These exercise real openpyxl against throwaway files rather than mocking it — openpyxl needs
no live Excel, so this is a real dependency being tested cheaply, same spirit as the project's
"no mocks in integration tests" convention applied at the unit level (Spec sec 7).
"""

from pathlib import Path

import openpyxl
import pytest

from excel_runner import backends


def _make_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "fixture.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    sheet["A1"] = "Region"
    sheet["B1"] = "Total"
    sheet["A2"] = "North"
    sheet["B2"] = 100
    workbook.save(path)
    return path


class TestOpenWorkbook:
    def test_opens_read_write_by_default(self, tmp_path: Path) -> None:
        path = _make_workbook(tmp_path)
        workbook = backends.open_workbook(str(path), mode="read_write")
        assert workbook["Summary"]["A1"].value == "Region"
        workbook.close()

    def test_opens_read_only(self, tmp_path: Path) -> None:
        path = _make_workbook(tmp_path)
        workbook = backends.open_workbook(str(path), mode="read_only")
        assert workbook["Summary"]["A1"].value == "Region"
        workbook.close()


class TestReadRange:
    def test_reads_a_single_cell_as_a_scalar(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.read_range(workbook, "Summary", "A1") == "Region"
        workbook.close()

    def test_reads_a_multi_cell_range_as_a_2d_list(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.read_range(workbook, "Summary", "A1:B2") == [
            ["Region", "Total"],
            ["North", 100],
        ]
        workbook.close()

    def test_reads_a_single_row_range_as_a_2d_list(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.read_range(workbook, "Summary", "A1:B1") == [["Region", "Total"]]
        workbook.close()


class TestWriteCell:
    def test_writes_a_value(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.write_cell(workbook, "Summary", "C1", "Status")
        assert workbook["Summary"]["C1"].value == "Status"
        workbook.close()

    def test_writes_a_formula_string_as_is(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.write_cell(workbook, "Summary", "C2", "=SUM(B2:B2)")
        assert workbook["Summary"]["C2"].value == "=SUM(B2:B2)"
        workbook.close()


class TestSaveWorkbook:
    def test_saves_changes_to_the_given_path(self, tmp_path: Path) -> None:
        source = _make_workbook(tmp_path)
        workbook = backends.open_workbook(str(source), mode="read_write")
        backends.write_cell(workbook, "Summary", "A1", "Changed")
        out_path = tmp_path / "saved.xlsx"
        backends.save_workbook(workbook, str(out_path))
        workbook.close()

        reopened = backends.open_workbook(str(out_path), mode="read_only")
        assert backends.read_range(reopened, "Summary", "A1") == "Changed"
        reopened.close()


class TestCloseWorkbook:
    def test_close_does_not_raise(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.close_workbook(workbook)  # should not raise

    def test_open_missing_file_raises_a_clear_file_not_found_error(self, tmp_path: Path) -> None:
        with pytest.raises(FileNotFoundError):
            backends.open_workbook(str(tmp_path / "does_not_exist.xlsx"), mode="read_write")
