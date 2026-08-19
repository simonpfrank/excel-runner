"""Unit tests for backends.py's second batch of file-backend primitives (Spec sec 3):
write_range, set_column_width, insert_range (whole row/column only — see module docstring
in excel_runner/backends.py for why partial-range insert isn't built yet), copy_range.
"""

from pathlib import Path

import openpyxl
import pytest

from excel_runner import backends


def _make_workbook(tmp_path: Path, name: str = "fixture.xlsx") -> Path:
    path = tmp_path / name
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


class TestWriteRange:
    def test_writes_a_2d_block_anchored_at_the_top_left(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.write_range(workbook, "Summary", "D1:E2", [[1, 2], [3, 4]])
        assert backends.read_range(workbook, "Summary", "D1:E2") == [[1, 2], [3, 4]]
        workbook.close()

    def test_single_cell_range_still_works(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.write_range(workbook, "Summary", "D1", [["x"]])
        assert backends.read_range(workbook, "Summary", "D1") == "x"
        workbook.close()


class TestSetColumnWidth:
    def test_sets_an_explicit_width(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.set_column_width(workbook, "Summary", "A:A", 25)
        assert workbook["Summary"].column_dimensions["A"].width == 25
        workbook.close()

    def test_sets_a_range_of_columns(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.set_column_width(workbook, "Summary", "A:B", 15)
        assert workbook["Summary"].column_dimensions["A"].width == 15
        assert workbook["Summary"].column_dimensions["B"].width == 15
        workbook.close()

    def test_autofit_sizes_by_longest_content(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.set_column_width(workbook, "Summary", "A:A", "autofit")
        width = workbook["Summary"].column_dimensions["A"].width
        assert width is not None
        assert width > 0
        workbook.close()


class TestInsertRange:
    def test_inserts_a_whole_column(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.insert_range(workbook, "Summary", "B:B")
        # original B1 ("Total") should have shifted right to C1
        assert backends.read_range(workbook, "Summary", "C1") == "Total"
        workbook.close()

    def test_inserts_a_whole_column_with_a_header(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.insert_range(workbook, "Summary", "B:B", header={"row": 1, "text": "Flag"})
        assert backends.read_range(workbook, "Summary", "B1") == "Flag"
        workbook.close()

    def test_inserts_a_whole_row(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        backends.insert_range(workbook, "Summary", "2:2")
        # original row 2 ("North", 100) should have shifted down to row 3
        assert backends.read_range(workbook, "Summary", "A3") == "North"
        workbook.close()

    def test_partial_range_raises_a_clear_error(self, tmp_path: Path) -> None:
        """Partial-range insert-with-shift needs hand-rolled cell-shifting logic that isn't
        built yet (PRD sec 11 item 12's flagged cost) — must fail clearly, not silently do
        the wrong thing or pretend to succeed."""
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        with pytest.raises(NotImplementedError):
            backends.insert_range(workbook, "Summary", "B2:B5", direction="rows")
        workbook.close()


class TestCopyRange:
    def test_copies_an_explicit_range_between_workbooks(self, tmp_path: Path) -> None:
        source = backends.open_workbook(str(_make_workbook(tmp_path, "source.xlsx")), mode="read_write")
        target_path = _make_workbook(tmp_path, "target.xlsx")
        target = backends.open_workbook(str(target_path), mode="read_write")

        backends.copy_range(source, "Summary", "A1:B2", target, "Summary", "D1")

        assert backends.read_range(target, "Summary", "D1:E2") == [["Region", "Total"], ["North", 100]]
        source.close()
        target.close()

    def test_copies_the_whole_sheet_when_source_range_is_none(self, tmp_path: Path) -> None:
        source = backends.open_workbook(str(_make_workbook(tmp_path, "source.xlsx")), mode="read_write")
        target_path = _make_workbook(tmp_path, "target.xlsx")
        target = backends.open_workbook(str(target_path), mode="read_write")

        backends.copy_range(source, "Summary", None, target, "Summary", "D1")

        assert backends.read_range(target, "Summary", "D1") == "Region"
        assert backends.read_range(target, "Summary", "E2") == 100
        source.close()
        target.close()
