"""Unit tests for backends.py's find_* primitives and read_metadata (Spec sec 3)."""

from pathlib import Path

import openpyxl

from excel_runner import backends


def _make_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "fixture.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    sheet["A1"] = "Notes"
    sheet["A2"] = "Region"
    sheet["B2"] = "Total"
    sheet["C2"] = "Status"
    sheet["A3"] = "North"
    sheet["B3"] = 100
    sheet["C3"] = "PASS"
    sheet["A4"] = "South"
    sheet["B4"] = 200
    sheet["C4"] = "FAIL"
    workbook.properties.title = "Q1 Report"
    workbook.properties.creator = "Simon"
    workbook.save(path)
    return path


class TestFindHeadersRow:
    def test_finds_the_row_matching_all_patterns(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.find_headers_row(workbook, "Summary", "A1:C4", ["Region", "Total"])
        assert result == (2, {"Region": "A", "Total": "B"})
        workbook.close()

    def test_returns_none_when_no_row_matches(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.find_headers_row(workbook, "Summary", "A1:C4", ["NoSuchColumn"])
        assert result is None
        workbook.close()

    def test_single_cell_search_range_does_not_crash(self, tmp_path: Path) -> None:
        """Regression test: a single-cell search_range must still iterate as one row of one
        cell, not the bare cell itself (mypy caught this as a real runtime bug, not just a
        type-annotation nit — no earlier test exercised this shape)."""
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.find_headers_row(workbook, "Summary", "A2", ["Region"])
        assert result == (2, {"Region": "A"})
        workbook.close()


class TestFindRow:
    def test_finds_a_matching_value(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_row(workbook, "Summary", "A", "South", header_row=2) == 4
        workbook.close()

    def test_returns_none_when_not_found(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_row(workbook, "Summary", "A", "West", header_row=2) is None
        workbook.close()

    def test_works_without_a_header_row(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_row(workbook, "Summary", "A", "North") == 3
        workbook.close()


class TestFindColumn:
    def test_finds_a_column_by_exact_pattern(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_column(workbook, "Summary", 2, "Status") == "C"
        workbook.close()

    def test_finds_a_column_by_regex_pattern(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_column(workbook, "Summary", 2, "Tot.*") == "B"
        workbook.close()

    def test_returns_none_when_not_found(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        assert backends.find_column(workbook, "Summary", 2, "NoSuchColumn") is None
        workbook.close()


class TestFindColumns:
    def test_finds_multiple_named_columns(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.find_columns(
            workbook, "Summary", 2, {"region": "Region", "status": "Status"}
        )
        assert result == {"region": "A", "status": "C"}
        workbook.close()

    def test_omits_unmatched_names(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.find_columns(workbook, "Summary", 2, {"missing": "NoSuchColumn"})
        assert result == {}
        workbook.close()


class TestReadMetadataProperties:
    def test_reads_document_properties(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.read_properties(workbook)
        assert result["title"] == "Q1 Report"
        assert result["creator"] == "Simon"
        workbook.close()


class TestReadMetadataCells:
    def test_reads_specific_scattered_cells(self, tmp_path: Path) -> None:
        workbook = backends.open_workbook(str(_make_workbook(tmp_path)), mode="read_write")
        result = backends.read_cells(workbook, "Summary", ["A1", "B3"])
        assert result == {"A1": "Notes", "B3": 100}
        workbook.close()
