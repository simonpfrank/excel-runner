"""Tests for text-file, replacement, and table actions."""

from pathlib import Path

import openpyxl
import pytest

from excel_runner.actions import (
    copy_table_columns,
    read_table,
    read_text_file,
    replace_in_range,
    replace_table_text,
    replace_text,
    update_table_cells,
    write_cell,
    write_range,
)
from excel_runner.core import ActionExecutionError, WorkbookSession


def _table_session(tmp_path: Path) -> WorkbookSession:
    path = tmp_path / "table.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "BASIS"
    sheet.append(["ignore"])
    sheet.append([None])
    sheet.append([None])
    sheet.append([None])
    sheet.append([None])
    sheet.append([None])
    sheet.append([None, "BASIS_ITEM", "SOURCE", "TARGET"])
    sheet.append([None, "ACC_RATIO", "old_202512", ""])
    sheet.append([None, "ECO_TBL", "value", ""])
    other = workbook.create_sheet("Other")
    other["A1"] = "old_202512"
    workbook.defined_names["TargetCell"] = openpyxl.workbook.defined_name.DefinedName(
        "TargetCell", attr_text="BASIS!$D$8"
    )
    workbook.save(path)
    from excel_runner import backends

    return WorkbookSession("table", "file", backends.open_workbook(str(path), "read_write"), str(path), "read_write")


def test_read_text_file_preserves_strings_and_csv_quotes(tmp_path: Path) -> None:
    path = tmp_path / "input.fac"
    path.write_text('id,name\n00123,"one,two"\n')
    result = read_text_file(str(path))
    assert result.output == {"values": [["id", "name"], ["00123", "one,two"]]}


def test_read_text_file_without_delimiter_returns_one_column(tmp_path: Path) -> None:
    path = tmp_path / "input.log"
    path.write_text("first\n\nsecond\n")
    assert read_text_file(str(path)).output == {"values": [["first"], ["second"]]}


@pytest.mark.parametrize("delimiter,quotechar", [("", '"'), (",,", '"'), (",", "")])
def test_read_text_file_rejects_invalid_csv_options(
    tmp_path: Path, delimiter: str, quotechar: str
) -> None:
    path = tmp_path / "input.csv"
    path.write_text("header\n")
    with pytest.raises(ActionExecutionError):
        read_text_file(str(path), delimiter=delimiter, quotechar=quotechar)


def test_read_text_file_rejects_malformed_csv(tmp_path: Path) -> None:
    path = tmp_path / "input.csv"
    path.write_text('header\n"unclosed\n')
    with pytest.raises(ActionExecutionError):
        read_text_file(str(path))


def test_write_actions_accept_a_defined_name(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    write_cell(session, "BASIS", "TargetCell", "hello")
    write_range(session, "BASIS", "TargetCell", [["range"]])
    assert session.handle["BASIS"]["D8"].value == "range"


def test_write_cell_rejects_a_multi_cell_defined_name(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    session.handle.defined_names["TargetRange"] = openpyxl.workbook.defined_name.DefinedName(
        "TargetRange", attr_text="BASIS!$D$8:$D$9"
    )
    with pytest.raises(ValueError, match="single cell"):
        write_cell(session, "BASIS", "TargetRange", "hello")


def test_replacements_use_sheet_selection_and_range(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    assert replace_text(session, {"matching": "^(BASIS|Other)$"}, "old_\\d{6}", "new").output == {"replacements": 2}
    assert replace_in_range(session, "BASIS", "C8", "new", "final").output == {"replacements": 1}
    assert session.handle["Other"]["A1"].value == "new"
    assert session.handle["BASIS"]["C8"].value == "final"


def test_replace_in_range_leaves_blank_cells_unchanged(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    result = replace_in_range(session, "BASIS", "D8:D9", "None", "changed")
    assert result.output == {"replacements": 0}
    assert session.handle["BASIS"]["D8"].value is None


def test_table_actions_discover_copy_update_and_replace(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    table = read_table(session, "BASIS", "B7")
    assert table.output["headers"] == ["BASIS_ITEM", "SOURCE", "TARGET"]
    copy_table_columns(session, "BASIS", "B7", ["SOURCE"], ["TARGET"])
    update_table_cells(session, "BASIS", "B7", "BASIS_ITEM", ["ECO_TBL"], ["TARGET"], "set")
    replace_table_text(session, "BASIS", "B7", "BASIS_ITEM", ["ACC_RATIO"], ["TARGET"], "\\d{6}", "202606")
    assert session.handle["BASIS"]["D8"].value == "old_202606"
    assert session.handle["BASIS"]["D9"].value == "set"


def test_table_actions_reject_case_insensitive_duplicate_headers(tmp_path: Path) -> None:
    session = _table_session(tmp_path)
    session.handle["BASIS"]["D7"] = "source"
    with pytest.raises(ActionExecutionError, match="unique"):
        read_table(session, "BASIS", "B7")