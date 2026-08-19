"""Backend primitives: openpyxl (file) and, later, xlwings (COM). See Spec sec 3.

Plain functions, not classes, one per primitive operation, so swapping or testing either side
never requires touching action code (PRD sec 6.1). File-backend functions are unprefixed;
COM-backend functions (added in a later increment) will carry a `com_` prefix to keep the two
sides unambiguous within one file once both exist.
"""

import re
import shutil
from typing import Any, Literal

import openpyxl
import xlwings as xw
from openpyxl.cell.cell import Cell
from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.workbook import Workbook

_SingleCell = (Cell, ReadOnlyCell)
_WHOLE_COLUMN_RE = re.compile(r"^([A-Za-z]+):([A-Za-z]+)$")
_WHOLE_ROW_RE = re.compile(r"^(\d+):(\d+)$")


def open_workbook(path: str, mode: Literal["read_only", "read_write"]) -> Workbook:
    """Open an existing workbook file.

    Creating a workbook that doesn't exist yet is a `workbooks:` registry concern
    (`create_if_missing`), handled by session management once it exists — not this function.

    Args:
        path: Path to an existing workbook file.
        mode: "read_only" opens without allowing writes; "read_write" allows them.

    Returns:
        The opened openpyxl Workbook.

    Raises:
        FileNotFoundError: If path does not exist.
    """
    return openpyxl.load_workbook(path, read_only=(mode == "read_only"))


def create_workbook(path: str, template_path: str | None = None) -> None:
    """Create a new workbook file — blank, or copied from a template.

    Args:
        path: Where to create the new workbook.
        template_path: If given, copy this existing workbook's content instead of a blank one.
    """
    if template_path is not None:
        shutil.copy2(template_path, path)
    else:
        openpyxl.Workbook().save(path)


def save_workbook(workbook: Workbook, path: str) -> None:
    """Save a workbook to the given path.

    Args:
        workbook: The workbook to save.
        path: Destination path.
    """
    workbook.save(path)


def close_workbook(workbook: Workbook) -> None:
    """Close a workbook, releasing its file handle.

    Args:
        workbook: The workbook to close.
    """
    workbook.close()


def read_range(workbook: Workbook, sheet: str, range: str) -> Any:
    """Read a cell or range of cells.

    Args:
        workbook: The workbook to read from.
        sheet: Worksheet name.
        range: An A1-style cell (e.g. "B2") or range (e.g. "A1:D50").

    Returns:
        The cell's value for a single cell, or a 2D list of row values for a range.
    """
    selection = workbook[sheet][range]
    if isinstance(selection, _SingleCell):
        return selection.value
    return [[cell.value for cell in row] for row in selection]


def write_cell(workbook: Workbook, sheet: str, cell: str, value: Any) -> None:
    """Write a value to a single cell.

    A value starting with "=" is stored as a formula automatically — openpyxl's normal
    behavior, not special-cased here (PRD sec 6.5's "formulas need no dedicated action").

    Args:
        workbook: The workbook to write to.
        sheet: Worksheet name.
        cell: An A1-style cell reference (e.g. "B2").
        value: The value to write.
    """
    workbook[sheet][cell] = value


def write_range(workbook: Workbook, sheet: str, range: str, values: list[list[Any]]) -> None:
    """Write a 2D block of values, anchored at the top-left cell of `range`.

    The block's size comes from `values`, not from parsing the full extent of `range` — so
    `range` only needs its top-left cell to be correct (PRD sec 11 item 8).

    Args:
        workbook: The workbook to write to.
        sheet: Worksheet name.
        range: An A1-style cell or range — only the top-left cell is used as the anchor.
        values: A 2D list of row values to write.
    """
    anchor = workbook[sheet][range.split(":")[0]]
    start_row, start_col = anchor.row, anchor.column
    for row_offset, row_values in enumerate(values):
        for col_offset, value in enumerate(row_values):
            workbook[sheet].cell(row=start_row + row_offset, column=start_col + col_offset, value=value)


def set_column_width(workbook: Workbook, sheet: str, columns: str, width: float | Literal["autofit"]) -> None:
    """Set the width of a column or range of columns.

    "autofit" approximates Excel's real autofit (which needs a rendering engine openpyxl
    doesn't have) by sizing to the longest value currently in each column, plus padding.

    Args:
        workbook: The workbook to modify.
        sheet: Worksheet name.
        columns: A single column letter (e.g. "B") or range (e.g. "A:C").
        width: An explicit width, or "autofit".
    """
    worksheet = workbook[sheet]
    start_letter, _, end_letter = columns.partition(":")
    end_letter = end_letter or start_letter
    for idx in range(column_index_from_string(start_letter), column_index_from_string(end_letter) + 1):
        letter = get_column_letter(idx)
        if width == "autofit":
            lengths = [len(str(cell.value)) for cell in worksheet[letter] if cell.value is not None]
            worksheet.column_dimensions[letter].width = max(lengths, default=8) + 2
        else:
            worksheet.column_dimensions[letter].width = width


def insert_range(
    workbook: Workbook,
    sheet: str,
    at: str,
    direction: Literal["rows", "columns"] | None = None,
    header: dict[str, Any] | None = None,
) -> None:
    """Insert a whole row or whole column, shifting existing content.

    Args:
        workbook: The workbook to modify.
        sheet: Worksheet name.
        at: A whole-column reference (e.g. "C:C") or whole-row reference (e.g. "5:5").
        direction: Unused for whole-row/whole-column inserts — the direction is unambiguous
            from `at` itself (PRD sec 11 item 12). Accepted so partial-range calls (below)
            still reach the clear error instead of an unrelated TypeError.
        header: `{"row": int, "text": str}` — only meaningful for a column insert.

    Raises:
        NotImplementedError: If `at` is a partial range. Whole-row/whole-column insert is
            native to openpyxl and cheap; a true partial-range insert needs hand-rolled
            cell-shifting logic not built yet (PRD sec 11 item 12's flagged cost) — this must
            fail clearly rather than silently do the wrong thing.
    """
    worksheet = workbook[sheet]
    column_match = _WHOLE_COLUMN_RE.match(at)
    row_match = _WHOLE_ROW_RE.match(at)
    if column_match:
        idx = column_index_from_string(column_match.group(1))
        worksheet.insert_cols(idx)
        if header:
            worksheet.cell(row=header["row"], column=idx, value=header["text"])
    elif row_match:
        worksheet.insert_rows(int(row_match.group(1)))
    else:
        raise NotImplementedError(
            f'insert_range: partial range "{at}" is not supported yet — only whole-row '
            '("5:5") or whole-column ("C:C") ranges are built.'
        )


def copy_range(
    source_workbook: Workbook,
    source_sheet: str,
    source_range: str | None,
    target_workbook: Workbook,
    target_sheet: str,
    target_range: str,
) -> None:
    """Copy a range — or, if `source_range` is None, the whole used area of the sheet — from
    one workbook to another, anchored at `target_range`'s top-left cell.

    Args:
        source_workbook: The workbook to copy from.
        source_sheet: Source worksheet name.
        source_range: An A1-style range, or None to copy the whole sheet.
        target_workbook: The workbook to copy into.
        target_sheet: Target worksheet name.
        target_range: Where to start writing — only the top-left cell is used.
    """
    source_worksheet = source_workbook[source_sheet]
    if source_range is None:
        values = [[cell.value for cell in row] for row in source_worksheet.iter_rows()]
    else:
        selection = source_worksheet[source_range]
        values = [[selection.value]] if isinstance(selection, _SingleCell) else [
            [cell.value for cell in row] for row in selection
        ]
    write_range(target_workbook, target_sheet, target_range, values)


def find_headers_row(
    workbook: Workbook, sheet: str, search_range: str, patterns: list[str]
) -> tuple[int, dict[str, str]] | None:
    """Find the row within `search_range` where every pattern matches some cell in that row.

    Args:
        workbook: The workbook to search.
        sheet: Worksheet name.
        search_range: An A1-style range to search within.
        patterns: Regex patterns — every one must match a cell in a row for that row to count.

    Returns:
        `(row_number, {pattern: column_letter})` for the first matching row, or None.
    """
    selection = workbook[sheet][search_range]
    rows = [(selection,)] if isinstance(selection, _SingleCell) else selection
    for row in rows:
        matches: dict[str, str] = {}
        for pattern in patterns:
            for cell in row:
                if cell.value is not None and re.search(pattern, str(cell.value)):
                    matches[pattern] = get_column_letter(cell.column)
                    break
        if len(matches) == len(patterns):
            return row[0].row, matches
    return None


def find_row(
    workbook: Workbook, sheet: str, column: str, search_value: Any, header_row: int | None = None
) -> int | None:
    """Find the row number where `column` equals `search_value`.

    Args:
        workbook: The workbook to search.
        sheet: Worksheet name.
        column: A column letter (e.g. "B").
        search_value: The value to match, by equality.
        header_row: If given, search starts on the row after it.

    Returns:
        The matching row number, or None if not found.
    """
    worksheet = workbook[sheet]
    col_idx = column_index_from_string(column)
    start_row = (header_row or 0) + 1
    for (cell,) in worksheet.iter_rows(min_row=start_row, min_col=col_idx, max_col=col_idx):
        if cell.value == search_value:
            return cell.row
    return None


def find_column(workbook: Workbook, sheet: str, header_row: int, pattern: str) -> str | None:
    """Find the column letter whose header (in `header_row`) matches `pattern`.

    Args:
        workbook: The workbook to search.
        sheet: Worksheet name.
        header_row: The row number containing headers.
        pattern: A regex pattern to match against each header cell's value.

    Returns:
        The matching column letter, or None if not found.
    """
    for cell in workbook[sheet][header_row]:
        if cell.value is not None and re.search(pattern, str(cell.value)):
            return get_column_letter(cell.column)
    return None


def find_columns(
    workbook: Workbook, sheet: str, header_row: int, patterns: dict[str, str]
) -> dict[str, str]:
    """Find multiple named columns by header pattern in one call.

    Args:
        workbook: The workbook to search.
        sheet: Worksheet name.
        header_row: The row number containing headers.
        patterns: Logical name to regex pattern.

    Returns:
        Logical name to column letter, for every pattern that matched. Names whose pattern
        didn't match anything are simply absent, not an error at this layer.
    """
    result: dict[str, str] = {}
    for name, pattern in patterns.items():
        column = find_column(workbook, sheet, header_row, pattern)
        if column is not None:
            result[name] = column
    return result


def read_properties(workbook: Workbook) -> dict[str, Any]:
    """Read a workbook's document properties (title, creator, etc.).

    Args:
        workbook: The workbook to read from.

    Returns:
        Mapping of property name to value, for every non-None standard property.
    """
    properties = workbook.properties
    return {
        name: value
        for name in vars(properties)
        if not name.startswith("_") and (value := getattr(properties, name)) is not None
    }


def read_cells(workbook: Workbook, sheet: str, cells: list[str]) -> dict[str, Any]:
    """Read a scattered list of specific cells.

    Args:
        workbook: The workbook to read from.
        sheet: Worksheet name.
        cells: A1-style cell references to read.

    Returns:
        Mapping of cell reference to its value.
    """
    worksheet = workbook[sheet]
    return {cell: worksheet[cell].value for cell in cells}


# --- COM (xlwings) — owned-instance tracking (PRD sec 6.2.1, Spec sec 3.1) -----------------


class OwnedInstanceRegistry:
    """Tracks Excel App instances a run has spawned itself, so cleanup only ever touches
    instances it owns.

    xlwings (and the underlying Excel COM/Apple Event API) will happily attach to whatever
    Excel instance is already running rather than spawning its own — dangerous when more than
    one automation run, or a real user's own Excel session, can be active at once. Every
    instance this registry hands out comes from a fresh `xw.App(...)` call, never an
    unqualified lookup like `xw.apps.active` or a bare `xw.Book("name.xlsx")` that could bind
    to an instance it doesn't own.
    """

    def __init__(self) -> None:
        self._owned: dict[int, xw.App] = {}

    @property
    def pids(self) -> tuple[int, ...]:
        """Process IDs of every instance currently owned — the run's audit trail of what it
        spawned (PRD sec 6.2.1's "records that instance's process ID against the run")."""
        return tuple(self._owned)

    def spawn(self, visible: bool = False) -> xw.App:
        """Spawn a brand-new, dedicated Excel App instance and start tracking it.

        Args:
            visible: Whether the spawned Excel window is shown. Defaults to hidden, the normal
                automation case; a visible instance is for a deliberately different use (e.g.
                the parked "replay nice" idea, PRD sec 12), not the default run path.

        Returns:
            The newly spawned App.
        """
        app = xw.App(visible=visible, add_book=False)
        self._owned[app.pid] = app
        return app

    def close_owned(self) -> None:
        """Quit every owned instance, attempting all of them even if some fail.

        Raises:
            ExceptionGroup: If one or more instances failed to quit. Every instance still gets
                a quit attempt regardless (PRD sec 6.3's crash-safety requirement), mirroring
                `engine.SessionManager.close_all()` — this isn't a defensive catch-and-ignore,
                every failure is still surfaced, just after giving every other instance a
                chance to close too.
        """
        errors: list[Exception] = []
        for app in self._owned.values():
            try:
                app.quit()
            except Exception as exc:  # noqa: BLE001 - intentional, see docstring
                errors.append(exc)
        self._owned.clear()
        if errors:
            raise ExceptionGroup("failed to quit one or more owned Excel instances", errors)
