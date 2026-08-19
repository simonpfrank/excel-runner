"""Backend primitives: openpyxl (file) and, later, xlwings (COM). See Spec sec 3.

Plain functions, not classes, one per primitive operation, so swapping or testing either side
never requires touching action code (PRD sec 6.1). File-backend functions are unprefixed;
COM-backend functions (added in a later increment) will carry a `com_` prefix to keep the two
sides unambiguous within one file once both exist.
"""

from typing import Any, Literal

import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.workbook.workbook import Workbook

_SingleCell = (Cell, ReadOnlyCell)


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
