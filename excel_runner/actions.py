"""Action functions. See docs/Specification.md sec 4 and docs/PRD.md sec 7 for the catalog.

Every action has the shape `fn(session: WorkbookSession, **params) -> ActionResult`. Note
`workbook` is deliberately absent from every signature below, even though it's a required YAML
field on every step (PRD sec 7/sec 11): the (not-yet-built) runner resolves `workbook` into the
`session` it passes in, so the action itself never needs it as a separate parameter — passing
both would just be the same information twice.

Functions are grouped to match the PRD sec 7 catalog order (basic -> data -> ...) via comment
banners, since there's no file boundary to do it now that actions live in one module
(docs/Specification.md sec 4).
"""

from typing import Any

from excel_runner import backends
from excel_runner.core import ActionResult, WorkbookSession, file_action

# --- basic -------------------------------------------------------------------------------


@file_action
def open(session: WorkbookSession) -> ActionResult:
    """Confirm a workbook is open.

    The runner resolves and opens `session` before dispatching to any action (Spec sec 6.1) —
    this exists for explicit manual control and audit-log clarity, not to do the opening
    itself. `update_links` and a `mode` override are not yet parameters here: `update_links`
    has no effect without a live Excel session (COM, a later phase per PRD sec 8), and a mode
    override depends on the read/write inference that static validation (Spec sec 5.4) hasn't
    been built yet to override.

    Args:
        session: The already-open workbook session.

    Returns:
        A success result with no meaningful output.
    """
    return ActionResult(status="success", output={})


@file_action
def save(session: WorkbookSession) -> ActionResult:
    """Save the workbook to its current session path.

    Args:
        session: The workbook session to save.

    Returns:
        A success result with no meaningful output.
    """
    backends.save_workbook(session.handle, session.path)
    return ActionResult(status="success", output={})


@file_action
def close(session: WorkbookSession) -> ActionResult:
    """Close the workbook, releasing its file handle.

    Args:
        session: The workbook session to close.

    Returns:
        A success result with no meaningful output.
    """
    backends.close_workbook(session.handle)
    return ActionResult(status="success", output={})


# --- data ----------------------------------------------------------------------------------


@file_action
def read_range(session: WorkbookSession, sheet: str, range: str) -> ActionResult:
    """Read a cell or range of cells.

    `as: formulas` is not yet a parameter here — it depends on which `data_only` flag the
    workbook was opened with, a session-level decision not built until Spec sec 5.4.

    Args:
        session: The workbook session to read from.
        sheet: Worksheet name.
        range: An A1-style cell (e.g. "B2") or range (e.g. "A1:D50").

    Returns:
        `{"values": ...}` — the cell's value for a single cell, or a 2D list for a range
        (PRD sec 10.4's output-shape rule: always a keyed object).
    """
    values = backends.read_range(session.handle, sheet, range)
    return ActionResult(status="success", output={"values": values})


@file_action
def write_cell(session: WorkbookSession, sheet: str, cell: str, value: Any) -> ActionResult:
    """Write a value to a single cell.

    Args:
        session: The workbook session to write to.
        sheet: Worksheet name.
        cell: An A1-style cell reference (e.g. "B2").
        value: The value to write. A string starting with "=" is stored as a formula.

    Returns:
        A success result with no meaningful output.
    """
    backends.write_cell(session.handle, sheet, cell, value)
    session.dirty = True
    return ActionResult(status="success", output={})
