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

from typing import Any, Literal

from openpyxl.utils import column_index_from_string, get_column_letter

from excel_runner import backends
from excel_runner.core import (
    ActionExecutionError,
    ActionResult,
    ErrorDetail,
    WorkbookSession,
    control_action,
    file_action,
)

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


@file_action(writes=True)
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


@control_action
def stop(reason: str | None = None) -> ActionResult:
    """Halt the run — no later step runs (PRD sec 6.9).

    No `session` parameter and no `workbook:` field — this is pure control flow, not a backend
    call. Reaching this action is not itself a failure; `runner.py` is what actually ends the
    loop once this returns, marking every later step `"stopped"` (Spec sec 6.1).

    Args:
        reason: Optional note for the audit log explaining why the run stopped.

    Returns:
        A success result; `reason` is echoed into `output` when given, for the audit log.
    """
    return ActionResult(
        status="success", output={"reason": reason} if reason is not None else {}
    )


# --- data ----------------------------------------------------------------------------------


@file_action(writes=True)
def copy(
    session: WorkbookSession,
    target: WorkbookSession,
    source_sheet: str,
    target_sheet: str,
    target_range: str,
    source_range: str | None = None,
) -> ActionResult:
    """Copy a range — or, if `source_range` is omitted, the whole sheet — into another session.

    The one action needing two open sessions at once. The (not-yet-built) runner will need
    special-case wiring to resolve both `source.workbook` and `target.workbook` into `session`
    and `target` before calling this — every other action's single `workbook:` field maps to
    one `session` param, but copy's YAML shape has two nested workbook refs (PRD sec 7/sec 11).

    Args:
        session: The source workbook session.
        target: The target workbook session.
        source_sheet: Source worksheet name.
        target_sheet: Target worksheet name.
        target_range: Where to start writing — only the top-left cell is used.
        source_range: An A1-style range, or None to copy the whole sheet.

    Returns:
        A success result with no meaningful output.
    """
    backends.copy_range(
        session.handle,
        source_sheet,
        source_range,
        target.handle,
        target_sheet,
        target_range,
    )
    target.dirty = True
    return ActionResult(status="success", output={})


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
def read_metadata(
    session: WorkbookSession,
    target: Literal["properties", "cells"],
    sheet: str | None = None,
    cells: list[str] | None = None,
) -> ActionResult:
    """Read document properties, or a scattered list of specific cells.

    The `textboxes` sub-target from PRD sec 7 is COM-only (openpyxl can't see live control
    state) and not built here — deferred to the COM phase (Spec sec 8). `sheet` is a
    clarification found during implementation: PRD sec 7's catalog didn't list it for the
    `cells` sub-case, but reading specific cells needs to know which worksheet they're on,
    same as every other cell-addressing action.

    Args:
        session: The workbook session to read from.
        target: "properties" for document properties, "cells" for a scattered cell list.
        sheet: Worksheet name — required if target is "cells".
        cells: A1-style cell references to read — required if target is "cells".

    Returns:
        `{"values": ...}`-style keyed output: document properties by name, or cell reference
        to value, depending on `target`.

    Raises:
        ActionExecutionError: If target is "cells" but `sheet`/`cells` weren't given.
    """
    if target == "properties":
        return ActionResult(
            status="success", output=backends.read_properties(session.handle)
        )
    if target != "cells":
        # Python doesn't enforce type hints at runtime, and this function is directly
        # importable/callable on its own (PRD sec 3/sec 9's library goal) — so an unsupported
        # target must be rejected explicitly here, not silently fall through to the "cells"
        # handling below just because it wasn't "properties". Found while reasoning about
        # target="textboxes" specifically: it used to be mishandled exactly this way.
        raise ActionExecutionError(
            ErrorDetail(
                message=f'read_metadata: target "{target}" is not supported yet.',
                technical_reason=f"read_metadata called with unsupported target {target!r}",
            )
        )
    if sheet is None or cells is None:
        raise ActionExecutionError(
            ErrorDetail(
                message='read_metadata: target "cells" requires both `sheet` and `cells`.',
                technical_reason="read_metadata called with target=cells but sheet or cells was None",
            )
        )
    return ActionResult(
        status="success", output=backends.read_cells(session.handle, sheet, cells)
    )


@file_action(writes=True)
def write_cell(
    session: WorkbookSession, sheet: str, cell: str, value: Any
) -> ActionResult:
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


@file_action(writes=True)
def write_range(
    session: WorkbookSession, sheet: str, range: str, values: list[list[Any]]
) -> ActionResult:
    """Write a 2D block of values, anchored at the top-left cell of `range`.

    Args:
        session: The workbook session to write to.
        sheet: Worksheet name.
        range: An A1-style cell or range — only the top-left cell is used as the anchor.
        values: A 2D list of row values to write.

    Returns:
        A success result with no meaningful output.
    """
    backends.write_range(session.handle, sheet, range, values)
    session.dirty = True
    return ActionResult(status="success", output={})


@file_action(writes=True)
def write_row(
    session: WorkbookSession,
    sheet: str,
    row: int,
    values: dict[str, Any] | list[Any],
    start_column: str | None = None,
) -> ActionResult:
    """Write a row of values, either by explicit column mapping or positionally.

    The by-header mode (`values_by_header` + `headers_from`, PRD sec 7/sec 11 item 9) isn't
    built here — it needs another step's output, which doesn't exist as a concept until
    runner.py threads step-output context through (Spec sec 4/8).

    Args:
        session: The workbook session to write to.
        sheet: Worksheet name.
        row: The row number to write into.
        values: Either `{column: value}` (explicit mapping), or a plain ordered list written
            left-to-right starting at `start_column` (positional mode).
        start_column: Required when `values` is a list — the column to start writing at.

    Returns:
        A success result with no meaningful output.

    Raises:
        ActionExecutionError: If `values` is a list but `start_column` wasn't given.
    """
    if isinstance(values, dict):
        for column, value in values.items():
            backends.write_cell(session.handle, sheet, f"{column}{row}", value)
    else:
        if start_column is None:
            raise ActionExecutionError(
                ErrorDetail(
                    message="write_row: positional `values` (a list) needs `start_column`.",
                    technical_reason="write_row called with list values but start_column=None",
                )
            )
        start_idx = column_index_from_string(start_column)
        for offset, value in enumerate(values):
            column = get_column_letter(start_idx + offset)
            backends.write_cell(session.handle, sheet, f"{column}{row}", value)
    session.dirty = True
    return ActionResult(status="success", output={})


# --- structure -----------------------------------------------------------------------------


@file_action(writes=True)
def insert_range(
    session: WorkbookSession,
    sheet: str,
    at: str,
    direction: Literal["rows", "columns"] | None = None,
    header: dict[str, Any] | None = None,
) -> ActionResult:
    """Insert a whole row or whole column, shifting existing content.

    A partial range (e.g. "C5:C10") isn't built yet (PRD sec 11 item 12's flagged cost) —
    caught here and returned as a structured error, consistent with find_*'s "legitimately
    didn't work" pattern below, rather than a raw exception escaping.

    Args:
        session: The workbook session to modify.
        sheet: Worksheet name.
        at: A whole-column reference (e.g. "C:C") or whole-row reference (e.g. "5:5").
        direction: Unused for whole-row/whole-column inserts (unambiguous from `at` itself).
        header: `{"row": int, "text": str}` — only meaningful for a column insert.

    Returns:
        A success result, or a structured error if `at` is a partial range.
    """
    try:
        backends.insert_range(session.handle, sheet, at, direction, header)
    except NotImplementedError as exc:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=str(exc), technical_reason=f"{type(exc).__name__}: {exc}"
            ),
        )
    session.dirty = True
    return ActionResult(status="success", output={})


@file_action(writes=True)
def set_column_width(
    session: WorkbookSession,
    sheet: str,
    columns: str,
    width: float | Literal["autofit"],
) -> ActionResult:
    """Set the width of a column or range of columns.

    Args:
        session: The workbook session to modify.
        sheet: Worksheet name.
        columns: A single column letter (e.g. "B") or range (e.g. "A:C").
        width: An explicit width, or "autofit".

    Returns:
        A success result with no meaningful output.
    """
    backends.set_column_width(session.handle, sheet, columns, width)
    session.dirty = True
    return ActionResult(status="success", output={})


@file_action(writes=True)
def create_sheet(
    session: WorkbookSession, name: str, index: int | None = None
) -> ActionResult:
    """Add a new, empty worksheet to the workbook.

    Args:
        session: The workbook session to modify.
        name: Name for the new sheet.
        index: Position to insert at (0-based). Appended at the end if omitted.

    Returns:
        A success result, or a structured error if a sheet named `name` already exists —
        a normal "didn't work" outcome (a workflow author picked a name already in use), not
        an unexpected failure, consistent with the find_*/insert_range error pattern above.
    """
    try:
        backends.create_sheet(session.handle, name, index)
    except ValueError as exc:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=str(exc), technical_reason=f"{type(exc).__name__}: {exc}"
            ),
        )
    session.dirty = True
    return ActionResult(status="success", output={})


@file_action(writes=True)
def rename_sheet(session: WorkbookSession, sheet: str, new_name: str) -> ActionResult:
    """Rename an existing worksheet.

    Args:
        session: The workbook session to modify.
        sheet: Current worksheet name.
        new_name: New name for the worksheet.

    Returns:
        A success result with no meaningful output.
    """
    backends.rename_sheet(session.handle, sheet, new_name)
    session.dirty = True
    return ActionResult(status="success", output={})


@file_action(writes=True)
def delete_sheet(session: WorkbookSession, sheet: str) -> ActionResult:
    """Remove a worksheet from the workbook.

    Args:
        session: The workbook session to modify.
        sheet: Name of the worksheet to remove.

    Returns:
        A success result, or a structured error if `sheet` is the workbook's only remaining
        sheet — a workbook can't have zero sheets, and this is a normal, avoidable outcome for
        a workflow author to react to, not an unexpected failure.
    """
    try:
        backends.delete_sheet(session.handle, sheet)
    except ValueError as exc:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=str(exc), technical_reason=f"{type(exc).__name__}: {exc}"
            ),
        )
    session.dirty = True
    return ActionResult(status="success", output={})


# --- lookup ----------------------------------------------------------------------------------


@file_action
def find_headers_row(
    session: WorkbookSession, sheet: str, search_range: str, patterns: list[str]
) -> ActionResult:
    """Find the row within `search_range` where every pattern matches some cell in that row.

    Args:
        session: The workbook session to search.
        sheet: Worksheet name.
        search_range: An A1-style range to search within.
        patterns: Regex patterns — every one must match a cell in a row for that row to count.

    Returns:
        `{"row": int, "headers": {pattern: column_letter}}`, or a structured error if no row
        matches every pattern — a normal outcome of a search, not an unexpected failure.
    """
    result = backends.find_headers_row(session.handle, sheet, search_range, patterns)
    if result is None:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=f'No row in "{search_range}" of sheet "{sheet}" matches every pattern: {patterns}.',
                technical_reason="find_headers_row: no row matched all patterns",
            ),
        )
    row, headers = result
    return ActionResult(status="success", output={"row": row, "headers": headers})


@file_action
def find_row(
    session: WorkbookSession,
    sheet: str,
    column: str,
    search_value: Any,
    header_row: int | None = None,
) -> ActionResult:
    """Find the row number where `column` equals `search_value`.

    Args:
        session: The workbook session to search.
        sheet: Worksheet name.
        column: A column letter (e.g. "B").
        search_value: The value to match, by equality.
        header_row: If given, search starts on the row after it.

    Returns:
        `{"row": int}`, or a structured error if not found.
    """
    row = backends.find_row(session.handle, sheet, column, search_value, header_row)
    if row is None:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=f'No row found in column "{column}" of sheet "{sheet}" matching "{search_value}".',
                technical_reason="find_row: no matching row",
            ),
        )
    return ActionResult(status="success", output={"row": row})


@file_action
def find_column(
    session: WorkbookSession, sheet: str, header_row: int, pattern: str
) -> ActionResult:
    """Find the column letter whose header (in `header_row`) matches `pattern`.

    Args:
        session: The workbook session to search.
        sheet: Worksheet name.
        header_row: The row number containing headers.
        pattern: A regex pattern to match against each header cell's value.

    Returns:
        `{"column": str}`, or a structured error if not found.
    """
    column = backends.find_column(session.handle, sheet, header_row, pattern)
    if column is None:
        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(
                message=f'No column found in sheet "{sheet}" row {header_row} matching "{pattern}".',
                technical_reason="find_column: no matching column",
            ),
        )
    return ActionResult(status="success", output={"column": column})


@file_action
def find_columns(
    session: WorkbookSession, sheet: str, header_row: int, patterns: dict[str, str]
) -> ActionResult:
    """Find multiple named columns by header pattern in one call.

    Args:
        session: The workbook session to search.
        sheet: Worksheet name.
        header_row: The row number containing headers.
        patterns: Logical name to regex pattern.

    Returns:
        Logical name to column letter, for every pattern that matched. Names whose pattern
        didn't match anything are simply absent — not an error at this level (PRD sec 10.4).
    """
    result = backends.find_columns(session.handle, sheet, header_row, patterns)
    return ActionResult(status="success", output=result)
