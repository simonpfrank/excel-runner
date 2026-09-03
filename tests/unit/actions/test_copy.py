"""Unit tests for the `copy` action — the one action needing two open sessions at once.

Real xlwings against a real, locally-spawned Excel instance -- no mocks (project convention,
matches test_recalculate.py / test_owned_instance_registry.py). `copy` is a `com` capability
action (backends.com_copy_range, via Excel's own Copy so formulas/formatting come across too,
not just values) -- both sessions must already be on the `xlw` backend.

The runner (not built yet) will need special-case wiring to resolve both `source.workbook`
and `target.workbook` into two sessions before calling this — every other action's single
`workbook:` field maps to one `session` param, but copy's YAML shape has two nested workbook
refs (`source: {...}`, `target: {...}`). Noted in Spec sec 4; not blocking building the
action function itself, which just needs two already-open sessions handed to it.
"""

from collections.abc import Iterator
from pathlib import Path

import openpyxl
import pytest

from excel_runner.actions import copy as copy_action
from excel_runner.backends import (
    OwnedInstanceRegistry,
    xlw_close_workbook,
    xlw_open_workbook,
)
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession
from tests.unit.conftest import requires_excel, requires_working_xlwings_save


@pytest.fixture
def source_workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "source.xlsx"
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


@pytest.fixture
def target_workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "target.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    workbook.save(path)
    return path


@pytest.fixture
def copy_sessions(
    source_workbook_path: Path, target_workbook_path: Path
) -> Iterator[tuple[WorkbookSession, WorkbookSession]]:
    """Two live workbooks open in the same shared Excel instance — Excel's Copy/paste only
    works within one App instance, not across two separately-spawned ones."""
    registry = OwnedInstanceRegistry()
    app = registry.spawn()
    source_book = xlw_open_workbook(app, str(source_workbook_path), mode="read_write")
    target_book = xlw_open_workbook(app, str(target_workbook_path), mode="read_write")
    try:
        source_session = WorkbookSession(
            name="source",
            backend="xlw",
            handle=source_book,
            path=str(source_workbook_path),
            mode="read_write",
        )
        target_session = WorkbookSession(
            name="target",
            backend="xlw",
            handle=target_book,
            path=str(target_workbook_path),
            mode="read_write",
        )
        yield source_session, target_session
    finally:
        xlw_close_workbook(source_book)
        xlw_close_workbook(target_book)
        registry.close_owned()


class TestCopyAction:
    def test_registers_as_a_com_action(self) -> None:
        assert ACTION_CAPABILITIES["copy"] == "com"


@requires_excel
@requires_working_xlwings_save
class TestCopyActionBehavior:
    def test_copies_an_explicit_range(
        self, copy_sessions: tuple[WorkbookSession, WorkbookSession]
    ) -> None:
        source_session, target_session = copy_sessions
        result = copy_action(
            session=source_session,
            target=target_session,
            source_sheet="Summary",
            target_sheet="Summary",
            target_range="D1",
            source_range="A1:B2",
        )
        assert result.status == "success"
        target_sheet = target_session.handle.sheets["Summary"]
        assert target_sheet["D1"].value == "Region"
        assert target_sheet["E1"].value == "Total"
        assert target_sheet["D2"].value == "North"
        assert target_sheet["E2"].value == 100

    def test_marks_the_target_session_dirty_not_the_source(
        self, copy_sessions: tuple[WorkbookSession, WorkbookSession]
    ) -> None:
        source_session, target_session = copy_sessions
        copy_action(
            session=source_session,
            target=target_session,
            source_sheet="Summary",
            target_sheet="Summary",
            target_range="D1",
        )
        assert target_session.dirty is True
        assert source_session.dirty is False
