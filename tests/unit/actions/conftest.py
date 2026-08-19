"""Shared fixtures for action tests: a real openpyxl fixture workbook wrapped in a
WorkbookSession, exactly what the (not-yet-built) runner will hand to an action function."""

from collections.abc import Iterator
from pathlib import Path

import openpyxl
import pytest

from excel_runner import backends
from excel_runner.core import WorkbookSession


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
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


@pytest.fixture
def file_session(workbook_path: Path) -> Iterator[WorkbookSession]:
    handle = backends.open_workbook(str(workbook_path), mode="read_write")
    yield WorkbookSession(
        name="manip", backend="file", handle=handle, path=str(workbook_path), mode="read_write"
    )
