"""Unit tests for the xlwings (live-Excel) backend primitives in excel_runner.backends
(Spec sec 3). Real xlwings against a real, locally-spawned Excel instance — no mocks (project
convention) — same spirit as test_owned_instance_registry.py.
"""

from pathlib import Path

import openpyxl
import pytest

from excel_runner import backends
from excel_runner.backends import OwnedInstanceRegistry
from tests.unit.conftest import requires_excel, requires_working_xlwings_save


def _make_workbook(path: Path) -> Path:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    sheet["A1"] = "hello"
    workbook.save(path)
    return path


@requires_excel
class TestXlwOpenWorkbook:
    def test_opens_an_existing_workbook(self, tmp_path: Path) -> None:
        path = _make_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            assert book.name == "fixture.xlsx"
            assert book.sheets["Summary"]["A1"].value == "hello"
        finally:
            registry.close_owned()

    def test_missing_file_raises_file_not_found(self, tmp_path: Path) -> None:
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            with pytest.raises(FileNotFoundError):
                backends.xlw_open_workbook(
                    app, str(tmp_path / "does_not_exist.xlsx"), mode="read_write"
                )
        finally:
            registry.close_owned()


@requires_excel
class TestXlwCloseWorkbook:
    def test_closes_the_book_without_quitting_the_app(self, tmp_path: Path) -> None:
        path = _make_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            backends.xlw_close_workbook(book)
            # spawn() spawns with add_book=True (a Windows reliability fix, see
            # OwnedInstanceRegistry.spawn's docstring) — the app's own initial "Book1"
            # is expected to remain; the one this test opened and closed must not.
            assert "fixture.xlsx" not in [b.name for b in app.books]
        finally:
            registry.close_owned()


@requires_excel
@requires_working_xlwings_save
class TestXlwSaveWorkbook:
    def test_saves_in_place_and_the_write_persists(self, tmp_path: Path) -> None:
        path = _make_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["B1"].value = "written via xlwings"
            backends.xlw_save_workbook(book)
            backends.xlw_close_workbook(book)
        finally:
            registry.close_owned()

        reopened = openpyxl.load_workbook(path)
        assert reopened["Summary"]["B1"].value == "written via xlwings"
