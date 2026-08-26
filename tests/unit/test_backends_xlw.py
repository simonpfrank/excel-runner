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


def _make_formula_workbook(path: Path) -> Path:
    """A workbook whose B1 depends on A1, with the cached formula result deliberately stale
    (openpyxl can't compute formulas itself, so it never gets to compute a correct one) — a
    real recalculation is what's needed to make B1 match A1*2."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    sheet["A1"] = 10
    sheet["B1"] = "=A1*2"
    workbook.save(path)
    return path


@requires_excel
@requires_working_xlwings_save
class TestRecalculatePrimitives:
    def test_com_calculate_workbook_updates_a_formula(self, tmp_path: Path) -> None:
        path = _make_formula_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["A1"].value = 21
            backends.com_calculate_workbook(book)
            backends.com_wait_until_calculation_done(app)
            assert book.sheets["Summary"]["B1"].value == 42
        finally:
            registry.close_owned()

    def test_com_calculate_sheet_updates_a_formula(self, tmp_path: Path) -> None:
        path = _make_formula_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["A1"].value = 5
            backends.com_calculate_sheet(book, "Summary")
            backends.com_wait_until_calculation_done(app)
            assert book.sheets["Summary"]["B1"].value == 10
        finally:
            registry.close_owned()

    def test_xlw_calculate_all_updates_a_formula(self, tmp_path: Path) -> None:
        path = _make_formula_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["A1"].value = 3
            backends.xlw_calculate_all(app)
            backends.com_wait_until_calculation_done(app)
            assert book.sheets["Summary"]["B1"].value == 6
        finally:
            registry.close_owned()

    def test_com_calculate_full_updates_a_formula(self, tmp_path: Path) -> None:
        path = _make_formula_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["A1"].value = 7
            backends.com_calculate_full(app)
            backends.com_wait_until_calculation_done(app)
            assert book.sheets["Summary"]["B1"].value == 14
        finally:
            registry.close_owned()

    def test_com_calculate_full_rebuild_updates_a_formula(self, tmp_path: Path) -> None:
        path = _make_formula_workbook(tmp_path / "fixture.xlsx")
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = backends.xlw_open_workbook(app, str(path), mode="read_write")
            book.sheets["Summary"]["A1"].value = 9
            backends.com_calculate_full_rebuild(app)
            backends.com_wait_until_calculation_done(app)
            assert book.sheets["Summary"]["B1"].value == 18
        finally:
            registry.close_owned()

    def test_wait_raises_timeout_error_if_calculation_never_finishes(self, tmp_path: Path) -> None:
        """A fake app whose CalculationState never reports done — proves the timeout path
        is real, without needing an actual multi-minute Excel calculation."""

        class _AlwaysCalculatingApi:
            CalculationState = 1  # xlCalculating, never changes

        class _FakeApp:
            api = _AlwaysCalculatingApi()

        with pytest.raises(TimeoutError):
            backends.com_wait_until_calculation_done(_FakeApp(), timeout=0.05)  # type: ignore[arg-type]

