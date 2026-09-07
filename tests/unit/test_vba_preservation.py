"""Unit tests for macro (VBA) preservation across an openpyxl round trip.

openpyxl's `load_workbook` defaults to `keep_vba=False`, and on save it writes a fresh package
from the parsed model — so any part it doesn't know about, including `xl/vbaProject.bin`, is
simply not written back. For an .xlsm that is silent data loss: no exception, no warning, just
a workbook whose macros have gone. The file backend saves on commit, so the loss lands in the
user's real workbook, not a scratch copy.

These tests pin the round trip rather than the flag, so they'd still catch the regression if
the implementation moved somewhere else.
"""

import zipfile
from pathlib import Path

import pytest

from excel_runner import backends

from .conftest import VBA_PROJECT_PART, macro_enabled_workbook, plain_workbook


def _parts(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as archive:
        return set(archive.namelist())


DEFAULT_SHEET = "Sheet"


class TestMacroEnabledRoundTrip:
    def test_the_fixture_really_carries_a_vba_part(self, tmp_path: Path) -> None:
        """Guards the test itself: if the fixture stopped containing VBA, every other test
        here would pass for the wrong reason."""
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")

        assert VBA_PROJECT_PART in _parts(path)

    def test_macros_survive_an_edit_and_save(self, tmp_path: Path) -> None:
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")

        workbook = backends.open_workbook(str(path), mode="read_write")
        backends.write_cell(workbook, DEFAULT_SHEET, "A1", "edited")
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        assert VBA_PROJECT_PART in _parts(path)

    def test_the_edit_itself_still_lands(self, tmp_path: Path) -> None:
        """Preserving VBA must not come at the cost of the write being dropped."""
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")

        workbook = backends.open_workbook(str(path), mode="read_write")
        backends.write_cell(workbook, DEFAULT_SHEET, "A1", "edited")
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        reopened = backends.open_workbook(str(path), mode="read_only")
        assert backends.read_range(reopened, DEFAULT_SHEET, "A1") == "edited"
        backends.close_workbook(reopened)

    def test_a_macro_free_workbook_gains_no_vba_part(self, tmp_path: Path) -> None:
        """The flag must not invent a VBA part where the original had none."""
        path = plain_workbook(tmp_path / "plain.xlsx")

        workbook = backends.open_workbook(str(path), mode="read_write")
        backends.write_cell(workbook, DEFAULT_SHEET, "A1", "edited")
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        assert VBA_PROJECT_PART not in _parts(path)

    def test_an_xlsx_is_not_relabelled_as_macro_enabled(self, tmp_path: Path) -> None:
        """openpyxl derives the saved workbook part's content type from whether a cached VBA
        archive is attached (`Workbook.mime_type`), not from what the source file actually
        contained. Turning `keep_vba` on indiscriminately therefore stamps every .xlsx as
        macroEnabled, and Excel then refuses it with "the file format and extension don't
        match" — trading one silent corruption for another."""
        path = plain_workbook(tmp_path / "plain.xlsx")

        workbook = backends.open_workbook(str(path), mode="read_write")
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        with zipfile.ZipFile(path) as archive:
            content_types = archive.read("[Content_Types].xml").decode()
        assert "macroEnabled" not in content_types

    def test_an_xlsm_is_still_labelled_macro_enabled(self, tmp_path: Path) -> None:
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")

        workbook = backends.open_workbook(str(path), mode="read_write")
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        with zipfile.ZipFile(path) as archive:
            content_types = archive.read("[Content_Types].xml").decode()
        assert "macroEnabled" in content_types

    @pytest.mark.parametrize("data_only", [True, False])
    def test_macros_survive_regardless_of_the_data_only_view(
        self, tmp_path: Path, data_only: bool
    ) -> None:
        """`data_only` picks formula text vs cached value and is orthogonal to VBA — a session
        reading formulas must not lose macros as a side effect."""
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")

        workbook = backends.open_workbook(
            str(path), mode="read_write", data_only=data_only
        )
        backends.save_open_workbook(workbook, "file", str(path))
        backends.close_workbook(workbook)

        assert VBA_PROJECT_PART in _parts(path)


class TestClosingReleasesTheCachedArchive:
    """`keep_vba=True` attaches an in-memory ZipFile that openpyxl's own `close()` ignores.
    Left to the garbage collector it raises an unraisable `ValueError` from `ZipFile.__del__`
    at an arbitrary later point, which is how this was noticed."""

    def test_close_releases_the_vba_archive(self, tmp_path: Path) -> None:
        path = macro_enabled_workbook(tmp_path / "macros.xlsm")
        workbook = backends.open_workbook(str(path), mode="read_write")
        archive = workbook.vba_archive
        assert archive is not None

        backends.close_workbook(workbook)

        assert archive.fp is None

    def test_close_is_safe_when_there_is_no_cached_archive(self, tmp_path: Path) -> None:
        path = plain_workbook(tmp_path / "plain.xlsx")
        workbook = backends.open_workbook(str(path), mode="read_only")

        backends.close_workbook(workbook)

        assert workbook.vba_archive is None
