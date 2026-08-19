"""Unit tests for the `copy` action — the one action needing two open sessions at once.

The runner (not built yet) will need special-case wiring to resolve both `source.workbook`
and `target.workbook` into two sessions before calling this — every other action's single
`workbook:` field maps to one `session` param, but copy's YAML shape has two nested workbook
refs (`source: {...}`, `target: {...}`). Noted in Spec sec 4; not blocking building the
action function itself, which just needs two already-open sessions handed to it.
"""

from pathlib import Path

import openpyxl

from excel_runner import backends
from excel_runner.actions import copy as copy_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


def _make_target_session(tmp_path: Path) -> WorkbookSession:
    path = tmp_path / "target.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Summary"
    workbook.save(path)
    handle = backends.open_workbook(str(path), mode="read_write")
    return WorkbookSession(name="target", backend="file", handle=handle, path=str(path), mode="read_write")


class TestCopyAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["copy"] == "file"

    def test_copies_an_explicit_range(self, file_session: WorkbookSession, tmp_path: Path) -> None:
        target_session = _make_target_session(tmp_path)
        result = copy_action(
            session=file_session,
            target=target_session,
            source_sheet="Summary",
            target_sheet="Summary",
            target_range="D1",
            source_range="A1:B2",
        )
        assert result.status == "success"
        assert backends.read_range(target_session.handle, "Summary", "D1:E2") == [["Region", "Total"], ["North", 100]]

    def test_marks_the_target_session_dirty_not_the_source(
        self, file_session: WorkbookSession, tmp_path: Path
    ) -> None:
        target_session = _make_target_session(tmp_path)
        copy_action(
            session=file_session,
            target=target_session,
            source_sheet="Summary",
            target_sheet="Summary",
            target_range="D1",
        )
        assert target_session.dirty is True
        assert file_session.dirty is False
