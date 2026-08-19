"""Unit tests for the `write_range` action (PRD sec 7/sec 11 item 8)."""

from excel_runner import backends
from excel_runner.actions import write_range as write_range_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestWriteRangeAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["write_range"] == "file"

    def test_writes_a_2d_block(self, file_session: WorkbookSession) -> None:
        result = write_range_action(session=file_session, sheet="Summary", range="D1:E2", values=[[1, 2], [3, 4]])
        assert result.status == "success"
        assert backends.read_range(file_session.handle, "Summary", "D1:E2") == [[1, 2], [3, 4]]

    def test_marks_the_session_dirty(self, file_session: WorkbookSession) -> None:
        assert file_session.dirty is False
        write_range_action(session=file_session, sheet="Summary", range="D1", values=[["x"]])
        assert file_session.dirty is True
