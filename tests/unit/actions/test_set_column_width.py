"""Unit tests for the `set_column_width` action (PRD sec 7)."""

from excel_runner.actions import set_column_width as set_column_width_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestSetColumnWidthAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["set_column_width"] == "file"

    def test_sets_an_explicit_width(self, file_session: WorkbookSession) -> None:
        result = set_column_width_action(session=file_session, sheet="Summary", columns="A:A", width=20)
        assert result.status == "success"
        assert file_session.handle["Summary"].column_dimensions["A"].width == 20

    def test_autofit(self, file_session: WorkbookSession) -> None:
        result = set_column_width_action(session=file_session, sheet="Summary", columns="A:A", width="autofit")
        assert result.status == "success"
        assert file_session.handle["Summary"].column_dimensions["A"].width > 0
