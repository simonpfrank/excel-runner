"""Unit tests for the `close` action (Spec sec 7 catalog)."""

from excel_runner.actions import close as close_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestCloseAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["close"] == "file"

    def test_closes_the_session_handle_and_returns_success(
        self, file_session: WorkbookSession
    ) -> None:
        result = close_action(session=file_session)
        assert result.status == "success"
