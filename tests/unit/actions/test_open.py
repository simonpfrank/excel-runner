"""Unit tests for the `open` action (Spec sec 7 catalog).

`open`'s own function body is deliberately minimal: by the time any action runs, the runner
has already resolved and opened the WorkbookSession it's handed (Spec sec 6.1's orchestration
loop resolves the session before dispatching to the action) — so `open` mainly confirms
success rather than doing the opening itself. update_links/mode-override params are left out
of this signature for now: update_links has no meaningful effect without COM (later phase),
and mode-override depends on session inference that doesn't exist until Spec sec 5.2/5.4 land.
"""

from excel_runner.actions import open as open_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestOpenAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["open"] == "file"

    def test_returns_success(self, file_session: WorkbookSession) -> None:
        result = open_action(session=file_session)
        assert result.status == "success"

    def test_has_no_meaningful_output(self, file_session: WorkbookSession) -> None:
        result = open_action(session=file_session)
        assert result.output == {}
