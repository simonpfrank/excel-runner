"""Unit tests for the `stop` control-flow action (Spec sec 4, PRD sec 6.9)."""

from excel_runner.actions import stop as stop_action
from excel_runner.core import ACTION_CAPABILITIES


class TestStopAction:
    def test_registers_as_a_control_action(self) -> None:
        assert ACTION_CAPABILITIES["stop"] == "none"

    def test_returns_success_with_empty_output_by_default(self) -> None:
        result = stop_action()
        assert result.status == "success"
        assert result.output == {}

    def test_returns_the_reason_in_output_when_given(self) -> None:
        result = stop_action(reason="lookup failed")
        assert result.status == "success"
        assert result.output == {"reason": "lookup failed"}
