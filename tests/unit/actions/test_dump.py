"""Unit tests for the `dump` control-flow action — internal step-output introspection for
authoring/debugging (see excel_runner/actions.py's dump docstring)."""

import json
from pathlib import Path

import pytest

from excel_runner.actions import dump as dump_action
from excel_runner.core import ACTION_CAPABILITIES, ActionExecutionError

_STEP_OUTPUTS = {
    "a": {"status": "success", "output": {"values": 1}},
    "b": {"status": "success", "output": {"values": 2}},
}


class TestDumpAction:
    def test_registers_as_a_control_action(self) -> None:
        assert ACTION_CAPABILITIES["dump"] == "none"

    def test_returns_success_with_empty_output(self) -> None:
        result = dump_action(step_outputs=_STEP_OUTPUTS)
        assert result.status == "success"
        assert result.output == {}

    def test_prints_every_step_by_default(self, capsys: pytest.CaptureFixture[str]) -> None:
        dump_action(step_outputs=_STEP_OUTPUTS)
        printed = capsys.readouterr().out
        assert '"a"' in printed
        assert '"b"' in printed

    def test_filters_to_requested_ids(self, capsys: pytest.CaptureFixture[str]) -> None:
        dump_action(step_outputs=_STEP_OUTPUTS, ids=["a"])
        printed = capsys.readouterr().out
        assert '"a"' in printed
        assert '"b"' not in printed

    def test_unknown_id_is_skipped_with_a_warning(
        self, caplog: pytest.LogCaptureFixture, capsys: pytest.CaptureFixture[str]
    ) -> None:
        with caplog.at_level("WARNING"):
            result = dump_action(step_outputs=_STEP_OUTPUTS, ids=["a", "missing"])
        assert result.status == "success"
        assert "missing" in caplog.text
        printed = capsys.readouterr().out
        assert '"a"' in printed

    def test_writes_to_file_when_requested(self, tmp_path: Path) -> None:
        target = tmp_path / "nested" / "dump.json"
        dump_action(step_outputs=_STEP_OUTPUTS, to="file", path=str(target))
        assert json.loads(target.read_text()) == _STEP_OUTPUTS

    def test_file_without_path_raises(self) -> None:
        with pytest.raises(ActionExecutionError):
            dump_action(step_outputs=_STEP_OUTPUTS, to="file")
