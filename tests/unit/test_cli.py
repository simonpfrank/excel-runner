"""Unit tests for the CLI entrypoint (excel_runner.cli) — mocks run_workflow itself, since the
integration-level, zero-mock, real-YAML/real-workbook exercise of run_workflow already exists in
tests/integration/test_run_workflow.py. This file only tests the CLI's own argument-parsing and
exit-code responsibility. Nothing is printed to stdout (Spec sec 6.4's correction — results live
at the run's fixed working_dir/audit.jsonl path, not stdout) beyond whatever a console logging
handler is configured to show, which is the console-logging tests' job (test_runner_logging.py),
not this file's.
"""

import logging
from pathlib import Path
from unittest.mock import patch

import pytest

from excel_runner.cli import main
from excel_runner.core import ErrorDetail, ValidationError
from excel_runner.runner import RunResult, StepResult


def _success_result() -> RunResult:
    return RunResult(
        status="success",
        step_results=(
            StepResult(step_id="s1", status="success", output={"values": 1}),
        ),
        audit_log_path=Path("/tmp/run/audit.jsonl"),
    )


def _error_result() -> RunResult:
    return RunResult(
        status="error",
        step_results=(
            StepResult(
                step_id="s1",
                status="error",
                output={},
                error=ErrorDetail(message="bad", technical_reason="KeyError"),
            ),
        ),
        audit_log_path=Path("/tmp/run/audit.jsonl"),
    )


class TestMain:
    def test_success_returns_zero(self) -> None:
        with patch(
            "excel_runner.cli.run_workflow", return_value=_success_result()
        ) as mock_run:
            exit_code = main(["workflow.yaml"])

        mock_run.assert_called_once_with("workflow.yaml", None, working_dir=None)
        assert exit_code == 0

    def test_step_error_returns_one(self) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_error_result()):
            exit_code = main(["workflow.yaml"])

        assert exit_code == 1

    def test_env_overrides_are_parsed_from_key_value_pairs(self) -> None:
        with patch(
            "excel_runner.cli.run_workflow", return_value=_success_result()
        ) as mock_run:
            main(["workflow.yaml", "--env", "a=1", "--env", "b=2"])

        mock_run.assert_called_once_with("workflow.yaml", {"a": "1", "b": "2"}, working_dir=None)

    def test_working_dir_flag_is_passed_through(self) -> None:
        with patch(
            "excel_runner.cli.run_workflow", return_value=_success_result()
        ) as mock_run:
            main(["workflow.yaml", "--working-dir", "/some/base"])

        mock_run.assert_called_once_with("workflow.yaml", None, working_dir="/some/base")

    def test_logging_level_flag_sets_the_package_logger_level(self) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml", "--logging-level", "DEBUG"])

        assert logging.getLogger("excel_runner").level == logging.DEBUG

    def test_logging_level_defaults_to_info(self) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml"])

        assert logging.getLogger("excel_runner").level == logging.INFO

    def test_validation_error_returns_one_and_logs_the_message(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        error = ValidationError(
            ErrorDetail(message="bad workflow", technical_reason="oops")
        )
        with patch("excel_runner.cli.run_workflow", side_effect=error):
            with caplog.at_level(logging.ERROR, logger="excel_runner.cli"):
                exit_code = main(["workflow.yaml"])

        assert exit_code == 1
        assert any("bad workflow" in record.message for record in caplog.records)
