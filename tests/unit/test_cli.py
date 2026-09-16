"""Unit tests for the CLI entrypoint (excel_runner.cli) — mocks run_workflow itself, since the
integration-level, zero-mock, real-YAML/real-workbook exercise of run_workflow already exists in
tests/integration/test_run_workflow.py. This file only tests the CLI's own argument-parsing and
exit-code responsibility. Nothing is printed to stdout (Spec sec 6.4's correction — results live
at the run's fixed working_dir/audit.jsonl path, not stdout) beyond whatever a console logging
handler is configured to show, which is the console-logging tests' job (test_runner_logging.py),
not this file's.
"""

import logging
import shutil
import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pytest

from excel_runner.cli import main
from excel_runner.core import ErrorDetail, ValidationError
from excel_runner.runner import RunResult, StepResult


def test_cli_launcher_can_run_from_a_deployment_folder(tmp_path: Path) -> None:
    deployment_path = tmp_path / "deployment"
    package_source = Path(__file__).parents[2] / "excel_runner"
    shutil.copytree(package_source, deployment_path / "excel_runner")
    launcher_source = Path(__file__).parents[2] / "run_excel_runner.py"
    launcher_path = deployment_path / "run_excel_runner.py"
    shutil.copy(launcher_source, launcher_path)
    workflow_source = Path(__file__).parents[2] / "unify_smoke_test.yaml"
    workflow_path = deployment_path / "unify_smoke_test.yaml"
    shutil.copy(workflow_source, workflow_path)
    working_path = tmp_path / "working"
    working_path.mkdir()
    copied_launcher_path = working_path / "run_excel_runner.py"
    shutil.copy(launcher_path, copied_launcher_path)

    result = subprocess.run(
        [
            sys.executable,
            str(copied_launcher_path),
            "--runner-home",
            str(deployment_path),
            str(workflow_path),
        ],
        check=False,
        capture_output=True,
        text=True,
        cwd=working_path,
    )

    assert result.returncode == 0, result.stderr
    workbook = openpyxl.load_workbook(working_path / "hello_world.xlsx", data_only=False)
    assert workbook["Sheet"]["A1"].value == "Hello World"

    import_result = subprocess.run(
        [
            sys.executable,
            "-c",
            "import excel_runner.cli as cli; print(cli.run_workflow.__module__)",
        ],
        check=False,
        capture_output=True,
        text=True,
        cwd=working_path,
    )

    assert import_result.returncode == 0, import_result.stderr
    assert import_result.stdout.strip() == "excel_runner.runner"


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

        mock_run.assert_called_once_with(
            "workflow.yaml", None, working_dir=None, check_existence=False
        )
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

        mock_run.assert_called_once_with(
            "workflow.yaml",
            {"a": "1", "b": "2"},
            working_dir=None,
            check_existence=False,
        )

    def test_working_dir_flag_is_passed_through(self) -> None:
        with patch(
            "excel_runner.cli.run_workflow", return_value=_success_result()
        ) as mock_run:
            main(["workflow.yaml", "--working-dir", "/some/base"])

        mock_run.assert_called_once_with(
            "workflow.yaml", None, working_dir="/some/base", check_existence=False
        )

    def test_check_existence_flag_is_passed_through(self) -> None:
        with patch(
            "excel_runner.cli.run_workflow", return_value=_success_result()
        ) as mock_run:
            main(["workflow.yaml", "--check-existence"])

        mock_run.assert_called_once_with(
            "workflow.yaml", None, working_dir=None, check_existence=True
        )

    def test_dry_run_preflights_without_running_the_workflow(self) -> None:
        with (
            patch("excel_runner.cli.preflight_workflow") as mock_preflight,
            patch("excel_runner.cli.run_workflow") as mock_run,
        ):
            exit_code = main(["workflow.yaml", "--dry-run", "--env", "period=202606"])

        assert exit_code == 0
        mock_preflight.assert_called_once_with("workflow.yaml", {"period": "202606"})
        mock_run.assert_not_called()

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


class TestConsoleLogging:
    """The CLI is the one place in this project that attaches logging handlers itself (every
    library module just calls `logging.getLogger(__name__)`, see AGENTS.md's logging section).
    WARNING/ERROR go to stderr; DEBUG/INFO go to stdout — so a caller piping only stdout
    doesn't see anything that actually needs attention, and vice versa.
    """

    def test_debug_and_info_go_to_stdout_not_stderr(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml", "--logging-level", "DEBUG"])
        logging.getLogger("excel_runner.cli").info("hello-info-marker")
        logging.getLogger("excel_runner.cli").debug("hello-debug-marker")

        captured = capsys.readouterr()
        assert "hello-info-marker" in captured.out
        assert "hello-debug-marker" in captured.out
        assert "hello-info-marker" not in captured.err
        assert "hello-debug-marker" not in captured.err

    def test_warning_and_error_go_to_stderr_not_stdout(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml"])
        logging.getLogger("excel_runner.cli").warning("hello-warning-marker")
        logging.getLogger("excel_runner.cli").error("hello-error-marker")

        captured = capsys.readouterr()
        assert "hello-warning-marker" in captured.err
        assert "hello-error-marker" in captured.err
        assert "hello-warning-marker" not in captured.out
        assert "hello-error-marker" not in captured.out

    def test_default_info_level_does_not_show_debug(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml"])
        logging.getLogger("excel_runner.cli").debug("should-not-appear")
        logging.getLogger("excel_runner.cli").info("should-appear")

        captured = capsys.readouterr()
        assert "should-not-appear" not in captured.out
        assert "should-appear" in captured.out

    def test_format_includes_level_module_function_and_line(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml"])
        logging.getLogger("excel_runner.cli").warning("format-marker")

        captured = capsys.readouterr()
        line = next(
            entry for entry in captured.err.splitlines() if "format-marker" in entry
        )
        assert "WARNING" in line
        assert "cli" in line
        assert "test_format_includes_level_module_function_and_line" in line


class TestFileLogging:
    def test_default_logfile_records_cli_logs(self, tmp_path: Path) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml", "--working-dir", str(tmp_path)])
        logging.getLogger("excel_runner.cli").info("file-log-marker")

        log_file = tmp_path / "excel_runner_runs" / "workflow" / "run.log"
        assert "file-log-marker" in log_file.read_text(encoding="utf-8")

    def test_no_logfile_does_not_create_a_run_log(self, tmp_path: Path) -> None:
        with patch("excel_runner.cli.run_workflow", return_value=_success_result()):
            main(["workflow.yaml", "--working-dir", str(tmp_path), "--no-logfile"])
        logging.getLogger("excel_runner.cli").info("not-in-a-file-marker")

        log_file = tmp_path / "excel_runner_runs" / "workflow" / "run.log"
        assert not log_file.exists()
