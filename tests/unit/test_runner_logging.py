"""Unit tests for console/application logging via stdlib logging (Spec sec 6.2.1, PRD sec
6.7.1) — real-time narration distinct from the audit log. Library code (`runner.py`,
`engine.py`, `backends.py`, `actions.py`) only ever emits via `logging.getLogger(...)`; it
never attaches handlers itself (standard Python library practice) — only `cli.py` does that,
via `configure_logging()` (see `test_cli.py::TestConsoleLogging`). Tests here use pytest's
`caplog` fixture, which attaches its own handler for the duration of the test, rather than
`configure_logging()`'s real `StreamHandler`s.
"""

import logging
from pathlib import Path

import openpyxl
import pytest

from excel_runner.runner import run_workflow


def _make_workbook(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet"
    sheet["A1"] = "hello"
    workbook.save(path)
    return path


def _write_yaml(path: Path, text: str) -> Path:
    path.write_text(text)
    return path


class TestStepLogging:
    def test_info_logs_step_start_and_completion(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_it
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "x"
            """,
        )

        with caplog.at_level(logging.INFO, logger="excel_runner.runner"):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        messages = " ".join(record.message for record in caplog.records)
        assert "write_it" in messages
        assert "write_cell" in messages
        assert "success" in messages

    def test_warning_level_logs_a_normal_anticipated_step_error(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        """An ActionResult(status="error") like find_row finding nothing is a normal,
        anticipated outcome the workflow can react to via if: (not a raised exception) — logged
        at WARNING, not ERROR, since ERROR is reserved for something that actually halts the
        run."""
        _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: find_it
                action: find_row
                workbook: manip
                sheet: "Sheet"
                column: "A"
                search_value: "does-not-exist"
            """,
        )

        with caplog.at_level(logging.INFO, logger="excel_runner.runner"):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        warning_records = [r for r in caplog.records if r.levelno == logging.WARNING]
        assert any("find_it" in r.message for r in warning_records)
        assert not any(r.levelno >= logging.ERROR for r in caplog.records)

    def test_debug_logs_resolved_params(
        self, tmp_path: Path, caplog: pytest.LogCaptureFixture
    ) -> None:
        _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_it
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "hello-debug-marker"
            """,
        )

        with caplog.at_level(logging.DEBUG, logger="excel_runner.runner"):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        debug_records = [r for r in caplog.records if r.levelno == logging.DEBUG]
        assert any("hello-debug-marker" in r.message for r in debug_records)
