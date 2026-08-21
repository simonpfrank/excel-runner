"""Unit tests for console/application logging via stdlib logging (Spec sec 6.2.1, PRD sec
6.7.1) — real-time narration distinct from the audit log. Library code only ever emits via
`logging.getLogger(...)`; it never attaches handlers itself (standard Python library practice,
also decided for the CLI — Spec sec 6.2.1's corrected scope). Tests use pytest's `caplog`
fixture, which attaches its own handler for the duration of the test, rather than this project
configuring one.
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

    def test_error_level_logs_a_failed_step(
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

        error_records = [r for r in caplog.records if r.levelno >= logging.ERROR]
        assert any("find_it" in r.message for r in error_records)

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
