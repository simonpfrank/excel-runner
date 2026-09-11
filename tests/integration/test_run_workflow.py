"""Integration tests for run_workflow() — the full stack, zero mocks (project convention):
real YAML files, real openpyxl workbooks, no fakes anywhere. This is the first point genuine
end-to-end tests are possible (Spec sec 6.1/6.2, build order item 7) and the shape the user and
Claude agreed integration tests for this project should mostly take: a real workflow.yaml run
through run_workflow(), asserted against the resulting real workbook state.
"""

import json
from pathlib import Path

import openpyxl
import pytest

from excel_runner.core import ActionExecutionError, ValidationError
from excel_runner.runner import preflight_workflow, run_workflow


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


class TestHappyPath:
    def test_reads_writes_and_implicitly_saves_on_success(self, tmp_path: Path) -> None:
        """No explicit save/close step — commit_all() at the end of a successful run must
        persist the change to the real file on its own (PRD sec 6.3's implicit save)."""
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
              - id: write_b1
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "{{ steps.get_a1.output.values }}"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        reopened = openpyxl.load_workbook(real)
        assert reopened["Sheet"]["B1"].value == "hello"

    def test_explicit_save_and_close_steps_still_work(self, tmp_path: Path) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_b1
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "explicit"
              - id: save_it
                action: save
                workbook: manip
              - id: close_it
                action: close
                workbook: manip
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        assert openpyxl.load_workbook(real)["Sheet"]["B1"].value == "explicit"


class TestPreflight:
    def test_preflight_validates_a_write_workflow_without_changing_the_workbook(
        self, tmp_path: Path
    ) -> None:
        workbook_path = _make_workbook(tmp_path / "input.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            f"""
            workbooks:
              input:
                file: "{workbook_path.as_posix()}"
            steps:
              - id: write_value
                action: write_cell
                workbook: input
                sheet: "Sheet"
                cell: "A1"
                value: "changed"
            """,
        )

        preflight_workflow(workflow_path)

        assert openpyxl.load_workbook(workbook_path)["Sheet"]["A1"].value == "hello"
        assert not (tmp_path / "excel_runner_runs").exists()

    @pytest.mark.parametrize(
        ("action", "params"),
        [
            (
                "read_text_file",
                'file: "{input_path}"',
            ),
            (
                "replace_text",
                (
                  'workbook: input\n                sheet: "Sheet"\n'
                  '                pattern: "["\n                replacement: "x"'
                ),
            ),
        ],
    )
    def test_preflight_rejects_invalid_literal_input(
        self, tmp_path: Path, action: str, params: str
    ) -> None:
        workbook_path = _make_workbook(tmp_path / "input.xlsx")
        input_path = tmp_path / "bad.csv"
        input_path.write_text('header\n"unclosed\n')
        rendered_params = params.format(input_path=input_path.as_posix())
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            f"""
            workbooks:
              input:
                file: "{workbook_path.as_posix()}"
            steps:
              - id: validate_input
                action: {action}
                {rendered_params}
            """,
        )

        with pytest.raises(ActionExecutionError):
            preflight_workflow(workflow_path)

    def test_preflight_rejects_missing_table_column(self, tmp_path: Path) -> None:
        workbook_path = _make_workbook(tmp_path / "input.xlsx")
        sheet = openpyxl.load_workbook(workbook_path)
        sheet["Sheet"]["B7"] = "ITEM"
        sheet["Sheet"]["C7"] = "VALUE"
        sheet["Sheet"]["B8"] = "first"
        sheet.save(workbook_path)
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            f"""
            workbooks:
              input:
                file: "{workbook_path.as_posix()}"
            steps:
              - id: update_table
                action: update_table_cells
                workbook: input
                sheet: "Sheet"
                header_cell: "B7"
                lookup_column: "ITEM"
                lookup_rows: ["first"]
                target_columns: ["MISSING"]
                value: "updated"
            """,
        )

        with pytest.raises(ValidationError, match="MISSING"):
            preflight_workflow(workflow_path)


class TestAdditionalFunctions:
    def test_text_table_and_replacement_actions_run_against_a_real_workbook(
        self, tmp_path: Path
    ) -> None:
        workbook_path = tmp_path / "table.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "BASIS"
        sheet["B7"] = "BASIS_ITEM"
        sheet["C7"] = "SOURCE"
        sheet["D7"] = "TARGET"
        sheet["B8"] = "ACC_RATIO"
        sheet["C8"] = "value_202512"
        sheet["B9"] = "ECO_TBL"
        sheet["C9"] = "other"
        other = workbook.create_sheet("Other")
        other["A1"] = "value_202512"
        workbook.defined_names["NamedTarget"] = openpyxl.workbook.defined_name.DefinedName(
            "NamedTarget", attr_text="BASIS!$E$8"
        )
        workbook.save(workbook_path)
        fac_path = tmp_path / "input.fac"
        fac_path.write_text("CODE,VALUE\n00123,from_fac\n")
        workflow_path = _write_yaml(
            tmp_path / "additional.yaml",
            f"""
            workbooks:
              table:
                file: "{workbook_path.as_posix()}"
            steps:
              - id: read_fac
                action: read_text_file
                file: "{fac_path.as_posix()}"
              - id: write_fac
                action: write_range
                workbook: table
                sheet: "BASIS"
                range: "G7"
                values: "{{{{ steps.read_fac.output.values }}}}"
              - id: write_name
                action: write_cell
                workbook: table
                sheet: "BASIS"
                cell: "NamedTarget"
                value: "named"
              - id: read_basis
                action: read_table
                workbook: table
                sheet: "BASIS"
                header_cell: "B7"
              - id: copy_column
                action: copy_table_columns
                workbook: table
                sheet: "BASIS"
                header_cell: "B7"
                source_columns: ["SOURCE"]
                target_columns: ["TARGET"]
              - id: update_cell
                action: update_table_cells
                workbook: table
                sheet: "BASIS"
                header_cell: "B7"
                lookup_column: "BASIS_ITEM"
                lookup_rows: ["ECO_TBL"]
                target_columns: ["TARGET"]
                value: "updated"
              - id: replace_table
                action: replace_table_text
                workbook: table
                sheet: "BASIS"
                header_cell: "B7"
                lookup_column: "BASIS_ITEM"
                lookup_rows: ["ACC_RATIO"]
                target_columns: ["TARGET"]
                pattern: "\\\\d{{6}}"
                replacement: "202606"
              - id: replace_range
                action: replace_in_range
                workbook: table
                sheet: "BASIS"
                range: "NamedTarget"
                pattern: "named"
                replacement: "range"
              - id: replace_sheets
                action: replace_text
                workbook: table
                sheet: {{ matching: "^Other$" }}
                pattern: "202512"
                replacement: "202606"
            """,
        )
        result = run_workflow(workflow_path, working_dir=tmp_path)
        assert result.status == "success"
        updated = openpyxl.load_workbook(workbook_path)
        assert updated["BASIS"]["D8"].value == "value_202606"
        assert updated["BASIS"]["D9"].value == "updated"
        assert updated["BASIS"]["E8"].value == "range"
        assert updated["BASIS"]["G8"].value == "00123"
        assert updated["Other"]["A1"].value == "value_202606"


class TestIfConditions:
    def test_skipped_step_does_not_run(self, tmp_path: Path) -> None:
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
              - id: write_if_found
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "C1"
                value: "should not appear"
                if: "{{ steps.find_it.status == 'success' }}"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "error"  # find_it itself failed
        skipped = [s for s in result.step_results if s.step_id == "write_if_found"][0]
        assert skipped.status == "skipped"


class TestStepFailureDoesNotCrashTheRun:
    def test_a_normal_search_miss_does_not_stop_later_steps(
        self, tmp_path: Path
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
              - id: still_runs
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "error"
        statuses = {s.step_id: s.status for s in result.step_results}
        assert statuses["find_it"] == "error"
        assert statuses["still_runs"] == "success"

    def test_a_failed_run_never_commits_to_the_real_file(self, tmp_path: Path) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_b1
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "should not be committed"
              - id: find_it
                action: find_row
                workbook: manip
                sheet: "Sheet"
                column: "A"
                search_value: "does-not-exist"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "error"
        assert openpyxl.load_workbook(real)["Sheet"]["B1"].value is None


class TestExceptionsPropagateAndStillCleanUp:
    def test_a_genuine_authoring_mistake_raises_and_leaves_the_real_file_untouched(
        self, tmp_path: Path
    ) -> None:
        """write_row's positional mode without start_column raises ActionExecutionError —
        a genuine mistake, not a normal search-miss, so it must propagate, not be swallowed
        into an error StepResult."""
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: bad_write_row
                action: write_row
                workbook: manip
                sheet: "Sheet"
                row: 1
                values: ["a", "b"]
            """,
        )

        with pytest.raises(ActionExecutionError):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        assert openpyxl.load_workbook(real)["Sheet"]["A1"].value == "hello"  # untouched


class TestValidationRunsBeforeAnyWorkbookIsTouched:
    def test_invalid_workflow_raises_before_touching_the_real_file(
        self, tmp_path: Path
    ) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: bad_step
                action: read_range
                workbook: manip
                sheet: "Sheet"
            """,
        )

        with pytest.raises(ValidationError):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        assert openpyxl.load_workbook(real)["Sheet"]["A1"].value == "hello"


class TestCopyAcrossTwoWorkbooks:
    def test_copies_a_range_from_one_workbook_into_another(
        self, tmp_path: Path
    ) -> None:
        _make_workbook(tmp_path / "output" / "historical.xlsx")
        _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              historical:
                file: "{{ env.output_folder }}/historical.xlsx"
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: copy_it
                action: copy
                source:
                  workbook: historical
                  sheet: "Sheet"
                  range: "A1"
                target:
                  workbook: manip
                  sheet: "Sheet"
                  range: "D1"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        reopened = openpyxl.load_workbook(tmp_path / "output" / "manip.xlsx")
        assert reopened["Sheet"]["D1"].value == "hello"


class TestStop:
    """The `stop` control-flow action (PRD sec 6.9, Spec sec 6.1 build order item 9)."""

    def test_stop_after_a_failed_lookup_marks_later_steps_stopped_and_does_not_commit(
        self, tmp_path: Path
    ) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
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
              - id: guard
                action: stop
                reason: "lookup failed"
                if: "{{ steps.find_it.status == 'error' }}"
              - id: never_runs
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "should not be committed"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "error"  # find_it itself failed
        statuses = {s.step_id: s.status for s in result.step_results}
        assert statuses == {
            "find_it": "error",
            "guard": "success",
            "never_runs": "stopped",
        }
        guard = [s for s in result.step_results if s.step_id == "guard"][0]
        assert guard.output == {"reason": "lookup failed"}
        assert openpyxl.load_workbook(real)["Sheet"]["B1"].value is None

    def test_stop_on_a_deliberate_early_exit_still_commits_prior_work(
        self, tmp_path: Path
    ) -> None:
        """Reaching `stop` is not itself a failure — only an earlier `status: "error"` blocks
        the commit (PRD sec 6.9)."""
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_b1
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "already processed"
              - id: guard
                action: stop
              - id: never_runs
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "C1"
                value: "should not run"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        statuses = {s.step_id: s.status for s in result.step_results}
        assert statuses == {
            "write_b1": "success",
            "guard": "success",
            "never_runs": "stopped",
        }
        reopened = openpyxl.load_workbook(real)
        assert reopened["Sheet"]["B1"].value == "already processed"
        assert reopened["Sheet"]["C1"].value is None

    def test_stop_with_a_false_if_is_skipped_not_triggered(
        self, tmp_path: Path
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
              - id: guard
                action: stop
                if: "{{ false }}"
              - id: still_runs
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        statuses = {s.step_id: s.status for s in result.step_results}
        assert statuses == {"guard": "skipped", "still_runs": "success"}

    def test_stopped_steps_still_get_an_audit_record(self, tmp_path: Path) -> None:
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
              - id: guard
                action: stop
              - id: never_runs
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        lines = result.audit_log_path.read_text().splitlines()
        records = {
            json.loads(line)["step_id"]: json.loads(line)["status"] for line in lines
        }
        assert records == {"guard": "success", "never_runs": "stopped"}


class TestAuditLog:
    def test_audit_log_has_one_record_per_step(self, tmp_path: Path) -> None:
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
              - id: s1
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
              - id: s2
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "x"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        records = [
            json.loads(line) for line in result.audit_log_path.read_text().splitlines()
        ]
        # The log also carries run-level events (workbook_opened, backend_switched — Spec
        # sec 6.2, docs/backend_eligibility_build_plan.md W8), so step records are selected by
        # key rather than by counting lines.
        step_records = [record for record in records if "step_id" in record]
        assert len(step_records) == 2
        assert [record["step_id"] for record in step_records] == ["s1", "s2"]


class TestCrashSafety:
    """PRD sec 6.3/6.3.1's actual crash-safety requirement: a run interrupted mid-step must
    never leave the real files touched, must leave the scratch copies in place as the
    recovery/debugging artifact, and must not leave anything in a state that blocks a later,
    valid run against the same workbook. "No orphaned Excel process" (PRD sec 6.3) isn't
    testable yet — there's no COM backend built, so no Excel process is ever spawned by the
    current (file-backend only) action set; that part of the requirement gets a real test once
    build order item 9 exists.

    A raised exception (not an ActionResult(status="error")) is what "crashes" a run, per the
    error-handling policy established in Spec sec 4/runner.py's design — write_row's positional
    mode without start_column is used here as a realistic, already-covered way to trigger one.
    """

    def test_crash_mid_run_leaves_real_file_untouched_and_scratch_copy_in_place(
        self, tmp_path: Path
    ) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        # working_dir is now a fixed, deterministic path derived from the yaml's own filename
        # (PRD sec 6.3.4), not a fresh tempfile.mkdtemp() location per run — passing tmp_path
        # as the base gives a predictable place to find the surviving scratch copy.
        run_dir = tmp_path / "excel_runner_runs" / "workflow"

        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: write_first
                action: write_cell
                workbook: manip
                sheet: "Sheet"
                cell: "B1"
                value: "in progress"
              - id: crash_here
                action: write_row
                workbook: manip
                sheet: "Sheet"
                row: 1
                values: ["a", "b"]
            """,
        )

        with pytest.raises(ActionExecutionError):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        # the real file is completely untouched by the in-progress write
        assert openpyxl.load_workbook(real)["Sheet"]["B1"].value is None

        # the scratch copy survives as the recovery artifact, with the in-progress work intact
        scratch_file = run_dir / "scratch" / "working" / "manip.xlsx"
        assert scratch_file.exists()
        assert (
            openpyxl.load_workbook(scratch_file)["Sheet"]["B1"].value == "in progress"
        )

        # the audit log survives too — it lives outside scratch/ specifically so nothing ever
        # takes it along with a deletion (Spec sec 6.1's bug fix); nothing in working_dir is
        # ever auto-deleted now anyway (PRD sec 6.3.4)
        assert (run_dir / "audit.jsonl").exists()

    def test_a_later_valid_run_against_the_same_workbook_succeeds_after_a_crash(
        self, tmp_path: Path
    ) -> None:
        """The strongest cross-platform evidence that sessions were actually closed and no
        lingering handle survives a crash: a subsequent run against the same real file just
        works. Directly detecting an OS-level file lock would be meaningful mainly on Windows
        (PRD sec 4) and isn't reliably testable on macOS, so this is the real behavior that
        matters, tested directly instead."""
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        crashing_workflow = _write_yaml(
            tmp_path / "crash.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: crash_here
                action: write_row
                workbook: manip
                sheet: "Sheet"
                row: 1
                values: ["a", "b"]
            """,
        )
        with pytest.raises(ActionExecutionError):
            run_workflow(
                crashing_workflow,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )

        valid_workflow = _write_yaml(
            tmp_path / "valid.yaml",
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
                cell: "C1"
                value: "worked"
            """,
        )

        result = run_workflow(
            valid_workflow,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        assert openpyxl.load_workbook(real)["Sheet"]["C1"].value == "worked"


class TestDumpAction:
    """The `dump` control-flow action — internal step-output introspection while
    authoring/debugging."""

    def test_dump_prints_prior_step_output_to_console(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
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
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
              - id: show_it
                action: dump
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        printed = capsys.readouterr().out
        assert '"get_a1"' in printed
        assert '"hello"' in printed

    def test_dump_writes_to_file_when_requested(self, tmp_path: Path) -> None:
        _make_workbook(tmp_path / "output" / "manip.xlsx")
        dump_path = tmp_path / "wip_dump.json"
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            f"""
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{{{ env.output_folder }}}}/manip.xlsx"
            steps:
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
              - id: show_it
                action: dump
                to: "file"
                path: "{dump_path.as_posix()}"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        dumped = json.loads(dump_path.read_text())
        assert dumped["get_a1"]["output"]["values"] == "hello"

    def test_steps_dump_json_is_always_written_alongside_audit_log(
        self, tmp_path: Path
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
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "Sheet"
                range: "A1"
            """,
        )

        result = run_workflow(
            workflow_path,
            env_overrides={"output_folder": str(tmp_path / "output")},
            working_dir=str(tmp_path),
        )

        assert result.status == "success"
        steps_dump_path = result.audit_log_path.parent / "steps_dump.json"
        dumped = json.loads(steps_dump_path.read_text())
        assert dumped["get_a1"]["output"]["values"] == "hello"


class TestCheckExistenceFlag:
    """Tier-3 existence validation (opt-in via `check_existence=True`) — real file access,
    read-only, before any session/scratch machinery."""

    def test_missing_sheet_raises_before_any_workbook_is_touched(
        self, tmp_path: Path
    ) -> None:
        real = _make_workbook(tmp_path / "output" / "manip.xlsx")
        workflow_path = _write_yaml(
            tmp_path / "workflow.yaml",
            """
            env:
              output_folder: "./output"
            workbooks:
              manip:
                file: "{{ env.output_folder }}/manip.xlsx"
            steps:
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "NoSuchSheet"
                range: "A1"
            """,
        )

        with pytest.raises(ValidationError) as exc_info:
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
                check_existence=True,
            )

        assert "NoSuchSheet" in exc_info.value.detail.message
        assert openpyxl.load_workbook(real)["Sheet"]["A1"].value == "hello"

    def test_not_run_by_default(self, tmp_path: Path) -> None:
        """Without the flag, a missing sheet is only caught when the step actually executes,
        not upfront — confirms the check is genuinely opt-in. (A missing sheet currently
        surfaces as a raw KeyError from openpyxl at execution time, not an ActionExecutionError
        — a pre-existing gap in read_range, unrelated to this feature.)"""
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
              - id: get_a1
                action: read_range
                workbook: manip
                sheet: "NoSuchSheet"
                range: "A1"
            """,
        )

        with pytest.raises(KeyError):
            run_workflow(
                workflow_path,
                env_overrides={"output_folder": str(tmp_path / "output")},
                working_dir=str(tmp_path),
            )
