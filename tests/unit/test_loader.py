"""Unit tests for the YAML file -> Workflow pipeline in excel_runner.core (Spec sec 2.2)."""

from pathlib import Path

from excel_runner.core import load

_BASIC_WORKFLOW = """
env:
  input_folder: "./input"
  output_folder: "./output"

workbooks:
  historical:
    file: "{{ env.input_folder }}/historical.xlsx"
  manip:
    file: "{{ env.output_folder }}/manip.xlsx"
    create_if_missing: true
    template: historical

steps:
  - id: open_hist
    action: open
    workbook: historical

  - id: recalc
    action: recalculate
    workbook: manip
    mode: full
    if: "{{ steps.copy_data.status == 'success' }}"
"""


def _write(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "workflow.yaml"
    path.write_text(text)
    return path


class TestLoadBasics:
    def test_env_block_is_captured(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        assert workflow.env == {"input_folder": "./input", "output_folder": "./output"}

    def test_workbook_file_paths_are_resolved_against_env(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        assert workflow.workbooks["historical"].file == "./input/historical.xlsx"
        assert workflow.workbooks["manip"].file == "./output/manip.xlsx"

    def test_workbook_registry_fields_are_captured(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        manip = workflow.workbooks["manip"]
        assert manip.create_if_missing is True
        assert manip.template == "historical"

    def test_steps_are_captured_in_order(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        assert [step.id for step in workflow.steps] == ["open_hist", "recalc"]
        assert workflow.steps[1].action == "recalculate"

    def test_step_params_are_left_raw_and_unresolved(self, tmp_path: Path) -> None:
        """steps.* references can't be resolved at load time — no step has run yet — so
        they must survive load() as literal, untouched text (Spec sec 2.2)."""
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        recalc = workflow.steps[1]
        assert recalc.if_expr == "{{ steps.copy_data.status == 'success' }}"
        assert recalc.params == {"workbook": "manip", "mode": "full"}

    def test_id_action_and_if_are_not_duplicated_into_params(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW))
        recalc = workflow.steps[1]
        assert "id" not in recalc.params
        assert "action" not in recalc.params
        assert "if" not in recalc.params


class TestLoadEnvOverrides:
    def test_env_overrides_take_precedence_over_the_file(self, tmp_path: Path) -> None:
        workflow = load(
            _write(tmp_path, _BASIC_WORKFLOW),
            env_overrides={"input_folder": "/override/input"},
        )
        assert workflow.env["input_folder"] == "/override/input"
        assert workflow.workbooks["historical"].file == "/override/input/historical.xlsx"

    def test_env_overrides_can_add_new_keys(self, tmp_path: Path) -> None:
        workflow = load(_write(tmp_path, _BASIC_WORKFLOW), env_overrides={"run_id": "42"})
        assert workflow.env["run_id"] == "42"


class TestLoadYamlBooleanGotcha:
    """PRD sec 7's quoting note: yes/no/on/off must never silently become booleans."""

    def test_unquoted_yes_no_on_off_stay_strings(self, tmp_path: Path) -> None:
        text = """
        env: {}
        workbooks: {}
        steps:
          - id: find_headers
            action: find_headers_row
            workbook: manip
            patterns: [yes, no, on, off]
        """
        workflow = load(_write(tmp_path, text))
        assert workflow.steps[0].params["patterns"] == ["yes", "no", "on", "off"]

    def test_true_false_still_resolve_as_real_booleans(self, tmp_path: Path) -> None:
        text = """
        env: {}
        workbooks: {}
        steps:
          - id: open_hist
            action: open
            workbook: historical
            update_links: true
        """
        workflow = load(_write(tmp_path, text))
        assert workflow.steps[0].params["update_links"] is True
