"""Unit tests for the public API surface (Spec sec 6.3): everything importable directly from
`excel_runner`, and nothing more than intended — everything else in the package tree is an
implementation detail that may change without notice (PRD sec 3/sec 9's stable-surface goal).
"""

import excel_runner
from excel_runner import core, engine, runner


class TestPublicSurface:
    def test_run_workflow_is_the_real_function(self) -> None:
        assert excel_runner.run_workflow is runner.run_workflow

    def test_list_actions_is_the_real_function(self) -> None:
        assert excel_runner.list_actions is runner.list_actions

    def test_result_types_are_the_real_classes(self) -> None:
        assert excel_runner.RunResult is runner.RunResult
        assert excel_runner.StepResult is runner.StepResult

    def test_workflow_construction_types_are_the_real_classes(self) -> None:
        assert excel_runner.Workflow is core.Workflow
        assert excel_runner.Step is core.Step
        assert excel_runner.WorkbookRef is core.WorkbookRef

    def test_action_spec_is_the_real_class(self) -> None:
        """Needed to work with list_actions()'s return value meaningfully — its elements are
        ActionSpec instances, so the type itself has to be part of the public surface too."""
        assert excel_runner.ActionSpec is engine.ActionSpec

    def test_list_actions_works_through_the_public_import(self) -> None:
        specs = excel_runner.list_actions()
        assert len(specs) > 0
        assert all(isinstance(spec, excel_runner.ActionSpec) for spec in specs)
