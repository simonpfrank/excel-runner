"""excel_runner — declarative, YAML-driven Excel automation.

The public API surface (Spec sec 6.3): everything importable directly from this package.
Nothing else in `excel_runner`'s submodules is a stable contract — only these names are safe
to depend on from other code (PRD sec 3/sec 9's importable-library goal).

    from excel_runner import run_workflow
    result = run_workflow("workflow.yaml", env_overrides={"output_folder": "/tmp/out"})
"""
from . import core, engine, runner

Step = core.Step
WorkbookRef = core.WorkbookRef
Workflow = core.Workflow
ActionSpec = engine.ActionSpec
RunResult = runner.RunResult
StepResult = runner.StepResult
list_actions = runner.list_actions
run_workflow = runner.run_workflow

__all__ = [
    "ActionSpec",
    "RunResult",
    "Step",
    "StepResult",
    "Workflow",
    "WorkbookRef",
    "list_actions",
    "run_workflow",
]
