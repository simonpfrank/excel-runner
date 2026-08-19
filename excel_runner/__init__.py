"""excel_runner — declarative, YAML-driven Excel automation.

The public API surface (Spec sec 6.3): everything importable directly from this package.
Nothing else in `excel_runner`'s submodules is a stable contract — only these names are safe
to depend on from other code (PRD sec 3/sec 9's importable-library goal).

    from excel_runner import run_workflow
    result = run_workflow("workflow.yaml", env_overrides={"output_folder": "/tmp/out"})
"""

from excel_runner.core import Step, WorkbookRef, Workflow
from excel_runner.engine import ActionSpec
from excel_runner.runner import RunResult, StepResult, list_actions, run_workflow

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
