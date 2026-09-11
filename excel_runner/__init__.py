"""excel_runner — declarative, YAML-driven Excel automation.

The public API surface (Spec sec 6.3): everything importable directly from this package.
Nothing else in `excel_runner`'s submodules is a stable contract — only these names are safe
to depend on from other code (PRD sec 3/sec 9's importable-library goal).

    from excel_runner import run_workflow
    result = run_workflow("workflow.yaml", env_overrides={"output_folder": "/tmp/out"})
"""
import sys
from importlib import import_module
from pathlib import Path

_PACKAGE_DIRECTORY = str(Path(__file__).resolve().parent)
if _PACKAGE_DIRECTORY not in sys.path:
    sys.path.insert(0, _PACKAGE_DIRECTORY)

_core = import_module("core")
sys.modules[f"{__name__}.core"] = _core
core = _core
_backends = import_module("backends")
sys.modules[f"{__name__}.backends"] = _backends
backends = _backends
_actions = import_module("actions")
sys.modules[f"{__name__}.actions"] = _actions
actions = _actions
_engine = import_module("engine")
sys.modules[f"{__name__}.engine"] = _engine
engine = _engine
_runner = import_module("runner")
sys.modules[f"{__name__}.runner"] = _runner
runner = _runner

Step = _core.Step
WorkbookRef = _core.WorkbookRef
Workflow = _core.Workflow
ActionSpec = _engine.ActionSpec
RunResult = _runner.RunResult
StepResult = _runner.StepResult
list_actions = _runner.list_actions
run_workflow = _runner.run_workflow

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
