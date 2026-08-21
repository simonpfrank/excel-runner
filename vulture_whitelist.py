"""Vulture whitelist: known false positives, not real dead code.

Dataclass fields are read via attribute access on instances (e.g. ``ref.name``), which
vulture's static analysis doesn't trace, so every dataclass field looks "unused" to it.
Exception/data classes and functions only consumed by modules not built yet (runner.py, later
increments per docs/Specification.md's build order) also flag until those land — including
several `actions.*` functions, which are only ever called dynamically via
`engine.discover_actions`'s `inspect.getmembers` scan, never by direct name from other source
code, so vulture can't see that they're used at all. Run: `vulture excel_runner
vulture_whitelist.py`.
"""

from excel_runner import actions, backends, engine, runner
from excel_runner.core import (
    ActionExecutionError,
    ActionResult,
    ErrorDetail,
    Step,
    ValidationError,
    Workflow,
    WorkbookRef,
    WorkbookSession,
    com_action,
    evaluate_condition,
    xlw_action,
)
from openpyxl.worksheet.worksheet import Worksheet

WorkbookRef.name
WorkbookRef.file
WorkbookRef.create_if_missing
WorkbookRef.template
Step.id
Step.action
Step.params
Step.if_expr
Workflow.env
Workflow.workbooks
Workflow.steps
Workflow
ErrorDetail.technical_reason
ErrorDetail.field
ErrorDetail.suggestion
ValidationError
ActionExecutionError
evaluate_condition

ActionResult.status
ActionResult.output
ActionResult.error
WorkbookSession.backend
WorkbookSession.scratch_path
WorkbookSession.dirty
com_action
xlw_action

actions.open
actions.copy
actions.read_metadata
actions.write_row
actions.stop

backends.open_workbook

engine.ActionSpec.param_schema
engine.ActionSpec.description
engine.discover_actions
engine.ScratchManager.cleanup
engine.SessionManager
engine.SessionManager.get_or_open
engine.SessionManager.close_all
engine.validate_static
engine.plan

runner.StepResult.step_id
runner.run_workflow

backends.OwnedInstanceRegistry
backends.OwnedInstanceRegistry.pids
backends.OwnedInstanceRegistry.spawn
backends.OwnedInstanceRegistry.close_owned
backends.xlw_open_workbook
backends.xlw_close_workbook
backends.xlw_save_workbook

# `backends.rename_sheet` assigns `worksheet.title = new_name` — openpyxl's own writable
# property on its Worksheet class, not something we own, but vulture can't tell that and flags
# the assignment as an "unused attribute" since nothing in our own code ever *reads* `.title`
# back. Referencing it here marks it used.
Worksheet.title

