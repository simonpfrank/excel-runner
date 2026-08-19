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

from excel_runner import actions, backends, engine
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
)

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

actions.open
actions.copy
actions.read_metadata
actions.write_row

backends.open_workbook

engine.ActionSpec.param_schema
engine.discover_actions
