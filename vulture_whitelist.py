"""Vulture whitelist: known false positives, not real dead code.

Dataclass fields are read via attribute access on instances (e.g. ``ref.name``), which
vulture's static analysis doesn't trace, so every dataclass field looks "unused" to it.
Exception/data classes only consumed by modules not built yet (engine.py, runner.py, per
docs/Specification.md's build order) also flag until those land. Run: `vulture excel_runner
vulture_whitelist.py`.
"""

from excel_runner.core import ActionExecutionError, ErrorDetail, Step, ValidationError, Workflow, WorkbookRef

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
