"""Core, I/O-free layer: the workflow data model and error types.

See docs/Specification.md sec 2 for the full design. Loading/templating (sec 2.2) is added in
a later increment — this module currently covers sec 2.1 (data model) and sec 2.3 (errors).
"""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class WorkbookRef:
    """A logical workbook declared in a workflow's ``workbooks:`` registry.

    Args:
        name: Logical name, used as the registry key and referenced by steps.
        file: Path to the workbook file. May contain ``{{ env.* }}`` templating.
        create_if_missing: Create the file on first reference if it doesn't exist.
        template: Logical name of another WorkbookRef to copy from when creating.
    """

    name: str
    file: str
    create_if_missing: bool = False
    template: str | None = None


@dataclass(frozen=True)
class Step:
    """A single step in a workflow.

    Args:
        id: Unique step identifier, referenced by later steps via templating.
        action: Name of the action to run, resolved against the action registry.
        params: Raw, not-yet-validated parameters for the action.
        if_expr: Raw Jinja2 boolean expression gating whether the step runs.
            Unevaluated at parse time — resolved during execution.
    """

    id: str
    action: str
    params: dict[str, Any]
    if_expr: str | None = None


@dataclass(frozen=True)
class Workflow:
    """A fully parsed workflow: environment, workbook registry, and ordered steps.

    Args:
        env: Environment values available to ``{{ env.* }}`` templating.
        workbooks: Logical workbook name to WorkbookRef.
        steps: Steps in execution order.
    """

    env: dict[str, Any]
    workbooks: dict[str, WorkbookRef]
    steps: tuple[Step, ...]


@dataclass(frozen=True)
class ErrorDetail:
    """A structured error, split into a plain-English message and its technical cause.

    Args:
        message: Plain-English explanation a human or an AI agent can act on directly.
        technical_reason: The original exception type/message — never shown as the headline.
        field: Name of the offending field, if the error is field-specific.
        suggestion: A concrete fix, if one can be offered.
    """

    message: str
    technical_reason: str
    field: str | None = None
    suggestion: str | None = None


class ExcelRunnerError(Exception):
    """Base class for all excel_runner errors. Carries a structured ErrorDetail.

    Args:
        detail: The structured error detail. ``str(error)`` returns ``detail.message``.
    """

    def __init__(self, detail: ErrorDetail) -> None:
        super().__init__(detail.message)
        self.detail = detail


class ValidationError(ExcelRunnerError):
    """A workflow failed schema or step-graph validation before any workbook was touched."""


class ActionExecutionError(ExcelRunnerError):
    """An action failed while a workflow was running."""
