"""Core layer: the workflow data model, error types, and the loading/templating pipeline.

See docs/Specification.md sec 2 for the full design.

There is no whole-file "render as text, then parse" step. {{ }} expressions are resolved per
field: once at load time for env:/workbooks: fields (env-only context), and once per step
during execution for step params/`if:` (env + accumulated step-output context) — see the
module docstring correction recorded in docs/PRD.md sec 10.1 and docs/Specification.md sec 2.2.
"""

import re
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, ParamSpec

import jinja2
import yaml


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


# --- Templating -------------------------------------------------------------------------

class _DictItemFirstEnvironment(jinja2.Environment):
    """A dict field named "values"/"keys"/"items"/etc. would otherwise be shadowed by the real
    dict method of the same name — Jinja2's default `getattr` tries attribute access before
    item access, and every dict has real `.values()`/`.keys()`/etc. methods. Found via a real
    bug: `{{ steps.x.output.values }}` returned the bound `dict.values` method instead of the
    output dict's "values" entry (PRD sec 10.4's own output-shape convention for `read_range`).
    Every value flowing through our templates — env, steps, an action's output — is a plain
    dict/list, never an object whose real attributes we'd want prioritized over its data, so
    item-first is the *correct* order for this templating engine's actual use, not a patch
    around one unlucky field name.
    """

    def getattr(self, obj: Any, attribute: str) -> Any:
        try:
            return obj[attribute]
        except (TypeError, LookupError):
            return super().getattr(obj, attribute)


_ENV = _DictItemFirstEnvironment(undefined=jinja2.StrictUndefined)
_WHOLE_EXPRESSION_RE = re.compile(r"^\{\{(.*)\}\}$", re.DOTALL)


def _whole_expression(text: str) -> str | None:
    """Return the inner expression if `text` is entirely one {{ }} block, else None."""
    match = _WHOLE_EXPRESSION_RE.match(text.strip())
    if match is None:
        return None
    inner = match.group(1)
    if "{{" in inner or "}}" in inner:
        return None  # more than one block — treat as an embedded/partial string instead
    return inner.strip()


def _wrap_template_error(original_text: str, exc: Exception) -> ValidationError:
    stripped = original_text.strip()
    if isinstance(exc, jinja2.exceptions.UndefinedError):
        message = f'"{stripped}" references something that does not exist: {exc}'
    else:
        message = f'"{stripped}" is not a valid template expression: {exc}'
    return ValidationError(ErrorDetail(message=message, technical_reason=f"{type(exc).__name__}: {exc}"))


def _evaluate_expression(expr_text: str, context: dict[str, Any], original_text: str) -> Any:
    try:
        expression = _ENV.compile_expression(expr_text, undefined_to_none=False)
        return expression(**context)
    except (jinja2.exceptions.UndefinedError, jinja2.exceptions.TemplateSyntaxError) as exc:
        raise _wrap_template_error(original_text, exc) from exc


def _resolve_string(text: str, context: dict[str, Any]) -> Any:
    if "{{" not in text and "{%" not in text and "{#" not in text:
        return text
    inner = _whole_expression(text)
    if inner is not None:
        return _evaluate_expression(inner, context, text)
    try:
        return _ENV.from_string(text).render(**context)
    except (jinja2.exceptions.UndefinedError, jinja2.exceptions.TemplateSyntaxError) as exc:
        raise _wrap_template_error(text, exc) from exc


def resolve_value(value: Any, context: dict[str, Any]) -> Any:
    """Resolve ``{{ }}`` templating in a value, recursing through dicts and lists.

    A value that is entirely one ``{{ }}`` expression resolves to the native Python object
    it evaluates to; an expression embedded in a larger string always stringifies
    (Ansible-style native type preservation — see docs/PRD.md sec 10.1).

    Args:
        value: The raw value — a string, dict, list, or any other scalar.
        context: Template context, e.g. ``{"env": {...}}`` or
            ``{"env": {...}, "steps": {...}}``.

    Returns:
        The resolved value, same shape as the input for dicts/lists.

    Raises:
        ValidationError: If the value references an undefined variable, or contains a
            syntactically invalid expression.
    """
    if isinstance(value, str):
        return _resolve_string(value, context)
    if isinstance(value, dict):
        return {resolve_value(k, context): resolve_value(v, context) for k, v in value.items()}
    if isinstance(value, list):
        return [resolve_value(item, context) for item in value]
    return value


def evaluate_condition(if_expr: str, context: dict[str, Any]) -> bool:
    """Evaluate a step's ``if:`` expression to a bool.

    Args:
        if_expr: The raw condition, with or without a surrounding ``{{ }}`` wrapper.
        context: Template context (env + accumulated step outputs).

    Returns:
        The Python-truthy value of the evaluated expression.

    Raises:
        ValidationError: If the expression references an undefined variable, or is
            syntactically invalid.
    """
    inner = _whole_expression(if_expr)
    expr_text = inner if inner is not None else if_expr
    return bool(_evaluate_expression(expr_text, context, if_expr))


# --- Loading ------------------------------------------------------------------------------


class _Yaml12BoolLoader(yaml.SafeLoader):
    """SafeLoader without YAML 1.1's yes/no/on/off boolean coercion (PRD sec 7's quoting note).

    ``yes``/``no``/``on``/``off`` stay plain strings; only true/false variants are booleans.
    """


_Yaml12BoolLoader.yaml_implicit_resolvers = {
    first_char: [(tag, regexp) for tag, regexp in resolvers if tag != "tag:yaml.org,2002:bool"]
    for first_char, resolvers in yaml.SafeLoader.yaml_implicit_resolvers.items()
}
_Yaml12BoolLoader.add_implicit_resolver(
    "tag:yaml.org,2002:bool",
    re.compile(r"^(?:true|True|TRUE|false|False|FALSE)$"),
    list("tTfF"),
)


def _build_step(raw_step: dict[str, Any]) -> Step:
    params = {k: v for k, v in raw_step.items() if k not in {"id", "action", "if"}}
    return Step(id=raw_step["id"], action=raw_step["action"], params=params, if_expr=raw_step.get("if"))


def load(path: str | Path, env_overrides: dict[str, Any] | None = None) -> Workflow:
    """Load and parse a workflow YAML file.

    ``env:``/``workbooks:`` fields are resolved immediately (env-only context); step params
    and ``if:`` are left raw — they may reference another step's output, which doesn't exist
    until execution reaches that step (see the module docstring).

    Args:
        path: Path to the workflow YAML file.
        env_overrides: Values merged over (and taking precedence over) the file's own
            ``env:`` block — how an external caller parameterizes a run (PRD sec 6.6).

    Returns:
        The parsed Workflow, with workbook paths resolved and step params left raw.
    """
    raw_text = Path(path).read_text()
    raw = yaml.load(raw_text, Loader=_Yaml12BoolLoader) or {}

    env: dict[str, Any] = {**(raw.get("env") or {}), **(env_overrides or {})}
    context = {"env": env}

    workbooks: dict[str, WorkbookRef] = {}
    for name, entry in (raw.get("workbooks") or {}).items():
        resolved = resolve_value(entry, context)
        workbooks[name] = WorkbookRef(
            name=name,
            file=resolved["file"],
            create_if_missing=resolved.get("create_if_missing", False),
            template=resolved.get("template"),
        )

    steps = tuple(_build_step(raw_step) for raw_step in raw.get("steps") or [])

    return Workflow(env=env, workbooks=workbooks, steps=steps)


# --- Execution-time types ----------------------------------------------------------------
#
# Defined here, not in engine.py/actions.py (Spec sec 5.1/sec 4's original homes), to avoid a
# circular import: engine.py's registry must import actions.py to discover its functions, and
# actions.py's functions are typed against ActionResult/WorkbookSession — putting the shared
# types in core.py (which neither engine.py nor actions.py depend on each other for) keeps the
# dependency graph a clean line: core.py <- {backends.py, actions.py} <- engine.py <- runner.py.


@dataclass(frozen=True)
class ActionResult:
    """The result of running one action. Always a keyed object (PRD sec 10.4) — even a
    single-value output like find_row's `row` lives under a named key, never returned bare.

    Args:
        status: Whether the action succeeded.
        output: The action's result data. Empty dict if the action has no meaningful output.
        error: Present when status is "error".
    """

    status: Literal["success", "error"]
    output: dict[str, Any]
    error: ErrorDetail | None = None


@dataclass
class WorkbookSession:
    """A workbook currently open for the duration of a run. Deliberately not frozen — this
    models live, mutable run state (PRD sec 6.5's "avoid mutable dicts for state" is about
    ad-hoc dicts; a real class with named fields is exactly the alternative it asks for).

    Args:
        name: Logical workbook name, matching a WorkbookRef.name.
        backend: Which backend currently holds this workbook open.
        handle: The live backend object (an openpyxl Workbook, or later an xlwings Book).
        path: The file path the backend is currently pointed at — the real path until the
            scratch-copy execution model (PRD sec 6.3.1, not built yet) starts routing
            writes through `scratch_path` instead.
        mode: Whether this session was opened read-only or read-write.
        scratch_path: Set once the scratch-copy execution model is built. None means work
            happens directly against `path`.
        dirty: Whether a write has happened since the last save.
    """

    name: str
    backend: Literal["file", "com"]
    handle: Any
    path: str
    mode: Literal["read_only", "read_write"]
    scratch_path: Path | None = None
    dirty: bool = False


# --- Action capability tagging ------------------------------------------------------------
#
# A plain name->capability dict, populated by decorators, rather than attributes stamped onto
# the function object — keeps mypy --strict happy (no dynamic-attribute type: ignore noise)
# and keeps the registration mechanism trivially introspectable (engine.py's discover_actions
# just reads this dict). "depends_on_param" (PRD sec 7's read_metadata exception) gets its own
# decorator when that action is built — not added speculatively now.

ACTION_CAPABILITIES: dict[str, Literal["file", "com", "depends_on_param"]] = {}

_P = ParamSpec("_P")


def file_action(fn: Callable[_P, ActionResult]) -> Callable[_P, ActionResult]:
    """Register a function as a file-backend (openpyxl) action for discovery (Spec sec 5.1).

    Typed with ParamSpec, not `Callable[..., ActionResult]`, so the decorated function keeps
    its real parameter signature for mypy — found the hard way: an untyped `...` erased every
    action's params to "accepts anything", silently defeating static type checking at every
    call site, not just inside the action itself.
    """
    ACTION_CAPABILITIES[fn.__name__] = "file"
    return fn


def com_action(fn: Callable[_P, ActionResult]) -> Callable[_P, ActionResult]:
    """Register a function as a COM-backend (xlwings) action for discovery (Spec sec 5.1)."""
    ACTION_CAPABILITIES[fn.__name__] = "com"
    return fn
