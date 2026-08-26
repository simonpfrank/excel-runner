"""Run-preparation and run-state layer: action discovery, session management, the
scratch-copy execution model, and both validation tiers. See Spec sec 5.
"""

import difflib
import inspect
import re
import shutil
import types as pytypes
import typing
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Any, Literal

from excel_runner import backends
from excel_runner.core import (
    ACTION_CAPABILITIES,
    ACTION_WRITES,
    ActionExecutionError,
    ActionResult,
    ErrorDetail,
    Step,
    ValidationError,
    WorkbookRef,
    WorkbookSession,
    Workflow,
)


@dataclass(frozen=True)
class ActionSpec:
    """A discovered action: its name, callable, capability, description, and parameter schema.

    Args:
        name: The action's name, matching an `action:` field in a workflow step.
        fn: The action function itself.
        capability: Which backend this action needs. "depends_on_param" is a named, single
            exception (PRD sec 7's `read_metadata`) — not a general mechanism.
        description: The action's docstring, first line only — what a future agent-tool
            wrapper would show as the tool's description (PRD sec 6.1's "close to free" tool
            schema generation). Added after the fact: the field didn't exist until §6.3's
            `list_actions()` was being built and needed it.
        param_schema: `{"properties": {name: {"type": ...}}, "required": [...]}`, derived from
            `fn`'s signature (excluding `session`).
        writes: Whether this action mutates the workbook it's given — declared via
            `writes=True` on the same `@file_action`/`@xlw_action`/`@com_action` decorator
            that registers `capability` (Spec sec 5.4), not a separate hardcoded list. Used by
            `plan()` below to infer read_only vs. read_write per workbook.
    """

    name: str
    fn: Callable[..., ActionResult]
    capability: Literal["file", "xlw", "com", "depends_on_param", "none"]
    description: str
    param_schema: dict[str, Any]
    writes: bool = False


def _generate_param_schema(fn: Callable[..., ActionResult]) -> dict[str, Any]:
    """Build a param schema from an action function's signature, excluding `session`."""
    signature = inspect.signature(fn)
    properties: dict[str, Any] = {}
    required: list[str] = []
    for param_name, param in signature.parameters.items():
        if param_name == "session":
            continue
        properties[param_name] = {"type": param.annotation}
        if param.default is inspect.Parameter.empty:
            required.append(param_name)
    return {"properties": properties, "required": required}


def discover_actions(module: ModuleType) -> dict[str, ActionSpec]:
    """Scan a module for capability-tagged action functions and build the registry.

    Args:
        module: The module to scan — normally `excel_runner.actions`.

    Returns:
        Mapping of action name to `ActionSpec`.
    """
    registry: dict[str, ActionSpec] = {}
    for name, fn in inspect.getmembers(module, inspect.isfunction):
        capability = ACTION_CAPABILITIES.get(name)
        if capability is None:
            continue
        registry[name] = ActionSpec(
            name=name,
            fn=fn,
            capability=capability,
            description=_first_line(inspect.getdoc(fn) or ""),
            param_schema=_generate_param_schema(fn),
            writes=ACTION_WRITES.get(name, False),
        )
    return registry


def _first_line(docstring: str) -> str:
    return docstring.split("\n", 1)[0].strip()


# --- Scratch-copy execution model (Spec sec 5.3, PRD sec 6.3.1) --------------------------


class ScratchManager:
    """Stages workbooks that will be written to into a scratch directory, and commits them
    back to their real path, only on success. Operates on plain file paths — no knowledge of
    openpyxl/xlwings, so file-backend and (later) COM-backend sessions stage and commit through
    the same code path (PRD sec 6.3.1).

    Args:
        working_dir: The run's working directory (PRD sec 6.3.4) — scratch copies are staged
            into `working_dir/scratch/`, created lazily on first `stage()` call, never just by
            constructing a `ScratchManager`.
    """

    def __init__(self, working_dir: Path) -> None:
        self._scratch_dir = working_dir / "scratch"
        self._staged: dict[str, tuple[Path, Path, bool]] = (
            {}
        )  # name -> (real_path, scratch_path, writes)
        self._backups: dict[str, Path] = {}  # name -> .bak path, only set during a commit_all()

    def stage(self, name: str, real_path: Path, writes: bool = True) -> Path:
        """Copy a workbook into the scratch dir, or reserve a scratch path for a new one.

        Args:
            name: The workbook's logical name.
            real_path: Its real file path. If it doesn't exist yet (a `create_if_missing`
                workbook), no copy happens — the caller creates the workbook directly at the
                returned scratch path instead.
            writes: Whether this workbook may be written to and needs committing back later.
                False for a read-only session (PRD sec 6.2.3's correction — staged too now, to
                avoid holding a handle open on the real file, but never committed since
                nothing about it ever changes).

        Returns:
            The scratch path to open/create the workbook at instead of `real_path`.
        """
        self._scratch_dir.mkdir(parents=True, exist_ok=True)
        scratch_path = self._scratch_dir / f"{name}{real_path.suffix}"
        if real_path.exists():
            shutil.copy2(real_path, scratch_path)
        self._staged[name] = (real_path, scratch_path, writes)
        return scratch_path

    def commit(self, name: str) -> None:
        """Commit one staged workbook's scratch content back to its real path.

        Rename-based, not copy-based (PRD sec 6.3.3): the new content is prepared at a `.tmp`
        sibling *before* `real_path` is touched at all, so a failure at that step leaves
        `real_path` completely untouched. If `real_path` already exists, it's renamed aside to
        a `.bak` sibling (an instant, zero-copy move — the original was already sitting there
        untouched pre-commit) before the `.tmp` file is renamed into place. The `.bak` is left
        in place here — `commit_all()` deletes every one only after every workbook in the
        batch has committed successfully, or uses it to roll back on a later failure.

        Args:
            name: The workbook's logical name, as passed to `stage()`.
        """
        real_path, scratch_path, _ = self._staged[name]
        real_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = real_path.with_name(real_path.name + ".tmp")
        shutil.copy2(scratch_path, tmp_path)
        if real_path.exists():
            bak_path = real_path.with_name(real_path.name + ".bak")
            real_path.rename(bak_path)
            self._backups[name] = bak_path
        tmp_path.rename(real_path)

    def commit_all(self) -> None:
        """Commit every staged, write-intent workbook, rolling back on a later failure (PRD
        sec 6.3.3). Read-only staged workbooks (`stage(..., writes=False)`) are skipped
        entirely — nothing about them ever changes, so there's nothing to commit back.

        No separate upfront precheck pass — each workbook's commit is attempted directly. If
        one fails, every workbook already committed *in this call* is rolled back (its `.bak`
        renamed back over `real_path`, reverse order), and whether each individual rollback
        itself succeeded is recorded. A workbook whose rollback also fails needs a human — its
        `.bak` is deliberately left in place rather than deleted, so the original content is
        still recoverable from disk. On full success, every `.bak` created this call is deleted.

        Raises:
            ActionExecutionError: If any workbook's commit fails. The message names the
                workbook that failed, and — if any earlier workbook's rollback in this same
                call also failed — names which one(s) need manual attention.
        """
        committed: list[str] = []
        for name, (_, _, writes) in self._staged.items():
            if not writes:
                continue
            try:
                self.commit(name)
                committed.append(name)
            except OSError as exc:
                rollback_results = self._rollback(committed)
                needs_human = [n for n, ok in rollback_results.items() if not ok]
                rolled_back = [n for n, ok in rollback_results.items() if ok]
                message = f'Workbook "{name}" could not be committed: {exc}.'
                if rolled_back:
                    message += f" Rolled back: {', '.join(rolled_back)}."
                if needs_human:
                    message += (
                        f" MANUAL INTERVENTION NEEDED for: {', '.join(needs_human)} — "
                        "their .bak file(s) hold the original content."
                    )
                raise ActionExecutionError(
                    ErrorDetail(
                        message=message,
                        technical_reason=(
                            f"{type(exc).__name__}: {exc}; rollback_results={rollback_results!r}"
                        ),
                    )
                ) from exc
        for bak_path in self._backups.values():
            bak_path.unlink(missing_ok=True)
        self._backups.clear()

    def _rollback(self, committed_names: list[str]) -> dict[str, bool]:
        """Undo every already-committed workbook in `committed_names`, reverse order.

        Args:
            committed_names: Logical names committed so far in this `commit_all()` call.

        Returns:
            Logical name -> whether that workbook's rollback succeeded.
        """
        results: dict[str, bool] = {}
        for name in reversed(committed_names):
            real_path, _, _ = self._staged[name]
            bak_path = self._backups.get(name)
            try:
                real_path.unlink(missing_ok=True)
                if bak_path is not None:
                    bak_path.rename(real_path)
                results[name] = True
            except OSError:
                results[name] = False
        return results


# --- Session management (Spec sec 5.2) ----------------------------------------------------


def _needed_backend(
    capability: Literal["file", "xlw", "com", "depends_on_param", "none"],
) -> Literal["file", "xlw"]:
    """Which `WorkbookSession.backend` a given action capability needs (PRD sec 6.2.2).

    `"com"` needs an `xlw`-backed session too — the action itself reaches deeper via xlwings'
    `.api`, so `SessionManager` never needs a distinct backend state for it. `"depends_on_param"`
    (`read_metadata`'s own runtime resolution, Spec sec 5.1) and `"none"` (control actions, PRD
    sec 6.9 — never call `get_or_open` at all, no `workbook:` field) can't be mapped to a
    concrete backend here.

    Raises:
        ActionExecutionError: For `"depends_on_param"` or `"none"` — not yet resolvable, or
            never actually reachable from here.
    """
    if capability == "file":
        return "file"
    if capability in ("xlw", "com"):
        return "xlw"
    raise ActionExecutionError(
        ErrorDetail(
            message=f'Action capability "{capability}" can\'t be resolved to a session backend yet.',
            technical_reason=(
                f"_needed_backend: unsupported capability {capability!r} (PRD sec 6.2.2/Spec sec 5.1)"
            ),
        )
    )


class SessionManager:
    """Tracks one WorkbookSession per logical workbook name for the duration of a run.

    Read/write mode is caller-specified, not statically inferred — tier-2 validation
    (Spec sec 5.4, not built yet) will compute that and hand it in later; for now the caller
    decides. Bidirectional backend switching (file <-> xlw mid-run, PRD sec 6.2.2) is built:
    `get_or_open` switches a session's backend in place (closing/reopening its handle) rather
    than raising, when a capability doesn't match a session's current backend.

    Every xlw/com-capability session opened by this manager, for every workbook, shares one
    lazily-spawned Excel instance (`self._owned_instances`, spawned on first need) — not one
    instance per workbook. Verified empirically: two workbooks opened via `app.books.open()`
    on the same `xw.App` land in the same Excel process (same `app.pid` on both), which is
    also what lets Excel resolve/recalculate live links between them. A per-workbook instance
    would defeat that.

    Args:
        workbooks: The workflow's `workbooks:` registry, name to WorkbookRef.
        scratch: Where read-write sessions get staged (PRD sec 6.3.1).
    """

    def __init__(
        self, workbooks: dict[str, WorkbookRef], scratch: ScratchManager
    ) -> None:
        self._workbooks = workbooks
        self._scratch = scratch
        self._sessions: dict[str, WorkbookSession] = {}
        self._owned_instances = backends.OwnedInstanceRegistry()
        self._app: Any = None

    def _shared_app(self) -> Any:
        """Return the run's one shared, lazily-spawned Excel instance, spawning it on first
        need (PRD sec 6.2.1) — never one instance per workbook, see class docstring."""
        if self._app is None:
            self._app = self._owned_instances.spawn()
        return self._app

    def get_or_open(
        self,
        name: str,
        mode: Literal["read_only", "read_write"] = "read_write",
        capability: Literal["file", "xlw", "com", "depends_on_param", "none"] = "file",
    ) -> WorkbookSession:
        """Return the session for `name`, opening it on first reference.

        A `mode="read_write"` session is staged through the scratch-copy manager first (PRD
        sec 6.3.1) — real work happens on the scratch copy, never the original, until
        `commit_all()`. A `mode="read_only"` session opens directly against the real path.

        Args:
            name: The workbook's logical name, matching a key in the `workbooks:` registry.
            mode: Ignored if a session for `name` is already open — the mode it was first
                opened with sticks for the rest of the run.
            capability: The dispatching action's capability (Spec sec 5.1) — determines which
                backend (`_needed_backend`, PRD sec 6.2.2) this session must be on. Every
                session opens on the file backend today (bidirectional switching isn't built
                yet), so this only matters for detecting a mismatch, not yet for resolving one.

        Returns:
            The (possibly newly-opened, possibly just-switched) WorkbookSession, on the
            backend `capability` needs.

        Raises:
            ActionExecutionError: If `name` isn't in the registry, or its file doesn't exist
                and `create_if_missing` isn't set.
        """
        needed = _needed_backend(capability)
        if name in self._sessions:
            session = self._sessions[name]
            if session.backend != needed:
                self._switch_backend(session, needed)
            return session

        if name not in self._workbooks:
            raise ActionExecutionError(
                ErrorDetail(
                    message=f'Workbook "{name}" is not declared in the workbooks: registry.',
                    technical_reason=f"SessionManager.get_or_open: unknown workbook name {name!r}",
                )
            )
        ref = self._workbooks[name]
        session = (
            self._open_read_write(name, ref, needed)
            if mode == "read_write"
            else self._open_read_only(name, ref, needed)
        )
        self._sessions[name] = session
        return session

    def _open_handle(
        self, scratch_path: Path, mode: Literal["read_only", "read_write"], backend: Literal["file", "xlw"]
    ) -> Any:
        if backend == "file":
            return backends.open_workbook(str(scratch_path), mode=mode)
        return backends.xlw_open_workbook(self._shared_app(), str(scratch_path), mode)

    def _open_read_write(
        self, name: str, ref: WorkbookRef, backend: Literal["file", "xlw"] = "file"
    ) -> WorkbookSession:
        real_path = Path(ref.file)
        scratch_path = self._scratch.stage(name, real_path)
        if not scratch_path.exists():
            self._create(ref, scratch_path)
        handle = self._open_handle(scratch_path, "read_write", backend)
        return WorkbookSession(
            name=name,
            backend=backend,
            handle=handle,
            path=str(scratch_path),
            mode="read_write",
            scratch_path=scratch_path,
        )

    def _open_read_only(
        self, name: str, ref: WorkbookRef, backend: Literal["file", "xlw"] = "file"
    ) -> WorkbookSession:
        real_path = Path(ref.file)
        scratch_path = self._scratch.stage(name, real_path, writes=False)
        if not scratch_path.exists():
            self._create(ref, scratch_path)
        handle = self._open_handle(scratch_path, "read_only", backend)
        return WorkbookSession(
            name=name,
            backend=backend,
            handle=handle,
            path=str(scratch_path),
            mode="read_only",
            scratch_path=scratch_path,
        )

    def _switch_backend(self, session: WorkbookSession, needed: Literal["file", "xlw"]) -> None:
        """Switch an already-open session's backend in place (PRD sec 6.2.2).

        Save-then-close-then-reopen, strictly in that order, with nothing else interleaved —
        both backends' save()/close() calls are synchronous, so this ordering is what avoids a
        Windows file-lock race (the new backend opening the same scratch path before the old
        one has actually released it).

        Args:
            session: The session to switch — mutated in place (`WorkbookSession` is
                deliberately not frozen).
            needed: The backend to switch to.
        """
        if session.backend == needed:
            return
        if session.dirty:
            if session.backend == "file":
                backends.save_workbook(session.handle, session.path)
            else:
                backends.xlw_save_workbook(session.handle)
            session.dirty = False
        if session.backend == "file":
            backends.close_workbook(session.handle)
        else:
            backends.xlw_close_workbook(session.handle)
        session.handle = self._open_handle(Path(session.path), session.mode, needed)
        session.backend = needed

    def _create(self, ref: WorkbookRef, at_path: Path) -> None:
        if not ref.create_if_missing:
            raise ActionExecutionError(
                ErrorDetail(
                    message=f'Workbook file "{ref.file}" does not exist, and `create_if_missing` is not set.',
                    technical_reason=f"SessionManager: missing file {ref.file!r}, create_if_missing=False",
                )
            )
        template_path = self._workbooks[ref.template].file if ref.template else None
        at_path.parent.mkdir(parents=True, exist_ok=True)
        backends.create_workbook(str(at_path), template_path=template_path)

    def _save_dirty_staged_sessions(self) -> None:
        for session in self._sessions.values():
            if session.scratch_path is not None and session.dirty:
                if session.backend == "file":
                    backends.save_workbook(session.handle, session.path)
                else:
                    backends.xlw_save_workbook(session.handle)
                session.dirty = False

    def checkpoint(self) -> None:
        """Persist every dirty staged session's in-memory state to its scratch file.

        openpyxl writes stay in memory until an explicit save — nothing else flushes them to
        disk mid-run — so without this, a crash after a successful step would leave the
        scratch copy (the recovery artifact, PRD sec 6.3.1) no more informative than the
        original file. `runner.py` calls this after every step, not just at the end (Spec
        sec 6.1) — found necessary via a failing crash-safety integration test, not designed
        up front.
        """
        self._save_dirty_staged_sessions()

    def commit_all(self) -> None:
        """Save every staged session's in-memory state, then commit scratch copies to their
        real paths (PRD sec 6.3.1). Read-only sessions were never staged and aren't touched.
        Redundant with per-step `checkpoint()` calls in the normal case (both are dirty-gated,
        so this is mostly a no-op by the time a run reaches here) — kept anyway as the final
        safety net at the commit boundary, not something to remove just because it's usually
        a no-op.
        """
        self._save_dirty_staged_sessions()
        self._scratch.commit_all()

    def close_all(self) -> None:
        """Close every open session, then quit the shared owned Excel instance (if one was ever
        spawned), attempting all of it even if some steps fail.

        Raises:
            ExceptionGroup: If one or more sessions, or the owned Excel instance, failed to
                close. Every session still gets a close attempt regardless (PRD sec 6.3's
                crash-safety requirement) — this isn't a defensive catch-and-ignore, every
                failure is still surfaced, just after giving every other session (and the
                owned instance) a chance to close too.
        """
        errors: list[Exception] = []
        for session in self._sessions.values():
            try:
                if session.backend == "file":
                    backends.close_workbook(session.handle)
                else:
                    backends.xlw_close_workbook(session.handle)
            except Exception as exc:  # noqa: BLE001 - intentional, see docstring
                errors.append(exc)
        try:
            self._owned_instances.close_owned()
        except ExceptionGroup as exc:
            errors.extend(exc.exceptions)
        if errors:
            raise ExceptionGroup(
                "failed to close one or more workbook sessions", errors
            )


# --- Validation, tier 1: static schema (Spec sec 5.4) --------------------------------------
#
# No workbook access — PRD sec 9.1's fourth example message (checking a range against a
# workbook's *actual* defined names) needs one, which contradicts that constraint. That check
# is not implemented here; it's a documented gap, not a silent omission — see
# docs/Specification.md sec 5.4 for the correction. What's implemented below only needs the
# parsed Workflow structure and the action registry, nothing else.

_IMPLICIT_FIELDS = {
    "workbook"
}  # consumed by the (not yet built) runner before dispatch, not
# part of any action's own param_schema — see Spec sec 4.
_SCHEMA_EXEMPT_ACTIONS = {
    "copy",
    "stop",
}  # copy's raw YAML shape (source/target dicts) doesn't
# match its Python signature yet — needs the runner's
# translation layer (Spec sec 4/8 item 7). stop has no
# workbook: field at all (PRD sec 6.9), so it's exempt from
# the implicit workbook requirement every other action gets.
_STEP_REF_RE = re.compile(r"steps\.([A-Za-z_][A-Za-z0-9_]*)")


def _step_label(step: Step) -> str:
    return f'Step "{step.id}" (action: "{step.action}")'


def _matches_type(value: Any, expected: Any) -> bool:
    """Whether `value` is compatible with a param's declared type annotation.

    Handles plain types, `X | Y` unions (checks any branch matches), `Literal[...]` (checks
    membership), and generic aliases like `list[...]`/`dict[...]` (checks only the origin
    type — e.g. "is this a list at all", not "is every element the right type").
    """
    if expected is inspect.Parameter.empty or expected is Any:
        return True
    origin = typing.get_origin(expected)
    if origin is typing.Literal:
        return value in typing.get_args(expected)
    if origin is pytypes.UnionType or origin is typing.Union:
        return any(_matches_type(value, arg) for arg in typing.get_args(expected))
    if origin is not None:
        return isinstance(value, origin)
    if expected is float:
        # PEP 484's numeric tower: an int is valid anywhere a float is expected (e.g.
        # `set_column_width`'s `width: 20` in YAML, which parses as int, not `20.0`) — bool is
        # technically an int subclass too, but not a sane width/etc. value, so excluded.
        return isinstance(value, float) or (
            isinstance(value, int) and not isinstance(value, bool)
        )
    if isinstance(expected, type):
        return isinstance(value, expected)
    # No current action's signature has an annotation shape that reaches here (every real one
    # is a plain type, a Union, a Literal, or a generic alias — all handled above). Kept as an
    # explicit "don't block on an unrecognized annotation" fallback rather than removed, since
    # mypy needs a return on every path and a future action's signature might reach it.
    return True  # pragma: no cover


def _type_name(expected: Any) -> str:
    origin = typing.get_origin(expected)
    if origin is pytypes.UnionType or origin is typing.Union:
        return " or ".join(
            _type_name(arg)
            for arg in typing.get_args(expected)
            if arg is not type(None)
        )
    if origin is typing.Literal:
        return " or ".join(repr(arg) for arg in typing.get_args(expected))
    if origin is not None:
        return getattr(origin, "__name__", str(origin))
    return getattr(expected, "__name__", str(expected))


def _expects_list(expected: Any) -> bool:
    origin = typing.get_origin(expected)
    if origin is pytypes.UnionType or origin is typing.Union:
        return any(_expects_list(arg) for arg in typing.get_args(expected))
    return origin is list or expected is list


def _type_mismatch_detail(
    step: Step, field: str, value: Any, expected: Any
) -> ErrorDetail:
    suggestion = (
        f"Wrap it in [ ], e.g. [{value}]."
        if _expects_list(expected) and isinstance(value, str)
        else None
    )
    return ErrorDetail(
        message=(
            f'{_step_label(step)}: field "{field}" must be a {_type_name(expected)}, '
            f"got {type(value).__name__} ({value!r})."
        ),
        technical_reason=f"type mismatch: field {field!r} expected {expected!r}, got {type(value)!r}",
        field=field,
        suggestion=suggestion,
    )


def _find_step_refs(value: Any) -> set[str]:
    """Recursively find every `steps.<id>` reference in a raw (unresolved) param value."""
    if isinstance(value, str):
        return set(_STEP_REF_RE.findall(value))
    if isinstance(value, dict):
        refs: set[str] = set()
        for key, val in value.items():
            refs |= _find_step_refs(key)
            refs |= _find_step_refs(val)
        return refs
    if isinstance(value, list):
        refs = set()
        for item in value:
            refs |= _find_step_refs(item)
        return refs
    return set()


def _check_action_exists(
    workflow: Workflow, registry: dict[str, ActionSpec]
) -> ValidationError | None:
    for step in workflow.steps:
        if step.action not in registry:
            suggestion = difflib.get_close_matches(
                step.action, registry.keys(), n=1, cutoff=0.5
            )
            message = f'Step "{step.id}": unknown action "{step.action}".'
            if suggestion:
                message += f' Did you mean "{suggestion[0]}"?'
            return ValidationError(
                ErrorDetail(
                    message=message,
                    technical_reason=f"action {step.action!r} not found in registry",
                    field="action",
                )
            )
    return None


def _check_unknown_params(
    workflow: Workflow, registry: dict[str, ActionSpec]
) -> ValidationError | None:
    for step in workflow.steps:
        if step.action in _SCHEMA_EXEMPT_ACTIONS:
            continue
        allowed = (
            set(registry[step.action].param_schema["properties"]) | _IMPLICIT_FIELDS
        )
        extra = sorted(set(step.params) - allowed)
        if extra:
            return ValidationError(
                ErrorDetail(
                    message=f'{_step_label(step)}: field "{extra[0]}" is not a recognized parameter for this action.',
                    technical_reason=f"unexpected param {extra[0]!r} for action {step.action!r}",
                    field=extra[0],
                )
            )
    return None


def _check_required_params(
    workflow: Workflow, registry: dict[str, ActionSpec]
) -> ValidationError | None:
    for step in workflow.steps:
        if step.action in _SCHEMA_EXEMPT_ACTIONS:
            continue
        required = (
            set(registry[step.action].param_schema["required"]) | _IMPLICIT_FIELDS
        )
        missing = sorted(required - set(step.params))
        if missing:
            return ValidationError(
                ErrorDetail(
                    message=f'{_step_label(step)}: missing required field "{missing[0]}".',
                    technical_reason=f"required param {missing[0]!r} missing for action {step.action!r}",
                    field=missing[0],
                    suggestion=f'Add a "{missing[0]}:" field to this step.',
                )
            )
    return None


def _check_param_types(
    workflow: Workflow, registry: dict[str, ActionSpec]
) -> ValidationError | None:
    for step in workflow.steps:
        if step.action in _SCHEMA_EXEMPT_ACTIONS:
            continue
        properties = registry[step.action].param_schema["properties"]
        for name, value in step.params.items():
            if name in _IMPLICIT_FIELDS:
                expected: Any = str
            elif name in properties:
                expected = properties[name]["type"]
            else:
                # Unreachable via validate_static's fixed check order (_check_unknown_params
                # always runs first and would already have raised) — kept for when this
                # function is called directly/in isolation, per Spec sec 5.4's "each check
                # individually testable" intent.
                continue  # pragma: no cover
            if not _matches_type(value, expected):
                return ValidationError(
                    _type_mismatch_detail(step, name, value, expected)
                )
    return None


def _check_step_references(
    workflow: Workflow, registry: dict[str, ActionSpec]
) -> ValidationError | None:
    step_index = {step.id: i for i, step in enumerate(workflow.steps)}
    for i, step in enumerate(workflow.steps):
        refs = _find_step_refs(step.params) | (
            _find_step_refs(step.if_expr) if step.if_expr else set()
        )
        for ref in sorted(refs):
            if ref not in step_index:
                suggestion = difflib.get_close_matches(
                    ref, list(step_index), n=1, cutoff=0.5
                )
                message = f'{_step_label(step)}: references step id "{ref}", which does not exist.'
                if suggestion:
                    message += f' Did you mean "{suggestion[0]}" (defined at step {step_index[suggestion[0]] + 1})?'
                return ValidationError(
                    ErrorDetail(
                        message=message,
                        technical_reason=f"unknown step reference {ref!r}",
                    )
                )
            if step_index[ref] >= i:
                return ValidationError(
                    ErrorDetail(
                        message=(
                            f'{_step_label(step)}: references step id "{ref}", which is defined at or after '
                            "this step. A step can only reference an earlier step's output."
                        ),
                        technical_reason=f"forward/self reference to step {ref!r}",
                    )
                )
    return None


_STATIC_CHECKS: list[
    Callable[[Workflow, dict[str, ActionSpec]], ValidationError | None]
] = [
    _check_action_exists,
    _check_unknown_params,
    _check_required_params,
    _check_param_types,
    _check_step_references,
]


def validate_static(workflow: Workflow, registry: dict[str, ActionSpec]) -> None:
    """Tier-1 validation: structural checks against the raw parsed Workflow, no workbook access.

    Args:
        workflow: The parsed workflow to validate.
        registry: The action registry (from `discover_actions`) to check against.

    Raises:
        ValidationError: On the first problem found, in check order — action existence,
            unrecognized params, missing required params, param type mismatches, then
            step-id references (existence and ordering).
    """
    for check in _STATIC_CHECKS:
        error = check(workflow, registry)
        if error is not None:
            raise error


# --- Validation, tier 2: dry-run / step-graph (Spec sec 5.4) --------------------------------


@dataclass(frozen=True)
class ExecutionPlan:
    """Per-workbook mode, inferred from a static pass over the step list (PRD sec 6.3).

    Args:
        modes: Workbook name to "read_only" or "read_write" — read_write iff some step's
            action writes to it (or, for `copy`, iff it's referenced at all — see `plan()`).
    """

    modes: dict[str, Literal["read_only", "read_write"]]


def _find_workbook_names(value: Any) -> set[str]:
    """Recursively find every value under a key literally named "workbook" in a raw param
    structure — handles both a flat `workbook:` field and copy's nested `source`/`target`
    dicts with the same generic walk.
    """
    names: set[str] = set()
    if isinstance(value, dict):
        for key, val in value.items():
            if key == "workbook" and isinstance(val, str):
                names.add(val)
            else:
                names |= _find_workbook_names(val)
    elif isinstance(value, list):
        for item in value:
            names |= _find_workbook_names(item)
    return names


def _check_workbooks_declared(workflow: Workflow) -> ValidationError | None:
    for step in workflow.steps:
        for name in sorted(_find_workbook_names(step.params)):
            if name not in workflow.workbooks:
                return ValidationError(
                    ErrorDetail(
                        message=(
                            f'{_step_label(step)}: references workbook "{name}", which is not declared '
                            "in the workbooks: registry."
                        ),
                        technical_reason=f"undeclared workbook {name!r}",
                    )
                )
    return None


def plan(workflow: Workflow, registry: dict[str, ActionSpec]) -> ExecutionPlan:
    """Tier-2 validation: reasons over the whole step list together, still no workbook access.

    Args:
        workflow: The parsed workflow to plan.
        registry: The action registry (Spec sec 5.1) — consulted for each step's `writes` flag.

    Returns:
        An `ExecutionPlan` with an inferred read/write mode per declared workbook.

    Raises:
        ValidationError: If a step references a workbook not in the `workbooks:` registry.
    """
    error = _check_workbooks_declared(workflow)
    if error is not None:
        raise error
    modes: dict[str, Literal["read_only", "read_write"]] = dict.fromkeys(
        workflow.workbooks, "read_only"
    )
    for step in workflow.steps:
        names = _find_workbook_names(step.params)
        if step.action == "copy":
            # Can't statically tell copy's source from its target without the runner's
            # translation layer (Spec sec 4) — the safe fallback (PRD sec 6.3) is read_write
            # for every workbook it touches, rather than silently under-provisioning one.
            for name in names:
                modes[name] = "read_write"
        elif step.action in registry and registry[step.action].writes:
            for name in names:
                modes[name] = "read_write"
    return ExecutionPlan(modes=modes)
