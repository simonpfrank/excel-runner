"""Run-preparation and run-state layer: action discovery, session management, and the
scratch-copy execution model. See Spec sec 5.

Both validation tiers land here in a later increment (Spec sec 5.4) — this module currently
covers sec 5.1 (registry), sec 5.2 (session management), and sec 5.3 (scratch-copy model).
"""

import inspect
import shutil
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Any, Literal

from excel_runner import backends
from excel_runner.core import (
    ACTION_CAPABILITIES,
    ActionExecutionError,
    ActionResult,
    ErrorDetail,
    WorkbookRef,
    WorkbookSession,
)


@dataclass(frozen=True)
class ActionSpec:
    """A discovered action: its name, callable, capability, and parameter schema.

    Args:
        name: The action's name, matching an `action:` field in a workflow step.
        fn: The action function itself.
        capability: Which backend this action needs. "depends_on_param" is a named, single
            exception (PRD sec 7's `read_metadata`) — not a general mechanism.
        param_schema: `{"properties": {name: {"type": ...}}, "required": [...]}`, derived from
            `fn`'s signature (excluding `session`).
    """

    name: str
    fn: Callable[..., ActionResult]
    capability: Literal["file", "com", "depends_on_param"]
    param_schema: dict[str, Any]


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
            param_schema=_generate_param_schema(fn),
        )
    return registry


# --- Scratch-copy execution model (Spec sec 5.3, PRD sec 6.3.1) --------------------------


class ScratchManager:
    """Stages workbooks that will be written to into a scratch directory, and commits them
    back to their real path atomically, only on success. Operates on plain file paths — no
    knowledge of openpyxl/xlwings, so file-backend and (later) COM-backend sessions stage and
    commit through the same code path (PRD sec 6.3.1).

    Args:
        scratch_dir: Directory to stage scratch copies into. Created lazily on first `stage()`
            call — never created just by constructing a `ScratchManager`.
    """

    def __init__(self, scratch_dir: Path) -> None:
        self._scratch_dir = scratch_dir
        self._staged: dict[str, tuple[Path, Path]] = {}  # name -> (real_path, scratch_path)
        self._all_committed = False

    def stage(self, name: str, real_path: Path) -> Path:
        """Copy a workbook into the scratch dir, or reserve a scratch path for a new one.

        Args:
            name: The workbook's logical name.
            real_path: Its real file path. If it doesn't exist yet (a `create_if_missing`
                workbook), no copy happens — the caller creates the workbook directly at the
                returned scratch path instead.

        Returns:
            The scratch path to open/create the workbook at instead of `real_path`.
        """
        self._scratch_dir.mkdir(parents=True, exist_ok=True)
        scratch_path = self._scratch_dir / f"{name}{real_path.suffix}"
        if real_path.exists():
            shutil.copy2(real_path, scratch_path)
        self._staged[name] = (real_path, scratch_path)
        return scratch_path

    def commit(self, name: str) -> None:
        """Atomically move one staged workbook's scratch content back to its real path.

        Args:
            name: The workbook's logical name, as passed to `stage()`.
        """
        real_path, scratch_path = self._staged[name]
        real_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = real_path.with_name(real_path.name + ".tmp")
        shutil.copy2(scratch_path, tmp_path)
        tmp_path.replace(real_path)

    def commit_all(self) -> None:
        """Commit every staged workbook. Marks the run as fully committed for `cleanup()`."""
        for name in self._staged:
            self.commit(name)
        self._all_committed = True

    def cleanup(self, keep_on_failure: bool = True) -> None:
        """Remove the scratch directory — but only if it's safe to.

        Args:
            keep_on_failure: If True (default) and `commit_all()` hasn't successfully run,
                the scratch dir is left in place as the recovery/debugging artifact
                (PRD sec 6.3.1) instead of being deleted. Pass False to force deletion
                regardless — the success path does this explicitly, since at that point
                keeping it was never in question.
        """
        if keep_on_failure and not self._all_committed:
            return
        if self._scratch_dir.exists():
            shutil.rmtree(self._scratch_dir)


# --- Session management (Spec sec 5.2) ----------------------------------------------------


class SessionManager:
    """Tracks one WorkbookSession per logical workbook name for the duration of a run.

    Read/write mode is caller-specified, not statically inferred — tier-2 validation
    (Spec sec 5.4, not built yet) will compute that and hand it in later; for now the caller
    decides. `promote_to_com` (file -> COM, mid-run) isn't built yet either — it needs the
    COM backend, which doesn't exist until the later COM phase (PRD sec 8, Spec sec 8 item 9).

    Args:
        workbooks: The workflow's `workbooks:` registry, name to WorkbookRef.
        scratch: Where read-write sessions get staged (PRD sec 6.3.1).
    """

    def __init__(self, workbooks: dict[str, WorkbookRef], scratch: ScratchManager) -> None:
        self._workbooks = workbooks
        self._scratch = scratch
        self._sessions: dict[str, WorkbookSession] = {}

    def get_or_open(
        self, name: str, mode: Literal["read_only", "read_write"] = "read_write"
    ) -> WorkbookSession:
        """Return the session for `name`, opening it on first reference.

        A `mode="read_write"` session is staged through the scratch-copy manager first (PRD
        sec 6.3.1) — real work happens on the scratch copy, never the original, until
        `commit_all()`. A `mode="read_only"` session opens directly against the real path.

        Args:
            name: The workbook's logical name, matching a key in the `workbooks:` registry.
            mode: Ignored if a session for `name` is already open — the mode it was first
                opened with sticks for the rest of the run.

        Returns:
            The (possibly newly-opened) WorkbookSession.

        Raises:
            ActionExecutionError: If `name` isn't in the registry, or its file doesn't exist
                and `create_if_missing` isn't set.
        """
        if name in self._sessions:
            return self._sessions[name]
        if name not in self._workbooks:
            raise ActionExecutionError(
                ErrorDetail(
                    message=f'Workbook "{name}" is not declared in the workbooks: registry.',
                    technical_reason=f"SessionManager.get_or_open: unknown workbook name {name!r}",
                )
            )
        ref = self._workbooks[name]
        session = self._open_read_write(name, ref) if mode == "read_write" else self._open_read_only(name, ref)
        self._sessions[name] = session
        return session

    def _open_read_write(self, name: str, ref: WorkbookRef) -> WorkbookSession:
        real_path = Path(ref.file)
        scratch_path = self._scratch.stage(name, real_path)
        if not scratch_path.exists():
            self._create(ref, scratch_path)
        handle = backends.open_workbook(str(scratch_path), mode="read_write")
        return WorkbookSession(
            name=name, backend="file", handle=handle, path=str(scratch_path), mode="read_write",
            scratch_path=scratch_path,
        )

    def _open_read_only(self, name: str, ref: WorkbookRef) -> WorkbookSession:
        real_path = Path(ref.file)
        if not real_path.exists():
            self._create(ref, real_path)
        handle = backends.open_workbook(str(real_path), mode="read_only")
        return WorkbookSession(name=name, backend="file", handle=handle, path=str(real_path), mode="read_only")

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

    def commit_all(self) -> None:
        """Save every staged session's in-memory state, then commit scratch copies to their
        real paths (PRD sec 6.3.1). Read-only sessions were never staged and aren't touched.
        """
        for session in self._sessions.values():
            if session.scratch_path is not None:
                backends.save_workbook(session.handle, session.path)
        self._scratch.commit_all()

    def close_all(self) -> None:
        """Close every open session, attempting all of them even if some fail.

        Raises:
            ExceptionGroup: If one or more sessions failed to close. Every session still gets
                a close attempt regardless (PRD sec 6.3's crash-safety requirement) — this
                isn't a defensive catch-and-ignore, every failure is still surfaced, just
                after giving every other session a chance to close too.
        """
        errors: list[Exception] = []
        for session in self._sessions.values():
            try:
                backends.close_workbook(session.handle)
            except Exception as exc:  # noqa: BLE001 - intentional, see docstring
                errors.append(exc)
        if errors:
            raise ExceptionGroup("failed to close one or more workbook sessions", errors)
