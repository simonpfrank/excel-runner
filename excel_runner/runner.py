"""The composition root, audit logging, and (once built, §6.3) the public API surface.
See docs/Specification.md sec 6.

This module currently covers sec 6.1 (orchestration) and sec 6.2 (audit logging) — the public
surface (sec 6.3, build order item 8) isn't built yet.
"""

import json
import logging
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

from excel_runner import actions as actions_module
from excel_runner import core, engine
from excel_runner.core import ActionResult, ErrorDetail, Step, WorkbookSession, Workflow

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class StepResult:
    """The outcome of one step — a superset of ActionResult's statuses, since a step can also
    be "skipped" (its `if:` was false), which never reaches an action at all.

    Args:
        step_id: The step's id.
        status: "success", "error" (a normal, anticipated "didn't work" outcome — see the
            error-handling policy in Spec sec 4 — or an exception would have been raised
            instead and this StepResult would never be constructed), "skipped" (this step's own
            `if:` was false), or "stopped" (a `stop` step ended the run before this step was
            ever reached — PRD sec 6.9).
        output: The action's output, or `{}` if skipped or stopped.
        error: Present when status is "error".
    """

    step_id: str
    status: Literal["success", "error", "skipped", "stopped"]
    output: dict[str, Any]
    error: ErrorDetail | None = None


@dataclass(frozen=True)
class RunResult:
    """The outcome of a full run.

    Args:
        status: "success" iff every step succeeded, was deliberately skipped, or never ran
            because a `stop` step (PRD sec 6.9) ended the run early — "error" if any *dispatched*
            step's action returned an error result, even though the run continued past it.
        step_results: One StepResult per step, in execution order.
        audit_log_path: Where the structured, per-step JSONL audit log was written.
    """

    status: Literal["success", "error"]
    step_results: tuple[StepResult, ...]
    audit_log_path: Path


class AuditLogger:
    """Writes one structured JSON record per step to a JSONL file (Spec sec 6.2).

    Args:
        path: Where to write the audit log. Parent directories are created if needed.
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._path.parent.mkdir(parents=True, exist_ok=True)
        # Truncate any previous run's leftover audit.jsonl at this same fixed working_dir path
        # (PRD sec 6.3.4) — a new run's records must never mix with an old run's.
        self._path.write_text("")

    def record_step(
        self, step: Step, result: StepResult, started_at: Any, ended_at: Any
    ) -> None:
        """Append one step's outcome to the audit log.

        Args:
            step: The step that ran (or was skipped).
            result: Its outcome.
            started_at: When the step started — logged as-is (a datetime or a plain string).
            ended_at: When the step ended — logged as-is.
        """
        record = {
            "step_id": step.id,
            "action": step.action,
            "params": step.params,
            "status": result.status,
            "output": result.output,
            "error": asdict(result.error) if result.error is not None else None,
            "started_at": str(started_at),
            "ended_at": str(ended_at),
        }
        with self._path.open("a") as handle:
            handle.write(json.dumps(record, default=str) + "\n")


def _dispatch_copy(
    resolved_params: dict[str, Any],
    session_manager: engine.SessionManager,
    plan: engine.ExecutionPlan,
) -> ActionResult:
    """Resolve copy's two nested workbook refs into two sessions and call the action.

    `copy` is the one action whose YAML shape (`source:`/`target:` dicts) doesn't map to a
    single `workbook:` field — see Spec sec 4's correction — so it's handled separately from
    every other action's dispatch.
    """
    source = resolved_params["source"]
    target = resolved_params["target"]
    source_session = session_manager.get_or_open(
        source["workbook"], mode=plan.modes[source["workbook"]]
    )
    target_session = session_manager.get_or_open(
        target["workbook"], mode=plan.modes[target["workbook"]]
    )
    return actions_module.copy(
        session=source_session,
        target=target_session,
        source_sheet=source["sheet"],
        source_range=source.get("range"),
        target_sheet=target["sheet"],
        target_range=target["range"],
    )


def _dispatch(
    step: Step,
    context: dict[str, Any],
    registry: dict[str, engine.ActionSpec],
    session_manager: engine.SessionManager,
    plan: engine.ExecutionPlan,
) -> ActionResult:
    """Resolve one step's params and call its action.

    `workbook` is resolved into a `WorkbookSession` and never forwarded to the action function
    itself (Spec sec 4) — every other resolved param is passed through as a keyword argument.
    """
    resolved = core.resolve_value(step.params, context)
    logger.debug('Step "%s" (%s): resolved params %r', step.id, step.action, resolved)
    if step.action == "copy":
        return _dispatch_copy(resolved, session_manager, plan)
    if step.action == "stop":
        # No session to resolve — stop is pure control flow, no workbook: field (PRD sec 6.9).
        return registry[step.action].fn(**resolved)
    workbook_name = resolved["workbook"]
    session: WorkbookSession = session_manager.get_or_open(
        workbook_name,
        mode=plan.modes[workbook_name],
        capability=registry[step.action].capability,
    )
    kwargs = {key: value for key, value in resolved.items() if key != "workbook"}
    return registry[step.action].fn(session=session, **kwargs)


def run_workflow(
    path: str | Path, env_overrides: dict[str, Any] | None = None, working_dir: str | Path | None = None
) -> RunResult:
    """Load, validate, and execute a workflow YAML file end to end.

    A step's action returning `ActionResult(status="error")` is a normal, anticipated outcome
    (Spec sec 4's error-handling policy) — the run continues so later steps can branch on it
    via `if:` — but the overall `RunResult.status` is still "error" if any step failed, and the
    run is never committed to the real workbook files in that case (PRD sec 6.3.1). A raised
    exception (a genuine authoring mistake, not a normal "didn't work") propagates to the
    caller instead of being captured in the result — session cleanup still runs regardless.

    Args:
        path: Path to the workflow YAML file.
        env_overrides: Values merged over (and taking precedence over) the file's own `env:`
            block (PRD sec 6.6).
        working_dir: Base directory this run's `working_dir` folder
            (`excel_runner_runs/<yaml_stem>/`) is created under (PRD sec 6.3.4). Defaults to
            the current working directory. Fed by the CLI's `--working-dir` flag (Spec sec 6.4).

    Returns:
        The RunResult, once every step has run (or been skipped) without raising.

    Raises:
        ValidationError: If the workflow fails tier-1 or tier-2 validation. Nothing is opened
            or touched before this check completes.
        ActionExecutionError: If an action raises during execution — a genuine mistake, not a
            normal "didn't work" outcome. The real workbook files are left untouched either way.
    """
    workflow: Workflow = core.load(path, env_overrides)
    registry = engine.discover_actions(actions_module)
    engine.validate_static(workflow, registry)
    plan = engine.plan(workflow, registry)

    # working_dir is a fixed, predictable path (not a random tempfile.mkdtemp() one) so
    # external tooling can construct it itself from just the yaml's filename, without reading
    # any output field (PRD sec 6.3.4). Nothing under it is ever auto-deleted — see
    # ScratchManager.commit_all() and AuditLogger's truncate-on-open below.
    base = Path(working_dir) if working_dir is not None else Path.cwd()
    run_dir = base / "excel_runner_runs" / Path(path).stem
    scratch = engine.ScratchManager(run_dir)
    session_manager = engine.SessionManager(workflow.workbooks, scratch)
    audit_log_path = run_dir / "audit.jsonl"
    audit = AuditLogger(audit_log_path)

    step_outputs: dict[str, dict[str, Any]] = {}
    step_results: list[StepResult] = []
    any_failed = False

    try:
        for i, step in enumerate(workflow.steps):
            context = {"env": workflow.env, "steps": step_outputs}
            started_at = datetime.now()
            stop_triggered = False

            if step.if_expr is not None and not core.evaluate_condition(
                step.if_expr, context
            ):
                step_result = StepResult(step_id=step.id, status="skipped", output={})
                logger.info('Step "%s" (%s): skipped (if: was false)', step.id, step.action)
            else:
                logger.info('Step "%s" (%s): starting', step.id, step.action)
                action_result = _dispatch(
                    step, context, registry, session_manager, plan
                )
                step_result = StepResult(
                    step_id=step.id,
                    status=action_result.status,
                    output=action_result.output,
                    error=action_result.error,
                )
                if action_result.status == "error":
                    any_failed = True
                    error_message = action_result.error.message if action_result.error else ""
                    logger.error('Step "%s" (%s): failed — %s', step.id, step.action, error_message)
                else:
                    logger.info('Step "%s" (%s): %s', step.id, step.action, action_result.status)
                stop_triggered = step.action == "stop"

            audit.record_step(step, step_result, started_at, datetime.now())
            step_results.append(step_result)
            step_outputs[step.id] = {
                "status": step_result.status,
                "output": step_result.output,
            }
            # Persist this step's writes to the scratch file now, not just at the end — so a
            # later step crashing still leaves everything that succeeded so far visible in the
            # recovery artifact (PRD sec 6.3.1). Found necessary via a failing crash-safety
            # integration test: without this, an in-memory-only write is invisible on disk
            # until commit_all(), which never runs on a crash.
            session_manager.checkpoint()

            if stop_triggered:
                # A stop step (PRD sec 6.9) ends the run right here — every later step gets a
                # "stopped" StepResult (distinct from "skipped": its own if: never even ran)
                # instead of being dispatched at all, so RunResult.step_results still has one
                # entry per workflow step.
                for later_step in workflow.steps[i + 1 :]:
                    later_now = datetime.now()
                    later_result = StepResult(
                        step_id=later_step.id, status="stopped", output={}
                    )
                    audit.record_step(later_step, later_result, later_now, later_now)
                    step_results.append(later_result)
                    step_outputs[later_step.id] = {"status": "stopped", "output": {}}
                break

        if not any_failed:
            session_manager.commit_all()
    finally:
        session_manager.close_all()

    return RunResult(
        status="error" if any_failed else "success",
        step_results=tuple(step_results),
        audit_log_path=audit_log_path,
    )


def list_actions() -> tuple[engine.ActionSpec, ...]:
    """List every built action's name, capability, description, and parameter schema.

    What a future agent/CLI wrapper would iterate over to generate its own tool definitions
    (PRD sec 6.1's "close to free" schema reuse) — just `discover_actions` wired to the real
    `actions` module, not a second source of truth.

    Returns:
        One `ActionSpec` per built action.
    """
    return tuple(engine.discover_actions(actions_module).values())
