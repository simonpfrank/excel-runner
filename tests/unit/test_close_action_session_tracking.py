"""Unit tests for the close action's effect on SessionManager's own tracking
(excel_runner.runner._dispatch).

Bug this pins down: `actions.close` closes the handle directly — it has no reference to the
SessionManager, by design (actions stay decoupled from session lifecycle, Spec sec 4). Without
`_dispatch` telling the manager afterwards, `close_all()` at end of run still iterates every
session it has ever opened, including ones already explicitly closed, and tries to close each
one again. Against the file backend that's a harmless no-op; against a live Excel (xlw)
session the underlying COM object is already disconnected, and the second close raises —
aborting a run in which every real step already succeeded (reported: `com_error
(-2147417848, 'The object invoked has disconnected from its clients.')` surfacing as an
`ExceptionGroup` from `close_all()`).

Built with a fake handle rather than live Excel: what's under test is `_dispatch`'s own
control flow (does it call `forget_session` at the right moment), not backend behaviour —
`test_session_manager.py::TestForgetSession` already covers `forget_session` itself, and
`tests/unit/test_backends_xlw.py`/`test_owned_instance_registry.py` cover real COM sessions.
"""

from pathlib import Path

from excel_runner import core, engine
from excel_runner.core import ActionResult, Step, WorkbookRef
from excel_runner.engine import ScratchManager, SessionManager
from excel_runner.runner import _dispatch


class _ClosesOnceThenExplodes:
    def __init__(self) -> None:
        self.close_count = 0

    def close(self) -> None:
        self.close_count += 1
        if self.close_count > 1:
            raise RuntimeError("second close — simulated disconnected COM object")


def _registry_with_close_only() -> dict[str, engine.ActionSpec]:
    def _close(session: core.WorkbookSession) -> ActionResult:
        session.handle.close()
        return ActionResult(status="success", output={})

    return {
        "close": engine.ActionSpec(
            name="close",
            fn=_close,
            capability="file",
            description="fake close for this test",
            param_schema={"properties": {}, "required": []},
            writes=False,
        )
    }


def test_dispatching_close_stops_close_all_from_closing_the_session_again(
    tmp_path: Path,
) -> None:
    workbooks = {"manip": WorkbookRef(name="manip", file=str(tmp_path / "manip.xlsx"))}
    manager = SessionManager(workbooks, ScratchManager(tmp_path / "working"))
    handle = _ClosesOnceThenExplodes()
    manager._sessions["manip"] = core.WorkbookSession(
        name="manip", backend="file", handle=handle, path="manip.xlsx", mode="read_write"
    )
    step = Step(id="close_it", action="close", params={"workbook": "manip"})
    plan = engine.ExecutionPlan(modes={"manip": "read_write"})

    result = _dispatch(
        step, {"env": {}, "steps": {}}, _registry_with_close_only(), manager, plan
    )

    assert result.status == "success"
    assert handle.close_count == 1
    manager.close_all()  # must not raise — proves the session was forgotten, not re-closed
    assert handle.close_count == 1  # and definitely not re-closed


def test_a_close_action_that_fails_is_not_forgotten(tmp_path: Path) -> None:
    """If `close` itself reported an error, the handle was never actually confirmed closed —
    forgetting it anyway would let `close_all()` skip a session that may still need closing."""

    def _failing_close(session: core.WorkbookSession) -> ActionResult:
        from excel_runner.core import ErrorDetail

        return ActionResult(
            status="error",
            output={},
            error=ErrorDetail(message="could not close", technical_reason="simulated"),
        )

    workbooks = {"manip": WorkbookRef(name="manip", file=str(tmp_path / "manip.xlsx"))}
    manager = SessionManager(workbooks, ScratchManager(tmp_path / "working"))
    manager._sessions["manip"] = core.WorkbookSession(
        name="manip",
        backend="file",
        handle=_ClosesOnceThenExplodes(),
        path="manip.xlsx",
        mode="read_write",
    )
    registry = {
        "close": engine.ActionSpec(
            name="close",
            fn=_failing_close,
            capability="file",
            description="fake failing close for this test",
            param_schema={"properties": {}, "required": []},
            writes=False,
        )
    }
    step = Step(id="close_it", action="close", params={"workbook": "manip"})
    plan = engine.ExecutionPlan(modes={"manip": "read_write"})

    result = _dispatch(step, {"env": {}, "steps": {}}, registry, manager, plan)

    assert result.status == "error"
    assert "manip" in manager._sessions
