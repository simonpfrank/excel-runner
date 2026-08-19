"""Unit tests for execution-time types in excel_runner.core: ActionResult, WorkbookSession,
and the capability-tagging decorators actions use to register with the discovery mechanism
(Spec sec 5.1's ActionSpec.capability, sec 4's ActionResult)."""

import dataclasses

import pytest

from excel_runner.core import (
    ACTION_CAPABILITIES,
    ActionResult,
    WorkbookSession,
    com_action,
    file_action,
)


class TestActionResult:
    def test_success_defaults(self) -> None:
        result = ActionResult(status="success", output={"values": [[1, 2]]})
        assert result.status == "success"
        assert result.output == {"values": [[1, 2]]}
        assert result.error is None

    def test_is_frozen(self) -> None:
        result = ActionResult(status="success", output={})
        with pytest.raises(dataclasses.FrozenInstanceError):
            result.status = "error"  # type: ignore[misc]


class TestWorkbookSession:
    def test_construction(self) -> None:
        session = WorkbookSession(
            name="manip",
            backend="file",
            handle=object(),
            path="./output/manip.xlsx",
            mode="read_write",
        )
        assert session.name == "manip"
        assert session.backend == "file"
        assert session.path == "./output/manip.xlsx"
        assert session.mode == "read_write"
        assert session.dirty is False
        assert session.scratch_path is None

    def test_dirty_is_mutable(self) -> None:
        """Unlike Workflow/Step, a session represents live run state — deliberately not
        frozen, matching the guideline to model state with a real class, not a raw dict."""
        session = WorkbookSession(
            name="manip", backend="file", handle=object(), path="x.xlsx", mode="read_write"
        )
        session.dirty = True
        assert session.dirty is True


class TestCapabilityTagging:
    def test_file_action_registers_file_capability(self) -> None:
        @file_action
        def _example_file_action(session: WorkbookSession) -> ActionResult:
            return ActionResult(status="success", output={})

        assert ACTION_CAPABILITIES["_example_file_action"] == "file"

    def test_com_action_registers_com_capability(self) -> None:
        @com_action
        def _example_com_action(session: WorkbookSession) -> ActionResult:
            return ActionResult(status="success", output={})

        assert ACTION_CAPABILITIES["_example_com_action"] == "com"

    def test_decorator_returns_the_original_function_unchanged(self) -> None:
        @file_action
        def _example_action(session: WorkbookSession) -> ActionResult:
            return ActionResult(status="success", output={"marker": True})

        result = _example_action(session=None)  # type: ignore[arg-type]  # this function ignores session
        assert result.output == {"marker": True}
