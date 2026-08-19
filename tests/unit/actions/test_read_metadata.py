"""Unit tests for the `read_metadata` action (PRD sec 7 — properties/cells sub-cases only;
the textbox sub-case is COM-only and deferred, per Spec sec 4/8)."""

import pytest

from excel_runner.actions import read_metadata as read_metadata_action
from excel_runner.core import ACTION_CAPABILITIES, ActionExecutionError, WorkbookSession


class TestReadMetadataAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["read_metadata"] == "file"

    def test_properties_target(self, richer_file_session: WorkbookSession) -> None:
        result = read_metadata_action(session=richer_file_session, target="properties")
        assert result.status == "success"
        assert result.output["title"] == "Q1 Report"
        assert result.output["creator"] == "Simon"

    def test_cells_target(self, richer_file_session: WorkbookSession) -> None:
        result = read_metadata_action(
            session=richer_file_session, target="cells", sheet="Summary", cells=["A1", "B3"]
        )
        assert result.status == "success"
        assert result.output == {"A1": "Notes", "B3": 100}

    def test_cells_target_without_sheet_or_cells_raises_a_clear_error(
        self, richer_file_session: WorkbookSession
    ) -> None:
        with pytest.raises(ActionExecutionError):
            read_metadata_action(session=richer_file_session, target="cells")

    def test_unsupported_target_raises_explicitly_rather_than_silently_acting_as_cells(
        self, richer_file_session: WorkbookSession
    ) -> None:
        """Regression test: target="textboxes" used to silently fall through to the "cells"
        handling (anything not "properties" was treated as "cells"). mypy now catches this at
        a literal call site like this one (file_action/com_action switched to ParamSpec so the
        decorator no longer erases parameter types, Spec sec 4/5.1) — the `type: ignore` below
        is deliberate, simulating the runner's actual dispatch, which calls every action via
        `**kwargs` unpacked from a dynamically-typed dict that no amount of ParamSpec can check.
        The runtime guard this test verifies is the only real defense on that path."""
        with pytest.raises(ActionExecutionError) as exc_info:
            read_metadata_action(
                session=richer_file_session,
                target="textboxes",  # type: ignore[arg-type]
                sheet="Summary",
                cells=["A1"],
            )
        assert "textboxes" in exc_info.value.detail.message
