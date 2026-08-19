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
