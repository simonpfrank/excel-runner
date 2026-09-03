"""Unit tests for the `read_metadata` action (PRD sec 7 — properties/cells sub-cases only;
the textbox sub-case is COM-only and deferred, per Spec sec 4/8)."""

import pytest

from excel_runner.actions import read_metadata as read_metadata_action
from excel_runner.core import ACTION_CAPABILITIES, ActionExecutionError, WorkbookSession
from tests.unit.conftest import requires_excel, requires_working_xlwings_save


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
            session=richer_file_session,
            target="cells",
            sheet="Summary",
            cells=["A1", "B3"],
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
        a literal call site like this one (the capability decorators switched to ParamSpec so
        they no longer erase parameter types, Spec sec 4/5.1) — the `type: ignore` below
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


class TestReadMetadataCellsNamedRange:
    def test_reads_a_cell_via_a_workbook_defined_name(
        self, richer_named_range_file_session: WorkbookSession
    ) -> None:
        result = read_metadata_action(
            session=richer_named_range_file_session,
            target="cells",
            sheet="Summary",
            cells=["NotesCell"],
        )
        assert result.status == "success"
        assert result.output == {"NotesCell": "Notes"}

    def test_neither_valid_a1_nor_a_real_defined_name_raises_a_clear_error(
        self, richer_named_range_file_session: WorkbookSession
    ) -> None:
        with pytest.raises(ActionExecutionError) as exc_info:
            read_metadata_action(
                session=richer_named_range_file_session,
                target="cells",
                sheet="Summary",
                cells=["NotARealRange"],
            )
        assert "NotARealRange" in exc_info.value.detail.message


@requires_excel
@requires_working_xlwings_save
class TestReadMetadataCellsFormulaParam:
    def test_defaults_to_the_computed_value_for_a_formula_cell(
        self, formula_file_session: WorkbookSession
    ) -> None:
        result = read_metadata_action(
            session=formula_file_session, target="cells", sheet="Summary", cells=["B1"]
        )
        assert result.output == {"B1": 20}

    def test_formula_true_returns_the_formula_text_instead(
        self, formula_file_session: WorkbookSession
    ) -> None:
        result = read_metadata_action(
            session=formula_file_session,
            target="cells",
            sheet="Summary",
            cells=["B1"],
            formula=True,
        )
        assert result.output == {"B1": "=A1*2"}
