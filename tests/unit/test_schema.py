"""Unit tests for the workflow data model in excel_runner.core (Spec sec 2.1)."""

import dataclasses

import pytest

from excel_runner.core import Step, WorkbookRef, Workflow


class TestWorkbookRef:
    def test_defaults(self) -> None:
        ref = WorkbookRef(name="historical", file="./input/historical.xlsx")
        assert ref.name == "historical"
        assert ref.file == "./input/historical.xlsx"
        assert ref.create_if_missing is False
        assert ref.template is None

    def test_explicit_fields(self) -> None:
        ref = WorkbookRef(
            name="results",
            file="./output/results.xlsx",
            create_if_missing=True,
            template="historical",
        )
        assert ref.create_if_missing is True
        assert ref.template == "historical"

    def test_is_frozen(self) -> None:
        ref = WorkbookRef(name="historical", file="./input/historical.xlsx")
        with pytest.raises(dataclasses.FrozenInstanceError):
            ref.name = "renamed"  # type: ignore[misc]

    def test_equality(self) -> None:
        a = WorkbookRef(name="historical", file="./input/historical.xlsx")
        b = WorkbookRef(name="historical", file="./input/historical.xlsx")
        assert a == b


class TestStep:
    def test_defaults(self) -> None:
        step = Step(id="get_totals", action="read_range", params={"workbook": "manip"})
        assert step.id == "get_totals"
        assert step.action == "read_range"
        assert step.params == {"workbook": "manip"}
        assert step.if_expr is None

    def test_explicit_if_expr(self) -> None:
        step = Step(
            id="recalc",
            action="recalculate",
            params={"workbook": "manip"},
            if_expr="{{ steps.copy_data.status == 'success' }}",
        )
        assert step.if_expr == "{{ steps.copy_data.status == 'success' }}"

    def test_is_frozen(self) -> None:
        step = Step(id="get_totals", action="read_range", params={})
        with pytest.raises(dataclasses.FrozenInstanceError):
            step.action = "write_cell"  # type: ignore[misc]

    def test_params_dict_stays_mutable_by_the_caller(self) -> None:
        """Step.params is a plain dict — the dataclass itself is frozen, but its dict
        field is not deep-frozen. Documents the actual (permissive) behavior rather
        than assuming deep immutability that isn't implemented."""
        step = Step(id="s1", action="write_cell", params={"value": 1})
        step.params["value"] = 2
        assert step.params["value"] == 2


class TestWorkflow:
    def test_construction(self) -> None:
        workbooks = {"manip": WorkbookRef(name="manip", file="./output/manip.xlsx")}
        steps = (Step(id="s1", action="open", params={"workbook": "manip"}),)
        workflow = Workflow(env={"output_folder": "./output"}, workbooks=workbooks, steps=steps)

        assert workflow.env == {"output_folder": "./output"}
        assert workflow.workbooks == workbooks
        assert workflow.steps == steps

    def test_is_frozen(self) -> None:
        workflow = Workflow(env={}, workbooks={}, steps=())
        with pytest.raises(dataclasses.FrozenInstanceError):
            workflow.env = {"a": 1}  # type: ignore[misc]

    def test_steps_is_a_tuple(self) -> None:
        workflow = Workflow(env={}, workbooks={}, steps=())
        assert isinstance(workflow.steps, tuple)
