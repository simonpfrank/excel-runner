"""Unit tests for both validation tiers (Spec sec 5.4).

Tier 1 (static schema): no workbook access, structural checks only.
Tier 2 (dry-run / step-graph): still no workbook access, reasons over the whole step list.
"""

import pytest

from excel_runner import actions as actions_module
from excel_runner import engine as validation
from excel_runner.core import Step, ValidationError, WorkbookRef, Workflow
from excel_runner.engine import discover_actions

_REGISTRY = discover_actions(actions_module)


def _workflow(
    steps: list[Step], workbooks: dict[str, WorkbookRef] | None = None
) -> Workflow:
    return Workflow(env={}, workbooks=workbooks or {}, steps=tuple(steps))


class TestActionExists:
    def test_passes_for_a_known_action(self) -> None:
        workflow = _workflow(
            [Step(id="s1", action="open", params={"workbook": "manip"})]
        )
        validation.validate_static(workflow, _REGISTRY)  # should not raise

    def test_unknown_action_raises_with_a_suggestion(self) -> None:
        workflow = _workflow(
            [Step(id="s1", action="opne", params={"workbook": "manip"})]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "opne" in exc_info.value.detail.message
        assert "open" in exc_info.value.detail.message

    def test_unknown_action_with_no_close_match_omits_a_suggestion(self) -> None:
        workflow = _workflow([Step(id="s1", action="zzzzzzzzzz", params={})])
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "zzzzzzzzzz" in exc_info.value.detail.message
        assert "mean" not in exc_info.value.detail.message


class TestUnknownParams:
    def test_extra_param_not_in_the_action_schema_raises(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "range": "A1",
                        "bogus": 1,
                    },
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "bogus" in exc_info.value.detail.message

    def test_copy_is_exempt_from_schema_shape_checks(self) -> None:
        """copy's raw YAML shape (source/target dicts) doesn't match its Python signature yet
        — that translation is the runner's job (Spec sec 4/8 item 7), not built. Must not
        false-positive as "unknown params" until then."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="copy",
                    params={
                        "source": {
                            "workbook": "historical",
                            "sheet": "S",
                            "range": "A1",
                        },
                        "target": {"workbook": "manip", "sheet": "S", "range": "B1"},
                    },
                )
            ]
        )
        validation.validate_static(workflow, _REGISTRY)  # should not raise


class TestRequiredParams:
    def test_missing_required_param_raises(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S"},
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "range" in exc_info.value.detail.message

    def test_stop_is_exempt_from_the_implicit_workbook_requirement(self) -> None:
        """Every other action implicitly requires workbook: (_IMPLICIT_FIELDS) — stop has no
        workbook at all (PRD sec 6.9), so it must be exempt the same way copy is."""
        workflow = _workflow([Step(id="s1", action="stop", params={})])
        validation.validate_static(workflow, _REGISTRY)  # should not raise

    def test_missing_implicit_workbook_field_raises(self) -> None:
        workflow = _workflow(
            [Step(id="s1", action="read_range", params={"sheet": "S", "range": "A1"})]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "workbook" in exc_info.value.detail.message


class TestParamTypes:
    def test_field_that_should_be_a_list_but_is_a_string_raises_with_a_wrap_suggestion(
        self,
    ) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="find_headers_row",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "search_range": "A1:C4",
                        "patterns": "Region",
                    },
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "patterns" in exc_info.value.detail.message
        suggestion = exc_info.value.detail.suggestion
        assert suggestion is not None
        assert "[" in suggestion

    def test_correct_type_passes(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="find_headers_row",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "search_range": "A1:C4",
                        "patterns": ["Region"],
                    },
                )
            ]
        )
        validation.validate_static(workflow, _REGISTRY)  # should not raise

    def test_literal_type_rejects_a_value_outside_the_allowed_set(self) -> None:
        """read_metadata's `target` is Literal["properties", "cells"] — exercises the
        Literal-specific branches of _matches_type/_type_name, not just plain isinstance.
        """
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_metadata",
                    params={"workbook": "manip", "target": "bogus"},
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "target" in exc_info.value.detail.message

    def test_union_type_rejects_a_value_matching_neither_branch(self) -> None:
        """write_row's `values` is dict[str, Any] | list[Any] — exercises the Union-specific
        branches of _matches_type/_type_name/_expects_list."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="write_row",
                    params={"workbook": "manip", "sheet": "S", "row": 1, "values": 5},
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "values" in exc_info.value.detail.message
        assert "dict" in exc_info.value.detail.message
        assert "list" in exc_info.value.detail.message

    def test_type_name_fallback_for_a_plain_type(self) -> None:
        """read_range's `sheet` is a plain `str` — exercises _type_name's final fallback for
        a non-generic, non-Union, non-Literal annotation."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": 5, "range": "A1"},
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "must be a str" in exc_info.value.detail.message


class TestStepReferences:
    def test_reference_to_a_nonexistent_step_raises_with_a_suggestion(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="get_regional",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                ),
                Step(
                    id="totals",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "B1",
                        "value": "{{ steps.get_regional_data.output }}",
                    },
                ),
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "get_regional_data" in exc_info.value.detail.message
        assert "get_regional" in exc_info.value.detail.message  # the suggestion

    def test_reference_to_a_later_step_raises(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "B1",
                        "value": "{{ steps.s2.output }}",
                    },
                ),
                Step(
                    id="s2",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                ),
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "s2" in exc_info.value.detail.message

    def test_reference_to_an_earlier_step_passes(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                ),
                Step(
                    id="s2",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "B1",
                        "value": "{{ steps.s1.output }}",
                    },
                ),
            ]
        )
        validation.validate_static(workflow, _REGISTRY)  # should not raise

    def test_non_string_param_values_are_ignored_not_scanned_for_references(
        self,
    ) -> None:
        """int/bool/etc. param values can't contain a {{ steps.x }} reference — confirms the
        scanner just skips them rather than erroring, while still finding a real reference
        elsewhere in the same step."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                ),
                Step(
                    id="s2",
                    action="write_row",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "row": 5,
                        "values": {"B": "{{ steps.s1.output }}", "C": 1200},
                    },
                ),
            ]
        )
        validation.validate_static(workflow, _REGISTRY)  # should not raise

    def test_reference_inside_if_expr_is_checked_too(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="close",
                    params={"workbook": "manip"},
                    if_expr="{{ steps.nonexistent.status == 'success' }}",
                )
            ]
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_static(workflow, _REGISTRY)
        assert "nonexistent" in exc_info.value.detail.message


class TestDryRunWorkbooksDeclared:
    def test_workbook_not_in_registry_raises(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "nope", "sheet": "S", "range": "A1"},
                )
            ],
            workbooks={"manip": WorkbookRef(name="manip", file="manip.xlsx")},
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.plan(workflow, _REGISTRY)
        assert "nope" in exc_info.value.detail.message

    def test_all_declared_workbooks_passes(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                )
            ],
            workbooks={"manip": WorkbookRef(name="manip", file="manip.xlsx")},
        )
        validation.plan(workflow, _REGISTRY)  # should not raise


class TestExecutionPlanModeInference:
    def test_workbook_never_written_to_is_read_only(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                )
            ],
            workbooks={"manip": WorkbookRef(name="manip", file="manip.xlsx")},
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert plan.modes["manip"] == "read_only"

    def test_workbook_written_to_is_read_write(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "A1",
                        "value": 1,
                    },
                )
            ],
            workbooks={"manip": WorkbookRef(name="manip", file="manip.xlsx")},
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert plan.modes["manip"] == "read_write"

    def test_a_workbook_only_referenced_by_a_read_action_in_some_steps_and_write_in_others_is_read_write(
        self,
    ) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "manip", "sheet": "S", "range": "A1"},
                ),
                Step(
                    id="s2",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "B1",
                        "value": 1,
                    },
                ),
            ],
            workbooks={"manip": WorkbookRef(name="manip", file="manip.xlsx")},
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert plan.modes["manip"] == "read_write"

    def test_copy_marks_every_workbook_it_touches_as_read_write(self) -> None:
        """Can't statically tell copy's source from its target without the runner's
        translation layer (Spec sec 4) — the safe fallback (PRD sec 6.3) is read_write for
        both rather than silently under-provisioning one of them."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="copy",
                    params={
                        "source": {
                            "workbook": "historical",
                            "sheet": "S",
                            "range": "A1",
                        },
                        "target": {"workbook": "manip", "sheet": "S", "range": "B1"},
                    },
                )
            ],
            workbooks={
                "historical": WorkbookRef(name="historical", file="historical.xlsx"),
                "manip": WorkbookRef(name="manip", file="manip.xlsx"),
            },
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert plan.modes["historical"] == "read_write"
        assert plan.modes["manip"] == "read_write"

    def test_workbook_references_nested_inside_a_list_are_still_found(self) -> None:
        """The workbook-name walk is generic — it recurses into lists too, not just dicts —
        so a hypothetical future action with a list of workbook-referencing sub-objects
        (not something any current action actually has) still gets planned correctly."""
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="write_cell",
                    params={
                        "workbook": "manip",
                        "sheet": "S",
                        "cell": "A1",
                        "value": 1,
                        "targets": [
                            {"workbook": "historical"},
                            {"workbook": "results"},
                        ],
                    },
                )
            ],
            workbooks={
                "manip": WorkbookRef(name="manip", file="manip.xlsx"),
                "historical": WorkbookRef(name="historical", file="historical.xlsx"),
                "results": WorkbookRef(name="results", file="results.xlsx"),
            },
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert "historical" in plan.modes
        assert "results" in plan.modes

    def test_workbook_never_referenced_by_any_step_defaults_to_read_only(self) -> None:
        workflow = _workflow(
            [],
            workbooks={"unused": WorkbookRef(name="unused", file="unused.xlsx")},
        )
        plan = validation.plan(workflow, _REGISTRY)
        assert plan.modes["unused"] == "read_only"
