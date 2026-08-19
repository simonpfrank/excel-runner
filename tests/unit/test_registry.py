"""Unit tests for action discovery in excel_runner.engine (Spec sec 5.1)."""

import types

from excel_runner import actions
from excel_runner.core import ActionResult, WorkbookSession, file_action
from excel_runner.engine import ActionSpec, discover_actions


class TestDiscoverActions:
    def test_finds_every_capability_tagged_function(self) -> None:
        registry = discover_actions(actions)
        assert set(registry) == {
            "open",
            "save",
            "close",
            "copy",
            "read_range",
            "read_metadata",
            "write_cell",
            "write_range",
            "write_row",
            "insert_range",
            "set_column_width",
            "find_headers_row",
            "find_row",
            "find_column",
            "find_columns",
        }

    def test_each_entry_is_an_action_spec(self) -> None:
        registry = discover_actions(actions)
        assert isinstance(registry["open"], ActionSpec)

    def test_action_spec_carries_the_real_callable(self) -> None:
        registry = discover_actions(actions)
        assert registry["read_range"].fn is actions.read_range

    def test_action_spec_carries_its_capability(self) -> None:
        registry = discover_actions(actions)
        assert registry["open"].capability == "file"

    def test_action_spec_name_matches_the_function_name(self) -> None:
        registry = discover_actions(actions)
        assert registry["write_cell"].name == "write_cell"


class TestParamSchema:
    def test_excludes_the_session_parameter(self) -> None:
        registry = discover_actions(actions)
        assert "session" not in registry["read_range"].param_schema["properties"]

    def test_includes_the_other_parameters(self) -> None:
        registry = discover_actions(actions)
        schema = registry["read_range"].param_schema
        assert set(schema["properties"]) == {"sheet", "range"}

    def test_required_parameters_have_no_default(self) -> None:
        registry = discover_actions(actions)
        schema = registry["read_range"].param_schema
        assert set(schema["required"]) == {"sheet", "range"}

    def test_action_with_no_extra_params_has_an_empty_schema(self) -> None:
        registry = discover_actions(actions)
        schema = registry["open"].param_schema
        assert schema["properties"] == {}
        assert schema["required"] == []

    def test_optional_parameters_with_a_default_are_not_marked_required(self) -> None:
        """None of the 5 actions built so far has an optional param — exercised directly
        against a local fixture function instead of waiting for one to exist."""

        @file_action
        def _example_action_with_optional_param(
            session: WorkbookSession, sheet: str, as_: str = "values"
        ) -> ActionResult:
            return ActionResult(status="success", output={})

        module = types.SimpleNamespace(**{"_example_action_with_optional_param": _example_action_with_optional_param})
        registry = discover_actions(module)  # type: ignore[arg-type]
        schema = registry["_example_action_with_optional_param"].param_schema
        assert schema["required"] == ["sheet"]
        assert "as_" in schema["properties"]
        assert "as_" not in schema["required"]
