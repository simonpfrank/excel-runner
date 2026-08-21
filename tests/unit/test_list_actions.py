"""Unit tests for list_actions() — the discovery entry point a future agent/CLI wrapper would
use to generate its own tool definitions (Spec sec 6.3)."""

from excel_runner.engine import ActionSpec
from excel_runner.runner import list_actions


class TestListActions:
    def test_returns_a_tuple_of_action_specs(self) -> None:
        result = list_actions()
        assert isinstance(result, tuple)
        assert all(isinstance(spec, ActionSpec) for spec in result)

    def test_includes_every_built_action(self) -> None:
        names = {spec.name for spec in list_actions()}
        assert "read_range" in names
        assert "write_cell" in names
        assert len(names) == 19

    def test_each_entry_has_a_description(self) -> None:
        specs = {spec.name: spec for spec in list_actions()}
        assert specs["read_range"].description == "Read a cell or range of cells."

    def test_is_consistent_with_discover_actions(self) -> None:
        """list_actions() should just be discover_actions() wired to the real actions module —
        not a second, possibly-drifting source of truth."""
        from excel_runner import actions as actions_module
        from excel_runner.engine import discover_actions

        expected = discover_actions(actions_module)
        actual = {spec.name: spec for spec in list_actions()}
        assert set(actual) == set(expected)
