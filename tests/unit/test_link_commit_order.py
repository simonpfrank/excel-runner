"""Unit tests for `compute_link_commit_order` (docs/recalc_and_link_refresh_plan.md R5-R7):
the topological commit order over write-intent workbooks' R4 (absolute/UNC, to-be-modified)
links, with cycle (R6) and chain (R7) detection."""

import pytest

from excel_runner.core import ValidationError
from excel_runner.engine import compute_link_commit_order


class TestNoLinks:
    def test_workbooks_with_no_links_all_appear_in_the_order(self) -> None:
        order = compute_link_commit_order({"a": set(), "b": set(), "c": set()})

        assert set(order) == {"a", "b", "c"}
        assert len(order) == 3

    def test_empty_input_returns_an_empty_order(self) -> None:
        assert compute_link_commit_order({}) == []


class TestSimpleDependency:
    def test_a_links_to_b_so_b_is_committed_first(self) -> None:
        order = compute_link_commit_order({"a": {"b"}, "b": set()})

        assert order.index("b") < order.index("a")

    def test_order_is_correct_regardless_of_input_key_order(self) -> None:
        order = compute_link_commit_order({"b": set(), "a": {"b"}})

        assert order.index("b") < order.index("a")

    def test_multiple_independent_links_all_ordered_correctly(self) -> None:
        """a->b and c->d, two unrelated dependency pairs in the same commit."""
        order = compute_link_commit_order(
            {"a": {"b"}, "b": set(), "c": {"d"}, "d": set()}
        )

        assert order.index("b") < order.index("a")
        assert order.index("d") < order.index("c")

    def test_two_workbooks_link_to_the_same_target(self) -> None:
        order = compute_link_commit_order({"a": {"c"}, "b": {"c"}, "c": set()})

        assert order.index("c") < order.index("a")
        assert order.index("c") < order.index("b")


class TestCycleDetectionR6:
    def test_two_workbooks_linking_to_each_other_raises_validation_error(self) -> None:
        with pytest.raises(ValidationError, match="[Cc]yclical"):
            compute_link_commit_order({"a": {"b"}, "b": {"a"}})


class TestChainDetectionR7:
    def test_a_chain_of_three_raises_validation_error(self) -> None:
        """a->b->c: b is both a's target and c's source — a chain, not one-hop, per R7."""
        with pytest.raises(ValidationError, match="chain"):
            compute_link_commit_order({"a": {"b"}, "b": {"c"}, "c": set()})
