"""Unit tests for per-value template resolution in excel_runner.core (Spec sec 2.2).

Design note: there is no whole-file "render as text" step (see docs/Specification.md sec 2.2
and the corrected rationale in docs/PRD.md sec 10.1) — resolution happens per field, either
once at load time (env-only context) or per-step during execution (env + steps context).
resolve_value is the one primitive both of those use.
"""

import pytest

from excel_runner.core import ValidationError, evaluate_condition, resolve_value


class TestResolveValueNonTemplated:
    def test_plain_string_is_unchanged(self) -> None:
        assert resolve_value("Reserving Data", {}) == "Reserving Data"

    def test_non_string_scalars_pass_through(self) -> None:
        assert resolve_value(5, {}) == 5
        assert resolve_value(True, {}) is True
        assert resolve_value(None, {}) is None
        assert resolve_value(3.5, {}) == 3.5


class TestResolveValueWholeExpression:
    """A field whose entire value is one {{ }} expression resolves to the native
    Python object, not a stringified render (Ansible-style native types, PRD sec 10.1)."""

    def test_returns_native_int(self) -> None:
        context = {"steps": {"find_row": {"output": {"row": 5}}}}
        result = resolve_value("{{ steps.find_row.output.row }}", context)
        assert result == 5
        assert isinstance(result, int)

    def test_returns_native_dict(self) -> None:
        context = {"steps": {"read_links": {"output": {"a.xlsx": "b.xlsx"}}}}
        result = resolve_value("{{ steps.read_links.output }}", context)
        assert result == {"a.xlsx": "b.xlsx"}

    def test_returns_native_bool(self) -> None:
        context = {"steps": {"copy_data": {"status": "success"}}}
        result = resolve_value("{{ steps.copy_data.status == 'success' }}", context)
        assert result is True

    def test_surrounding_whitespace_still_counts_as_whole_expression(self) -> None:
        context = {"steps": {"find_row": {"output": {"row": 5}}}}
        result = resolve_value("  {{ steps.find_row.output.row }}  ", context)
        assert result == 5


class TestResolveValueEmbedded:
    """An expression embedded inside a larger string always stringifies."""

    def test_embedded_expression_is_stringified(self) -> None:
        context = {"env": {"output_folder": "./output"}}
        result = resolve_value("{{ env.output_folder }}/summary.pdf", context)
        assert result == "./output/summary.pdf"

    def test_embedded_int_is_stringified(self) -> None:
        context = {"steps": {"find_row": {"output": {"row": 5}}}}
        result = resolve_value("row-{{ steps.find_row.output.row }}", context)
        assert result == "row-5"


class TestResolveValueRecursion:
    def test_resolves_dict_values(self) -> None:
        context = {"env": {"name": "North"}}
        result = resolve_value({"B": "{{ env.name }}", "C": 1200}, context)
        assert result == {"B": "North", "C": 1200}

    def test_resolves_dict_keys_computed_key_case(self) -> None:
        """The find_columns->write_row computed-key scenario from Spec sec 2.2 / PRD sec 11.19."""
        context = {"steps": {"find_key_columns": {"output": {"total": "D"}}}}
        result = resolve_value({"{{ steps.find_key_columns.output.total }}": "value"}, context)
        assert result == {"D": "value"}

    def test_resolves_list_items(self) -> None:
        context = {"env": {"a": "North", "b": "South"}}
        result = resolve_value(["{{ env.a }}", "{{ env.b }}", "East"], context)
        assert result == ["North", "South", "East"]

    def test_resolves_nested_structures(self) -> None:
        context = {"env": {"folder": "./output"}}
        result = resolve_value(
            {"filter": {"path": "{{ env.folder }}/x.xlsx", "values": ["{{ env.folder }}", 1]}},
            context,
        )
        assert result == {"filter": {"path": "./output/x.xlsx", "values": ["./output", 1]}}


class TestResolveValueUndefined:
    def test_undefined_reference_raises_validation_error(self) -> None:
        with pytest.raises(ValidationError) as exc_info:
            resolve_value("{{ steps.nonexistent.output }}", {"steps": {}})
        assert "nonexistent" in exc_info.value.detail.message

    def test_undefined_reference_keeps_technical_reason_separate(self) -> None:
        with pytest.raises(ValidationError) as exc_info:
            resolve_value("{{ steps.nonexistent.output }}", {"steps": {}})
        assert exc_info.value.detail.technical_reason

    def test_undefined_reference_embedded_in_a_larger_string_also_raises(self) -> None:
        with pytest.raises(ValidationError) as exc_info:
            resolve_value("prefix-{{ steps.nonexistent.output }}-suffix", {"steps": {}})
        assert "nonexistent" in exc_info.value.detail.message

    def test_syntactically_invalid_expression_raises_validation_error(self) -> None:
        with pytest.raises(ValidationError) as exc_info:
            resolve_value("{{ steps. }}", {"steps": {}})
        assert "not a valid template expression" in exc_info.value.detail.message


class TestResolveValueMultipleBlocksIsNotAWholeExpression:
    """Two adjacent {{ }} blocks in one string are a partial/embedded render (concatenated
    to a string), not a single whole-expression native-type resolution."""

    def test_two_adjacent_expressions_render_as_concatenated_string(self) -> None:
        context = {"env": {"a": "X", "b": "Y"}}
        result = resolve_value("{{ env.a }}{{ env.b }}", context)
        assert result == "XY"
        assert isinstance(result, str)


class TestEvaluateCondition:
    def test_wrapped_expression_true(self) -> None:
        context = {"steps": {"copy_data": {"status": "success"}}}
        assert evaluate_condition("{{ steps.copy_data.status == 'success' }}", context) is True

    def test_wrapped_expression_false(self) -> None:
        context = {"steps": {"copy_data": {"status": "error"}}}
        assert evaluate_condition("{{ steps.copy_data.status == 'success' }}", context) is False

    def test_bare_expression_without_braces(self) -> None:
        context = {"steps": {"copy_data": {"status": "success"}}}
        assert evaluate_condition("steps.copy_data.status == 'success'", context) is True

    def test_truthy_coercion_of_non_boolean_result(self) -> None:
        context = {"steps": {"find_row": {"output": {"row": 5}}}}
        assert evaluate_condition("{{ steps.find_row.output.row }}", context) is True

    def test_falsy_coercion_of_zero(self) -> None:
        context = {"steps": {"aggregate": {"output": {"North": 0}}}}
        assert evaluate_condition("{{ steps.aggregate.output.North }}", context) is False
