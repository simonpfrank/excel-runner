"""Unit tests for error types in excel_runner.core (Spec sec 2.3)."""

import dataclasses

import pytest

from excel_runner.core import (
    ActionExecutionError,
    ErrorDetail,
    ExcelRunnerError,
    ValidationError,
)


class TestErrorDetail:
    def test_defaults(self) -> None:
        detail = ErrorDetail(message="field is wrong", technical_reason="KeyError: 'headers'")
        assert detail.message == "field is wrong"
        assert detail.technical_reason == "KeyError: 'headers'"
        assert detail.field is None
        assert detail.suggestion is None

    def test_explicit_fields(self) -> None:
        detail = ErrorDetail(
            message='field "headers" must be a list',
            technical_reason="TypeError: expected list, got str",
            field="headers",
            suggestion="Wrap it in [ ].",
        )
        assert detail.field == "headers"
        assert detail.suggestion == "Wrap it in [ ]."

    def test_is_frozen(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        with pytest.raises(dataclasses.FrozenInstanceError):
            detail.message = "changed"  # type: ignore[misc]


class TestExcelRunnerError:
    def test_carries_detail(self) -> None:
        detail = ErrorDetail(message="something broke", technical_reason="ValueError: bad range")
        error = ExcelRunnerError(detail)
        assert error.detail is detail

    def test_str_is_the_plain_english_message_not_the_technical_reason(self) -> None:
        detail = ErrorDetail(message="something broke", technical_reason="ValueError: bad range")
        error = ExcelRunnerError(detail)
        assert str(error) == "something broke"
        assert "ValueError" not in str(error)

    def test_is_an_exception(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        with pytest.raises(ExcelRunnerError):
            raise ExcelRunnerError(detail)


class TestErrorHierarchy:
    def test_validation_error_is_an_excel_runner_error(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        error = ValidationError(detail)
        assert isinstance(error, ExcelRunnerError)

    def test_action_execution_error_is_an_excel_runner_error(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        error = ActionExecutionError(detail)
        assert isinstance(error, ExcelRunnerError)

    def test_validation_error_and_action_execution_error_are_distinct(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        assert not isinstance(ValidationError(detail), ActionExecutionError)
        assert not isinstance(ActionExecutionError(detail), ValidationError)

    def test_catching_base_class_catches_subclasses(self) -> None:
        detail = ErrorDetail(message="m", technical_reason="t")
        try:
            raise ValidationError(detail)
        except ExcelRunnerError as caught:
            assert caught.detail is detail
        else:
            pytest.fail("ExcelRunnerError should have caught ValidationError")
