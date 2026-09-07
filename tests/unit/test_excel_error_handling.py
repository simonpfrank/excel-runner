"""Unit tests for the Excel-call error boundary (backends._excel_operation) and the CLI's
catch-all.

Why this file exists: every call into live Excel — `xlw_`-prefixed (portable xlwings) or
`com_`-prefixed (raw `.api`) — can fail with a `pywintypes.com_error` that means nothing to a
user. Before this boundary existed those propagated untouched through `runner.run_workflow`
(which deliberately has no try/except around step execution) and past `cli.main` (which only
caught `ExcelRunnerError`), so the user got a raw traceback instead of a structured error.

No live Excel is needed here: the boundary is exercised with functions that raise on purpose,
which is the point — it must behave identically whatever Excel threw.
"""

import logging
from unittest.mock import patch

import pytest

from excel_runner import backends
from excel_runner.cli import main
from excel_runner.core import ActionExecutionError, ErrorDetail, ExcelRunnerError


class _FakeComError(Exception):
    """Stands in for pywintypes.com_error, which can't be raised without Windows COM."""


class TestExcelOperationBoundary:
    def test_wraps_an_unexpected_excel_failure_in_a_structured_error(self) -> None:
        @backends._excel_operation("open workbook")
        def explode() -> None:
            raise _FakeComError("Open method of Workbooks class failed")

        with pytest.raises(ActionExecutionError) as caught:
            explode()

        detail = caught.value.detail
        assert "open workbook" in detail.message
        assert "Excel" in detail.message
        assert "Open method of Workbooks class failed" in detail.technical_reason
        assert "_FakeComError" in detail.technical_reason

    def test_original_exception_is_kept_as_the_cause(self) -> None:
        @backends._excel_operation("save workbook")
        def explode() -> None:
            raise _FakeComError("boom")

        with pytest.raises(ActionExecutionError) as caught:
            explode()

        assert isinstance(caught.value.__cause__, _FakeComError)

    def test_our_own_errors_pass_through_untouched(self) -> None:
        original = ActionExecutionError(
            ErrorDetail(message="already structured", technical_reason="deliberate")
        )

        @backends._excel_operation("write cell")
        def explode() -> None:
            raise original

        with pytest.raises(ExcelRunnerError) as caught:
            explode()

        assert caught.value is original

    @pytest.mark.parametrize(
        "already_precise",
        [
            FileNotFoundError("no such workbook"),
            TimeoutError("calculation too slow"),
            ValueError('A sheet named "Data" already exists.'),
            NotImplementedError('partial range "B2:C4" is not supported yet'),
        ],
    )
    def test_already_precise_errors_pass_through_untouched(
        self, already_precise: Exception
    ) -> None:
        """Two different reasons, same outcome. `xlw_open_workbook` documents
        FileNotFoundError and `com_wait_until_calculation_done` documents TimeoutError — both
        already say exactly what went wrong. The `ValueError`/`NotImplementedError` argument
        guards are shared verbatim by the file and xlw twins, so wrapping only the xlw side
        would let a user tell which backend they landed on from the error text, which the
        backend-parity contract (tests/unit/test_backend_primitive_contract.py) forbids."""

        @backends._excel_operation("open workbook")
        def explode() -> None:
            raise already_precise

        with pytest.raises(type(already_precise)) as caught:
            explode()

        assert caught.value is already_precise

    def test_a_successful_call_returns_its_value_unchanged(self) -> None:
        @backends._excel_operation("read range")
        def succeed(value: int, *, doubled: bool = False) -> int:
            return value * 2 if doubled else value

        assert succeed(21, doubled=True) == 42
        assert succeed(21) == 21

    def test_the_wrapped_function_keeps_its_identity(self) -> None:
        @backends._excel_operation("read range")
        def named_function() -> None:
            """Original docstring."""

        assert named_function.__name__ == "named_function"
        assert named_function.__doc__ == "Original docstring."


class TestEveryExcelCallIsGuarded:
    """The boundary is only worth anything if nothing slips past it.

    This walks the real module rather than a hand-maintained list, so a newly added `xlw_`/
    `com_` function that forgets the decorator fails here instead of surfacing as a raw
    traceback in front of a user.
    """

    def _excel_functions(self) -> list[str]:
        return sorted(
            name
            for name in dir(backends)
            if (name.startswith("xlw_") or name.startswith("com_"))
            and callable(getattr(backends, name))
        )

    def test_there_are_excel_functions_to_check(self) -> None:
        assert len(self._excel_functions()) > 20

    def test_every_excel_function_is_wrapped(self) -> None:
        unguarded = [
            name
            for name in self._excel_functions()
            if not getattr(getattr(backends, name), "__excel_operation__", None)
        ]

        assert unguarded == []


class TestCliCatchAll:
    def test_an_unexpected_exception_becomes_exit_1_not_a_traceback(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        with patch(
            "excel_runner.cli.run_workflow", side_effect=_FakeComError("raw COM failure")
        ):
            with caplog.at_level(logging.ERROR):
                exit_code = main(["some_workflow.yaml"])

        assert exit_code == 1

    def test_the_unexpected_exception_is_logged_for_diagnosis(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        with patch(
            "excel_runner.cli.run_workflow", side_effect=_FakeComError("raw COM failure")
        ):
            with caplog.at_level(logging.ERROR):
                main(["some_workflow.yaml"])

        logged = caplog.text
        assert "raw COM failure" in logged
        assert "_FakeComError" in logged
