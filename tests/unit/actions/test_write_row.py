"""Unit tests for the `write_row` action — base (column-mapping) and positional modes only.

The by-header mode (`values_by_header` + `headers_from`) needs another step's output, which
doesn't exist as a concept until runner.py threads step-output context through (Spec sec 4/8)
— deferred, not built here.
"""

import pytest

from excel_runner.actions import write_row as write_row_action
from excel_runner.core import ACTION_CAPABILITIES, ActionExecutionError, WorkbookSession


class TestWriteRowActionBaseMode:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["write_row"] == "file"

    def test_writes_by_explicit_column_mapping(self, file_session: WorkbookSession) -> None:
        result = write_row_action(
            session=file_session, sheet="Summary", row=5, values={"B": "North", "C": 1200, "D": "PASS"}
        )
        assert result.status == "success"
        assert file_session.handle["Summary"]["B5"].value == "North"
        assert file_session.handle["Summary"]["C5"].value == 1200
        assert file_session.handle["Summary"]["D5"].value == "PASS"


class TestWriteRowActionPositionalMode:
    def test_writes_values_in_order_from_a_start_column(self, file_session: WorkbookSession) -> None:
        result = write_row_action(
            session=file_session, sheet="Summary", row=5, values=["North", 1200, "PASS"], start_column="B"
        )
        assert result.status == "success"
        assert file_session.handle["Summary"]["B5"].value == "North"
        assert file_session.handle["Summary"]["C5"].value == 1200
        assert file_session.handle["Summary"]["D5"].value == "PASS"

    def test_positional_list_without_start_column_raises_a_clear_error(
        self, file_session: WorkbookSession
    ) -> None:
        with pytest.raises(ActionExecutionError):
            write_row_action(session=file_session, sheet="Summary", row=5, values=["North", 1200])
