"""Unit tests for the `parse_date` control action."""

from datetime import date

import pytest

from excel_runner.actions import parse_date
from excel_runner.core import ActionExecutionError


class TestParseDateAction:
    def test_parses_an_iso_date_to_a_native_date(self) -> None:
        result = parse_date(value="2026-06-30", format="%Y-%m-%d")

        assert result.output == {"value": date(2026, 6, 30)}
        assert result.status == "success"

    def test_parses_a_date_using_the_supplied_format(self) -> None:
        result = parse_date(value="30/06/2026", format="%d/%m/%Y")

        assert result.output == {"value": date(2026, 6, 30)}

    def test_rejects_a_value_that_does_not_match_the_format(self) -> None:
        with pytest.raises(ActionExecutionError, match="parse_date: could not parse"):
            parse_date(value="2026-06-30", format="%d/%m/%Y")

    def test_rejects_an_invalid_calendar_date(self) -> None:
        with pytest.raises(ActionExecutionError, match="parse_date: could not parse"):
            parse_date(value="2026-02-30", format="%Y-%m-%d")