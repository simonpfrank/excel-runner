"""Unit tests for the `insert_range` action (PRD sec 7/sec 11 item 12)."""

from excel_runner import backends
from excel_runner.actions import insert_range as insert_range_action
from excel_runner.core import ACTION_CAPABILITIES, WorkbookSession


class TestInsertRangeAction:
    def test_registers_as_a_file_action(self) -> None:
        assert ACTION_CAPABILITIES["insert_range"] == "file"

    def test_inserts_a_whole_column_with_a_header(self, file_session: WorkbookSession) -> None:
        result = insert_range_action(
            session=file_session, sheet="Summary", at="B:B", header={"row": 1, "text": "Flag"}
        )
        assert result.status == "success"
        assert backends.read_range(file_session.handle, "Summary", "B1") == "Flag"

    def test_partial_range_returns_a_structured_error_not_a_raw_exception(
        self, file_session: WorkbookSession
    ) -> None:
        """Consistent with find_*'s "legitimately didn't work" pattern (Spec sec 4) — this is
        an anticipated, named limitation (PRD sec 11 item 12), not an unexpected failure, so
        it becomes a structured ActionResult, not an exception escaping to the caller."""
        result = insert_range_action(session=file_session, sheet="Summary", at="B2:B5", direction="rows")
        assert result.status == "error"
        assert result.error is not None
        assert "partial range" in result.error.message
