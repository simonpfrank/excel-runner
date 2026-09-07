"""Unit tests for AuditLogger (Spec sec 6.2)."""

import json
from pathlib import Path

from excel_runner.core import ErrorDetail, Step
from excel_runner.runner import AuditLogger, StepResult


class TestAuditLogger:
    def test_writes_one_json_line_per_step(self, tmp_path: Path) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)
        step = Step(
            id="s1",
            action="read_range",
            params={"workbook": "manip", "sheet": "S", "range": "A1"},
        )
        result = StepResult(step_id="s1", status="success", output={"values": "x"})

        logger.record_step(
            step,
            result,
            started_at="2026-08-19T10:00:00",
            ended_at="2026-08-19T10:00:01",
        )

        lines = log_path.read_text().splitlines()
        assert len(lines) == 1
        record = json.loads(lines[0])
        assert record["step_id"] == "s1"
        assert record["action"] == "read_range"
        assert record["status"] == "success"
        assert record["output"] == {"values": "x"}

    def test_multiple_steps_append_rather_than_overwrite(self, tmp_path: Path) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)
        step1 = Step(id="s1", action="open", params={"workbook": "manip"})
        step2 = Step(id="s2", action="close", params={"workbook": "manip"})

        logger.record_step(
            step1, StepResult(step_id="s1", status="success", output={}), "t0", "t1"
        )
        logger.record_step(
            step2, StepResult(step_id="s2", status="success", output={}), "t1", "t2"
        )

        lines = log_path.read_text().splitlines()
        assert len(lines) == 2
        assert json.loads(lines[0])["step_id"] == "s1"
        assert json.loads(lines[1])["step_id"] == "s2"

    def test_error_detail_is_serialized(self, tmp_path: Path) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)
        step = Step(
            id="s1",
            action="find_row",
            params={
                "workbook": "manip",
                "sheet": "S",
                "column": "A",
                "search_value": "x",
            },
        )
        result = StepResult(
            step_id="s1",
            status="error",
            output={},
            error=ErrorDetail(
                message="not found", technical_reason="find_row: no matching row"
            ),
        )

        logger.record_step(step, result, "t0", "t1")

        record = json.loads(log_path.read_text().splitlines()[0])
        assert record["error"]["message"] == "not found"
        assert record["error"]["technical_reason"] == "find_row: no matching row"

    def test_skipped_step_is_recorded(self, tmp_path: Path) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)
        step = Step(
            id="s1",
            action="close",
            params={"workbook": "manip"},
            if_expr="{{ steps.x.status == 'success' }}",
        )
        result = StepResult(step_id="s1", status="skipped", output={})

        logger.record_step(step, result, "t0", "t1")

        record = json.loads(log_path.read_text().splitlines()[0])
        assert record["status"] == "skipped"

    def test_creates_parent_directories(self, tmp_path: Path) -> None:
        log_path = tmp_path / "nested" / "dir" / "audit.jsonl"
        logger = AuditLogger(log_path)
        step = Step(id="s1", action="open", params={"workbook": "manip"})

        logger.record_step(
            step, StepResult(step_id="s1", status="success", output={}), "t0", "t1"
        )

        assert log_path.exists()


class TestRecordEvent:
    """Run-level events (docs/backend_eligibility_build_plan.md W8). Backend routing is
    deliberately invisible in the workflow YAML, so the audit log is the only place a user can
    find out *why* a run needed Excel and was therefore slow.
    """

    def test_writes_a_record_carrying_the_event_name_and_its_detail(
        self, tmp_path: Path
    ) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)

        logger.record_event(
            "workbook_opened",
            {
                "workbook": "manip",
                "backend": "xlw",
                "save_blockers": ["outbound_external_links"],
            },
        )

        record = json.loads(log_path.read_text().splitlines()[0])
        assert record["event"] == "workbook_opened"
        assert record["workbook"] == "manip"
        assert record["backend"] == "xlw"
        assert record["save_blockers"] == ["outbound_external_links"]
        assert record["at"]

    def test_an_event_is_distinguishable_from_a_step_record(
        self, tmp_path: Path
    ) -> None:
        """A reader tells them apart by key, not by position — events interleave with steps."""
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)

        logger.record_event("workbook_opened", {"workbook": "manip"})
        logger.record_step(
            Step(id="s1", action="open", params={"workbook": "manip"}),
            StepResult(step_id="s1", status="success", output={}),
            "t0",
            "t1",
        )

        event, step = (json.loads(line) for line in log_path.read_text().splitlines())
        assert "event" in event and "step_id" not in event
        assert "step_id" in step and "event" not in step

    def test_events_append_alongside_steps_rather_than_overwriting(
        self, tmp_path: Path
    ) -> None:
        log_path = tmp_path / "audit.jsonl"
        logger = AuditLogger(log_path)

        logger.record_event("workbook_opened", {"workbook": "a"})
        logger.record_event("backend_switched", {"workbook": "a", "to": "xlw"})

        lines = log_path.read_text().splitlines()
        assert [json.loads(line)["event"] for line in lines] == [
            "workbook_opened",
            "backend_switched",
        ]
