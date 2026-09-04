"""Unit tests for save-blocker inspection (docs/backend_eligibility_build_plan.md W1).

A "save blocker" is a named reason openpyxl must not be the thing that writes a given workbook
back to disk. Inspection is a plain zipfile/XML read — no Excel, no COM — so every test here
runs everywhere, including the one that builds a real link-bearing .xlsx by rewriting the zip
openpyxl produces (Excel is only needed to *author* a link, not to *carry* one).
"""

from pathlib import Path

import pytest

from excel_runner.core import SaveBlocker
from excel_runner.engine import inspect_save_blockers
from tests.unit.conftest import plain_workbook, workbook_with_external_link


class TestInspectSaveBlockers:
    def test_workbook_with_no_external_links_has_no_blockers(
        self, tmp_path: Path
    ) -> None:
        path = plain_workbook(tmp_path / "plain.xlsx")

        assert inspect_save_blockers(path) == frozenset()

    def test_workbook_with_an_outbound_external_link_is_blocked(
        self, tmp_path: Path
    ) -> None:
        path = workbook_with_external_link(tmp_path / "linking.xlsx")

        assert inspect_save_blockers(path) == frozenset(
            {SaveBlocker.OUTBOUND_EXTERNAL_LINKS}
        )

    def test_result_is_an_immutable_set_of_named_blockers(self, tmp_path: Path) -> None:
        """Not a bool and not a mutable dict — the set is designed to grow (plan sec 1.1)."""
        path = workbook_with_external_link(tmp_path / "linking.xlsx")

        blockers = inspect_save_blockers(path)

        assert isinstance(blockers, frozenset)
        assert all(isinstance(blocker, SaveBlocker) for blocker in blockers)

    def test_file_that_does_not_exist_has_no_blockers(self, tmp_path: Path) -> None:
        assert inspect_save_blockers(tmp_path / "never_created.xlsx") == frozenset()

    def test_file_that_is_not_a_valid_zip_has_no_blockers(self, tmp_path: Path) -> None:
        """A non-OOXML file can't be opened by openpyxl at all, so it can never reach a
        save — nothing to block, and inspection must not raise on it either."""
        path = tmp_path / "not_really.xlsx"
        path.write_bytes(b"this is not a zip archive")

        assert inspect_save_blockers(path) == frozenset()

    def test_directory_path_has_no_blockers(self, tmp_path: Path) -> None:
        directory = tmp_path / "a_folder.xlsx"
        directory.mkdir()

        assert inspect_save_blockers(directory) == frozenset()


class TestSaveBlockerEnum:
    def test_outbound_external_links_is_the_only_verified_member(self) -> None:
        """Members are added only when empirically verified (plan sec 1.1) — the charts rumour
        deliberately is not one."""
        assert [blocker.name for blocker in SaveBlocker] == ["OUTBOUND_EXTERNAL_LINKS"]

    @pytest.mark.parametrize("blocker", list(SaveBlocker))
    def test_every_member_has_a_stable_string_value_for_the_audit_log(
        self, blocker: SaveBlocker
    ) -> None:
        assert isinstance(blocker.value, str) and blocker.value
