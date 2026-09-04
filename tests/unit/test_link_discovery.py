"""Unit tests for the static, pure-Python external-link discovery pass in excel_runner.engine
(docs/recalc_and_link_refresh_plan.md sec 2): classifying and resolving link Target strings
requires no Excel/COM at all (real-Excel is only needed to *create* fixtures for
`scan_external_link_targets`, whose own implementation is pure zipfile/XML).
"""

from pathlib import Path

from excel_runner.backends import OwnedInstanceRegistry
from excel_runner.engine import (
    classify_link_target,
    resolve_link_target,
    scan_external_link_targets,
)
from tests.unit.conftest import requires_excel, requires_working_xlwings_save


class TestClassifyLinkTarget:
    def test_bare_filename_is_same_folder(self) -> None:
        assert classify_link_target("target.xlsx") == "same_folder"

    def test_forward_slash_subpath_is_relative_subpath(self) -> None:
        assert classify_link_target("other/target.xlsx") == "relative_subpath"

    def test_parent_relative_subpath_is_relative_subpath(self) -> None:
        assert classify_link_target("../target.xlsx") == "relative_subpath"

    def test_windows_drive_absolute_path_is_absolute(self) -> None:
        assert classify_link_target(r"C:\data\target.xlsx") == "absolute"

    def test_unc_path_is_absolute(self) -> None:
        assert classify_link_target(r"\\server\share\target.xlsx") == "absolute"

    def test_file_uri_is_absolute(self) -> None:
        assert classify_link_target("file:///C:/data/target.xlsx") == "absolute"

    def test_driveless_rooted_path_is_absolute(self) -> None:
        # What real Excel actually writes for a genuine external link to a file on the *same
        # drive* as the linking workbook — drive letter omitted, but still a real absolute
        # link (R3/R4), not same-folder/relative — confirmed via a real Excel-authored fixture
        # (TestScanExternalLinkTargets.test_finds_a_same_drive_absolute_link_target below).
        assert (
            classify_link_target("/Dev/projects/excel-runner/demos/target.xlsx")
            == "absolute"
        )


class TestResolveLinkTarget:
    def test_same_folder_target_resolves_against_linking_workbooks_folder(
        self, tmp_path: Path
    ) -> None:
        linking_path = tmp_path / "linking.xlsx"

        resolved = resolve_link_target("target.xlsx", linking_path)

        assert resolved == (tmp_path / "target.xlsx").resolve()

    def test_relative_subpath_resolves_against_linking_workbooks_folder(
        self, tmp_path: Path
    ) -> None:
        linking_path = tmp_path / "linking.xlsx"

        resolved = resolve_link_target("other/target.xlsx", linking_path)

        assert resolved == (tmp_path / "other" / "target.xlsx").resolve()

    def test_windows_absolute_path_resolves_to_itself(self, tmp_path: Path) -> None:
        linking_path = tmp_path / "linking.xlsx"
        absolute = tmp_path / "elsewhere" / "target.xlsx"

        resolved = resolve_link_target(str(absolute), linking_path)

        assert resolved == absolute.resolve()

    def test_file_uri_resolves_to_the_plain_path(self, tmp_path: Path) -> None:
        linking_path = tmp_path / "linking.xlsx"
        absolute = tmp_path / "elsewhere" / "target.xlsx"

        resolved = resolve_link_target(absolute.as_uri(), linking_path)

        assert resolved == absolute.resolve()


@requires_excel
@requires_working_xlwings_save
class TestScanExternalLinkTargets:
    def test_finds_a_same_folder_link_target(self, tmp_path: Path) -> None:
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            target = app.books.add()
            target.sheets[0].range("A1").value = 5
            target.save(str(tmp_path / "target.xlsx"))
            target.close()

            linking = app.books.add()
            linking.sheets[0].range("A1").formula = "='[target.xlsx]Sheet1'!A1*2"
            linking.save(str(tmp_path / "linking.xlsx"))
            linking.close()

            targets = scan_external_link_targets(tmp_path / "linking.xlsx")

            assert targets == ["target.xlsx"]
        finally:
            registry.close_owned()

    def test_finds_a_subfolder_link_target(self, tmp_path: Path) -> None:
        (tmp_path / "other").mkdir()
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            target = app.books.add()
            target.sheets[0].range("A1").value = 7
            target.save(str(tmp_path / "other" / "target2.xlsx"))
            target.close()

            linking = app.books.add()
            linking.sheets[0].range(
                "A1"
            ).formula = f"='{tmp_path / 'other'}\\[target2.xlsx]Sheet1'!A1*2"
            linking.save(str(tmp_path / "linking.xlsx"))
            linking.close()

            targets = scan_external_link_targets(tmp_path / "linking.xlsx")

            assert targets == ["other/target2.xlsx"]
        finally:
            registry.close_owned()

    def test_finds_a_same_drive_absolute_link_target(self, tmp_path: Path) -> None:
        # Typing the absolute path directly (rather than same-folder + ChangeLink) is
        # deliberate: Excel silently re-normalizes a ChangeLink'd target back to same-folder
        # when both files happen to sit in the same folder at ChangeLink time.
        (tmp_path / "linking_dir").mkdir()
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            target = app.books.add()
            target.sheets[0].range("A1").value = 9
            target.save(str(tmp_path / "target.xlsx"))
            target.close()

            linking = app.books.add()
            linking.sheets[0].range(
                "A1"
            ).formula = f"='{tmp_path}\\[target.xlsx]Sheet1'!A1*2"
            linking.save(str(tmp_path / "linking_dir" / "linking.xlsx"))
            linking.close()

            targets = scan_external_link_targets(
                tmp_path / "linking_dir" / "linking.xlsx"
            )

            # Excel writes two Target relationships for one same-drive absolute link: a
            # relative fallback (ignored — classifies as relative_subpath, R2) and the real
            # drive-omitted-rooted absolute one (R3/R4) that discover_write_intent_link_graph
            # actually keys off of.
            absolute_targets = [
                t for t in targets if classify_link_target(t) == "absolute"
            ]
            assert len(absolute_targets) == 1
            assert (
                resolve_link_target(
                    absolute_targets[0], tmp_path / "linking_dir" / "linking.xlsx"
                )
                == (tmp_path / "target.xlsx").resolve()
            )
        finally:
            registry.close_owned()

    def test_returns_empty_list_with_no_external_links(self, tmp_path: Path) -> None:
        registry = OwnedInstanceRegistry()
        app = registry.spawn()
        try:
            book = app.books.add()
            book.sheets[0].range("A1").value = "no links here"
            book.save(str(tmp_path / "standalone.xlsx"))
            book.close()

            assert scan_external_link_targets(tmp_path / "standalone.xlsx") == []
        finally:
            registry.close_owned()
