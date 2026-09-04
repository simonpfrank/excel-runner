"""Unit tests for tier-3 existence validation (`engine.validate_existence`, opt-in via CLI
`--check-existence`). Unlike tiers 1/2, this one really opens workbooks — read-only, via
openpyxl — to confirm every sheet/defined name a step references by literal name actually
exists.
"""

from pathlib import Path

import openpyxl
import pytest

from excel_runner import engine as validation
from excel_runner.core import Step, ValidationError, WorkbookRef, Workflow
from tests.unit.conftest import workbook_with_external_link


def _workflow(steps: list[Step], workbooks: dict[str, WorkbookRef]) -> Workflow:
    return Workflow(env={}, workbooks=workbooks, steps=tuple(steps))


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
    path = tmp_path / "fixture.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Products"
    sheet["A1"] = "Header"
    workbook.defined_names.add(
        openpyxl.workbook.defined_name.DefinedName("MyRange", attr_text="Products!$A$1")
    )
    workbook.save(path)
    return path


class TestSheetExistence:
    def test_passes_when_sheet_exists(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Products", "range": "A1"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        validation.validate_existence(workflow)  # should not raise

    def test_raises_when_sheet_is_missing(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Nope", "range": "A1"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_existence(workflow)
        assert "Nope" in exc_info.value.detail.message

    def test_a1_range_is_never_checked(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={
                        "workbook": "wb",
                        "sheet": "Products",
                        "range": "ZZ9999:ZZ9999",
                    },
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        validation.validate_existence(workflow)  # should not raise — A1-shaped, skipped

    def test_named_range_must_exist(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={
                        "workbook": "wb",
                        "sheet": "Products",
                        "range": "NoSuchName",
                    },
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_existence(workflow)
        assert "NoSuchName" in exc_info.value.detail.message

    def test_real_named_range_passes(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Products", "range": "MyRange"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        validation.validate_existence(workflow)  # should not raise

    def test_create_sheet_satisfies_a_later_reference(
        self, workbook_path: Path
    ) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="create_sheet",
                    params={"workbook": "wb", "name": "Archive"},
                ),
                Step(
                    id="s2",
                    action="write_cell",
                    params={
                        "workbook": "wb",
                        "sheet": "Archive",
                        "cell": "A1",
                        "value": 1,
                    },
                ),
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        validation.validate_existence(workflow)  # should not raise

    def test_write_to_sheet_not_yet_created_raises(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="write_cell",
                    params={
                        "workbook": "wb",
                        "sheet": "Archive",
                        "cell": "A1",
                        "value": 1,
                    },
                ),
                Step(
                    id="s2",
                    action="create_sheet",
                    params={"workbook": "wb", "name": "Archive"},
                ),
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        with pytest.raises(ValidationError):
            validation.validate_existence(workflow)

    def test_delete_sheet_then_reference_raises(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="delete_sheet",
                    params={"workbook": "wb", "sheet": "Products"},
                ),
                Step(
                    id="s2",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Products", "range": "A1"},
                ),
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        with pytest.raises(ValidationError):
            validation.validate_existence(workflow)

    def test_templated_sheet_is_skipped(self, workbook_path: Path) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={
                        "workbook": "wb",
                        "sheet": "{{ steps.prior.output.sheet }}",
                        "range": "A1",
                    },
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        validation.validate_existence(
            workflow
        )  # should not raise — can't know statically

    def test_nonexistent_workbook_file_is_skipped_entirely(self) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Anything", "range": "A1"},
                )
            ],
            {
                "wb": WorkbookRef(
                    name="wb", file="C:/does/not/exist.xlsx", create_if_missing=True
                )
            },
        )
        validation.validate_existence(
            workflow
        )  # should not raise — nothing to check yet

    def test_copy_checks_both_source_and_target_sheets(
        self, workbook_path: Path
    ) -> None:
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="copy",
                    params={
                        "source": {"workbook": "wb", "sheet": "Products"},
                        "target": {"workbook": "wb", "sheet": "Missing"},
                    },
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(workbook_path))},
        )
        with pytest.raises(ValidationError) as exc_info:
            validation.validate_existence(workflow)
        assert "Missing" in exc_info.value.detail.message


class TestSupportedLinkLayout:
    """Tier-3 also refuses link layouts this tool cannot honour
    (docs/backend_eligibility_build_plan.md W7).

    Every workbook a run touches is staged into one flat scratch folder, so a link that points
    *into another folder* (`data\\prices.xlsx`) has no correct meaning once staged: leaving it
    alone breaks it, rewriting it would silently point somewhere the author never wrote.
    Refusing up front is the only honest option.
    """

    def test_relative_subpath_link_is_refused(self, tmp_path: Path) -> None:
        path = workbook_with_external_link(
            tmp_path / "linking.xlsx", target="data/prices.xlsx"
        )
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Sheet", "range": "A1"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(path))},
        )

        with pytest.raises(ValidationError) as exc_info:
            validation.validate_existence(workflow)

        assert "data/prices.xlsx" in exc_info.value.detail.message
        assert "another folder" in exc_info.value.detail.message

    def test_same_folder_link_is_accepted(self, tmp_path: Path) -> None:
        """R1 survives the flat scratch layout intact — both workbooks land side by side."""
        path = workbook_with_external_link(
            tmp_path / "linking.xlsx", target="prices.xlsx"
        )
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Sheet", "range": "A1"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(path))},
        )

        validation.validate_existence(workflow)  # must not raise

    def test_absolute_link_is_accepted(self, tmp_path: Path) -> None:
        """R3/R4 point at a fixed location that staging never moves, so they stay valid."""
        path = workbook_with_external_link(
            tmp_path / "linking.xlsx", target=str(tmp_path / "elsewhere" / "prices.xlsx")
        )
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "wb", "sheet": "Sheet", "range": "A1"},
                )
            ],
            {"wb": WorkbookRef(name="wb", file=str(path))},
        )

        validation.validate_existence(workflow)  # must not raise

    def test_a_templates_unsupported_link_is_caught_before_the_workbook_exists(
        self, tmp_path: Path
    ) -> None:
        """The workbook inherits the layout it cannot honour, so the refusal must happen on
        the first run too — not only once the file happens to exist."""
        template = workbook_with_external_link(
            tmp_path / "template.xlsx", target="data/prices.xlsx"
        )
        workflow = _workflow(
            [
                Step(
                    id="s1",
                    action="read_range",
                    params={"workbook": "report", "sheet": "Sheet", "range": "A1"},
                )
            ],
            {
                "tmpl": WorkbookRef(name="tmpl", file=str(template)),
                "report": WorkbookRef(
                    name="report",
                    file=str(tmp_path / "report.xlsx"),
                    create_if_missing=True,
                    template="tmpl",
                ),
            },
        )

        with pytest.raises(ValidationError) as exc_info:
            validation.validate_existence(workflow)

        assert "data/prices.xlsx" in exc_info.value.detail.message
