"""Shared fixtures/markers for unit tests (Spec sec 7)."""

import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest


def _excel_available() -> bool:
    """Whether this machine can actually spawn and drive Excel via xlwings.

    Cheap existence check, not a spawn — spawning just to probe availability would slow down
    every collection run. Windows is assumed available since it's the real target (PRD sec 4);
    verified for real once a Windows environment exists (PRD sec 12).
    """
    if sys.platform == "darwin":
        return Path("/Applications/Microsoft Excel.app").exists()
    return sys.platform == "win32"


requires_excel = pytest.mark.skipif(not _excel_available(), reason="requires a live Excel install")

requires_working_xlwings_save = pytest.mark.skipif(
    sys.platform != "win32",
    reason=(
        "xlwings' save() is confirmed broken on this Mac's Excel build (Parameter error -50, "
        "reproduced via both save-as and in-place save, matches known xlwings GitHub issues) — "
        "Spec sec 3.1. Not a mock: the code and test are real, only the platform this can "
        "actually run on is restricted, per PRD sec 4/sec 12's macOS-now/Windows-later plan."
    ),
)

_EXTERNAL_LINK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/externalLinkPath" Target="{target}" TargetMode="External"/>'
    "</Relationships>"
)


def plain_workbook(path: Path) -> Path:
    """A minimal, real .xlsx with no external links — the "nothing special about it" control
    case, and a valid zip wherever a test needs one."""
    path.parent.mkdir(parents=True, exist_ok=True)
    openpyxl.Workbook().save(path)
    return path


def workbook_with_external_link(path: Path, target: str = "target.xlsx") -> Path:
    """A real .xlsx carrying one outbound external-link relationship, built by rewriting the
    zip openpyxl produces — the same on-disk shape `scan_external_link_targets` reads.

    Lives here rather than in one test module because save-blocker inspection, session
    promotion and link-layout validation all need the same fixture and must all be reading
    the same on-disk shape. Excel is only needed to *author* a link, not to *carry* one, so
    every test built on this runs everywhere.

    Args:
        path: Where to write the workbook.
        target: The link target to record, exactly as it would appear in the rels XML —
            a bare filename, a `../`-style relative path, or an absolute/UNC path.
    """
    source = plain_workbook(path.with_name("_source_" + path.name))
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(path, "w") as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item.filename))
        dst.writestr(
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            _EXTERNAL_LINK_RELS.format(target=target),
        )
    source.unlink()
    return path
