"""Shared fixtures/markers for unit tests (Spec sec 7)."""

import sys
from pathlib import Path

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
