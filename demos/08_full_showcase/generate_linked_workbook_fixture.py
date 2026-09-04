"""One-time script: (re)generates demos/08_full_showcase/linked_workbook/originals/linked.xlsx
with a genuine, absolute-path R4 external link to *this checkout's* demos/08_full_showcase/catalog.xlsx.

Why this can't be a plain create_if_missing/template workbook (like every other originals/*.xlsx
in this demo): a real external link is authored by Excel itself (ChangeLink over COM), and the
absolute path it stores is tied to wherever this repo happens to be checked out. Baking that into
a git-committed binary would go stale the moment the repo is cloned somewhere else — so this
script (re)builds the fixture fresh, locally, instead.

Run this once per checkout (or again if this repo folder ever moves), BEFORE running
demos/08_full_showcase.yaml:

    .venv\\Scripts\\python demos\\08_full_showcase\\generate_linked_workbook_fixture.py

What it does:
    1. Ensures demos/08_full_showcase/catalog.xlsx exists (the real working copy, not the
       originals/ template) — copying it fresh from originals/catalog.xlsx if missing, since
       the link below points at this exact path.
    2. Builds a new workbook whose A1 formula spells out catalog.xlsx's *absolute folder path*
       directly (`='<abs path>\\[catalog.xlsx]Products'!F2`), then saves it straight into
       demos/08_full_showcase/linked_workbook/originals/ — a different folder than catalog.xlsx
       itself. Typing the absolute path directly (rather than same-folder + `ChangeLink`) is
       deliberate: Excel silently re-normalizes a `ChangeLink`'d target back to a same-folder
       reference when both files happen to sit in the same folder at `ChangeLink` time, which
       would produce an R1 (same_folder) link instead of the genuine R3/R4 (absolute) one this
       fixture needs — confirmed by the direct-absolute-formula pattern used in
       tests/unit/test_link_discovery.py's `test_finds_a_subfolder_link_target`.
"""

import shutil
from pathlib import Path

from excel_runner import backends

DEMO_DIR = Path(__file__).parent
CATALOG_ORIGINAL_PATH = DEMO_DIR / "originals" / "catalog.xlsx"
CATALOG_REAL_PATH = DEMO_DIR / "catalog.xlsx"
LINKED_DIR = DEMO_DIR / "linked_workbook" / "originals"
LINKED_FINAL_PATH = LINKED_DIR / "linked.xlsx"


def main() -> None:
    if not CATALOG_REAL_PATH.exists():
        shutil.copy2(CATALOG_ORIGINAL_PATH, CATALOG_REAL_PATH)
    catalog_folder = CATALOG_REAL_PATH.resolve().parent

    registry = backends.OwnedInstanceRegistry()
    app = registry.spawn()
    try:
        linked = app.books.add()
        sheet_name = linked.sheets[0].name
        linked.sheets[0].range(
            "A1"
        ).formula = f"='{catalog_folder}\\[catalog.xlsx]Products'!F2"
        LINKED_DIR.mkdir(parents=True, exist_ok=True)
        linked.save(str(LINKED_FINAL_PATH))
        linked.close()
    finally:
        registry.close_owned()

    print(f"Wrote {LINKED_FINAL_PATH}")
    print(
        f"  -> external link (sheet {sheet_name!r}!A1) -> {CATALOG_REAL_PATH.resolve()}"
    )


if __name__ == "__main__":
    main()
