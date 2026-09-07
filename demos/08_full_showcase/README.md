# Demo 08 — Full Showcase

Exercises every registered `excel_runner` action at least once, in one coherent story:
discover a "Products" catalog's structure, pull totals out of it (as computed values and as
formula text), copy the table (formulas intact) into a fresh "report" workbook, add a genuine
cross-workbook link back to the catalog, write new formulas/labels/structure around it, then
recalculate catalog + report together (they're genuinely linked) before committing.

Two workflow files, run in sequence:

- **`08_full_showcase.yaml`** — builds everything. No assertions, only actions.
- **`08_full_showcase_validate.yaml`** — opens what the first file produced and asserts the
  numbers are correct via `read_range` + `stop`/`if:` guards. This is excel_runner validating
  its own output in YAML, not a bespoke Python assertion script.

The full step-by-step rationale (why each step is placed where it is, which output feeds
which later step) lives in [`08_full_showcase_script.md`](../08_full_showcase_script.md) — a
build-planning document, kept for that detail rather than duplicated here.

## Running it

```powershell
demos\reset_08_fixtures.bat
.venv\Scripts\python -m excel_runner demos\08_full_showcase.yaml
.venv\Scripts\python -m excel_runner demos\08_full_showcase_validate.yaml
```

**Always reset before a re-run.** Both workflows write to their working copies and commit —
a second run starts from the previous run's committed output, not a clean fixture, and the
structural steps (`rename_sheet`, `create_sheet`, ...) fail because that work is already done.
`reset_08_fixtures.bat` copies the three `originals/` fixtures back over the three working
copies and deletes `excel_runner_runs/08_full_showcase` (the run's audit log and scratch
staging area) so the next run starts clean.

## Files

```
08_full_showcase/
  catalog.xlsx                     working copy — source data (recreated from originals/ if missing)
  report.xlsx                      working copy — built up by the workflow
  originals/
    catalog.xlsx                   pristine fixture: "Products" sheet + a disposable "Scratch" sheet
    report.xlsx                    pristine fixture: blank, single default sheet
  linked_workbook/
    linked.xlsx                    working copy — carries a real external link to catalog.xlsx
    originals/
      linked.xlsx                  pristine fixture (see "The linked workbook fixture" below)
```

`catalog`/`report`/`linked_workbook` are declared with `create_if_missing: true` and a
`template:` pointing at their `originals/` counterpart, so a working copy that's been deleted
is recreated fresh on the next run — but once a run has committed to it, it stays as-is
(same as every other demo's `working_copy:` pattern, e.g. demos 05/07).

## The linked workbook fixture

`linked_workbook/originals/linked.xlsx` is not a plain checked-in template like the other two
`originals/*.xlsx` files. It carries a genuine, Excel-authored external link — an **absolute
path** formula (`='<abs path>\[catalog.xlsx]Products'!F2`) pointing at *this checkout's*
`demos/08_full_showcase/catalog.xlsx`. That's deliberate: it's what makes this demo exercise
the real R4 (absolute-path external link) machinery
(`docs/backend_eligibility_build_plan.md`) — `engine.SessionManager` staging-time link rewire,
`com_change_link`, revert-before-commit — rather than a same-folder (R1) link that never
triggers any of that.

The formula is typed as an absolute path rather than built via same-folder + `ChangeLink`,
because Excel silently re-normalizes a `ChangeLink`'d target back to a same-folder reference
when both files happen to sit in the same folder at `ChangeLink` time — which would produce
an R1 link instead of the R3/R4 one this fixture needs.

**Consequence: the fixture is tied to the absolute path of this checkout.** If this repo
folder is ever moved or cloned elsewhere, `linked.xlsx`'s stored link target goes stale — it
still points at the *old* location's `catalog.xlsx`. `reset_08_fixtures.bat` will still copy
it into place without error (it's just a file copy), but the link inside it will be wrong,
and running `08_full_showcase.yaml` will fail (or silently read no data) at
`open_linked_workbook`/`write_linked_workbook_label`.

**There is currently no script to rebuild it.** The generator that originally created this
fixture (`generate_linked_workbook_fixture.py`) has been removed. To rebuild `linked.xlsx`
after moving the repo, recreate it by hand:

1. Ensure `demos/08_full_showcase/catalog.xlsx` exists (copy it from `originals/catalog.xlsx`
   if not).
2. Using `xlwings`, `openpyxl.load_workbook(..., keep_vba=False)` won't help here — the link
   must be authored by real Excel so the relationship is genuine. Spawn an Excel instance via
   `excel_runner.backends.OwnedInstanceRegistry`, create a new workbook, set cell `A1`'s
   formula to `='<absolute path to demos/08_full_showcase>\[catalog.xlsx]Products'!F2`, and
   save it to `demos/08_full_showcase/linked_workbook/originals/linked.xlsx`.
3. Run `reset_08_fixtures.bat` again to pick up the new fixture.

If this repo is expected to move or be cloned by others, this is worth turning back into a
committed script rather than a manual recipe.

## Known limitation: relative link path degradation

Independent of the above, `com_change_link`'s round trip (real → scratch → real) rewrites the
link's *relative* fallback path on each save. The fixture's `../../catalog.xlsx` becomes
`../../../../demos/08_full_showcase/catalog.xlsx` after a run — both are wrong as relative
paths from `linked_workbook/`'s perspective. This is harmless in practice: the absolute path
(what this fixture actually relies on, and what Excel falls back to) stays correct. It's a
cosmetic staleness in the relative fallback only, left as-is rather than chased further.
