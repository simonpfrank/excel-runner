# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** In Progress
**Working on:** Build order items 1–7 (Spec §8) all done (items 4/5/6 for what's cleanly
buildable now). Git repo initialized, seven commits on `main`.

**Items 1–6, summarized** (full detail in git log / Spec §2–§5 correction notes): data model +
errors (1); loading/templating, found the whole-file "render then parse" plan couldn't work for
`{{ steps.* }}` (2); action registry + first 5 actions, found `ActionResult`/`WorkbookSession`
needed to live in `core.py` to avoid a circular import (3); 9 more actions taking the total to
14, established the `ActionResult(status="error")` vs. raised-exception policy, found
`read_links` empirically broken (4); `SessionManager`/`ScratchManager`, mode caller-specified
pending tier 2, coverage caught a missing-`mkdir` bug (5); both validation tiers, found PRD
§9.1's fourth example message isn't implementable without workbook access — corrected, not
faked (6).

**Item 7 (`runner.py` — orchestration + audit logging, Spec §6.1/§6.2)** — built, and this is
where several things deferred since items 3–6 finally got resolved for real: `copy`'s
two-session dispatch, the `workbook`-field-stripping translation, and (a genuine open design
question until now) whether an action's `ActionResult(status="error")` halts the run — resolved
by what `if:` is actually for: PRD's own `if: steps.x.status == 'success'` example only makes
sense if a failed step *doesn't* abort the run before later steps can check it. So the loop
continues past a failed step, but `RunResult.status` is `"error"` if any step failed, and
nothing gets committed in that case — only a *raised* exception aborts the loop outright.

Two real bugs surfaced writing the first integration test, both fixed with tests, not patched
around: the audit log was being written inside the same directory `ScratchManager.cleanup()`
deletes, so a *successful* run deleted its own audit trail (fixed by splitting scratch/ from
the parent run dir); and a dict output key literally named `"values"` (`read_range`'s own
output shape, PRD §10.4) was shadowed by Python's real `dict.values()` method under Jinja2's
default attribute resolution — `{{ steps.x.output.values }}` silently returned the bound method
instead. Fixed generically (a custom Jinja `Environment` that tries item access before
attribute access), not by renaming around the one collision, since any future action's output
key could hit the same class of bug. Also found and fixed a real typing gap while adding a
regression test: `@file_action`/`@com_action` were erasing every action's parameter types via
`Callable[..., ActionResult]`, silently defeating mypy at every call site — switched to
`ParamSpec`, which surfaced (and let us fix) `read_metadata` silently mishandling an
unsupported `target` value instead of rejecting it.

`tests/integration/test_run_workflow.py` is the first genuine end-to-end suite — real
`workflow.yaml` text run through `run_workflow()` against real openpyxl workbooks, exactly the
shape agreed on with the user. Fixture workbooks are generated in code, not committed as static
files in `tests/data/` (a correction to the original testing-approach sketch, Spec §7) — more
reviewable, consistent with every other test in this codebase.

All quality gates pass: 220/220 tests, ruff clean, mypy --strict clean (`excel_runner` +
`tests`), radon cc clean, vulture clean, **100% branch coverage** across all 6 modules.
**Next step:** Still owed before moving on, per the user's "pause for integration testing"
request: a dedicated mid-run-crash integration test (PRD §6.3/§6.3.1's actual crash-safety
requirement, not just the design note — Spec §7). Then build order item 8 — `runner.py` §6.3,
the public API surface (`__init__.py` re-exports, `list_actions()`).
**Notes:** `aggregate` and `update_summary_table`'s exact parameters are still explicitly
flagged as open in the PRD — don't block on them. Tracker below stays function/class-granular
even though source files are consolidated — see Spec §7.

## Status legend
❌ Not Done · 🟡 In Progress · ✅ Done — Results: ✅ Pass · ❌ Fail · ⏭️ N/A

## Phase 1 — `core.py`: data model, errors, templating (Spec §2)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Data model dataclasses (`WorkbookRef`, `Step`, `Workflow`) — §2.1 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Error types (`ErrorDetail`, exception classes) — §2.3 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Loading pipeline (`load`) — §2.2 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Templating (`resolve_value`, `evaluate_condition`) — §2.2 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |

## Phase 2 — Registry + first action slice (Spec §5.1, §4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Action registry (`ActionSpec`, `discover_actions`) — `engine.py` §5.1 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `open` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `save` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `close` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `read_range` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_cell` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `WorkbookSession`/`ActionResult` (moved to `core.py`) | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| File-backend primitives for the 5 actions above — `backends.py` §3 | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 3 — Remaining v1 file-backend actions (Spec §3, §4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| File-backend primitives for this batch — `backends.py` §3 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `copy` action (two-session signature — see notes above) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_range` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_row` action (base + positional modes) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_row` by-header mode | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_table` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `insert_range` action (whole-row/column only) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `insert_range` partial-range support | ❌ | ❌ | ❌ | ❌ | ❌ |
| `set_column_width` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_headers_row` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_row` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_column` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_columns` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `aggregate` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_metadata` action (properties/cells) | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 4 — Execution model (Spec §5.2, §5.3)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| `SessionManager` — `engine.py` §5.2 (`promote_to_com` excluded — needs item 9) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `ScratchManager` — `engine.py` §5.3 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `backends.create_workbook` (supports `create_if_missing`) | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 5 — Validation (Spec §5.4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Tier 1: static schema validation — `engine.py` §5.4 (real-defined-names check excluded, PRD §12) | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Tier 2: dry-run / step-graph validation — `engine.py` §5.4 | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 6 — Runner, audit, public API (Spec §6)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| `AuditLogger` — `runner.py` §6.2 | ✅ | ✅ | ✅ | ✅ | ✅ |
| `run_workflow` orchestration — `runner.py` §6.1 | ⏭️ | ✅ | ✅ | ⏭️ | ✅ |
| Public API surface (`__init__.py` re-exports) — `runner.py` §6.3 | ❌ | ❌ | ❌ | ❌ | ❌ |
| Crash-safety (mid-run interruption) integration test (Spec §7) | ⏭️ | ⏭️ | ❌ | ⏭️ | ❌ |

## Phase 7 — COM (Windows-dependent, later phase per PRD §8)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| COM-backend primitives — `backends.py` §3 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `OwnedInstanceRegistry` — `backends.py` §3.1 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `recalculate` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `run_macro` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `refresh_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_metadata` action (textbox sub-case) | ❌ | ❌ | ❌ | ❌ | ❌ |

## Backlog / deferred (per PRD §8 and §12)

| Item | Status |
|---|---|
| `update_summary_table` — exact parameters | Not designed yet, deliberately deferred |
| `aggregate`, `write_table`, `write_row` by-header mode | Need step-output context — blocked on `runner.py` (build order item 7), not just deferred by choice |
| `read_links` (+ already-deferred `write_links`) | Blocked on a real Excel-generated fixture (or manual XML/zip surgery) — empirically confirmed openpyxl can't create the relationship itself |
| `export_pdf` | Backlog |
| AI-authoring inspection actions (`list_sheets`, `describe_sheet`) | Planned, next phase after core engine works (PRD §9) |
| Instance-ownership tagging mechanism across crashes (PRD §6.2.1/§12) | Open question |
| Scratch-directory collision avoidance for concurrent runs (PRD §6.3.1/§12) | Open question |
| CLI / MCP wrapper | Deferred (PRD §3/§5) |
