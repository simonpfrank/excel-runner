# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** Ready for Next Phase
**Working on:** Build order items 1–8 (Spec §8) — the entire v1 file-backend core engine — are
now done, plus the crash-safety integration test. Git repo initialized, nine commits on `main`.
Only item 9 (COM phase, Windows-dependent, later) and item 10 (deferred/flagged items) remain
from the original build order.

**Items 1–6, summarized** (full detail in git log / Spec §2–§5 correction notes): data model +
errors (1); loading/templating, found the whole-file "render then parse" plan couldn't work for
`{{ steps.* }}` (2); action registry + first 5 actions, found `ActionResult`/`WorkbookSession`
needed to live in `core.py` to avoid a circular import (3); 9 more actions taking the total to
15 (miscounted as 14 at the time — corrected in item 8's notes below), established the
`ActionResult(status="error")` vs. raised-exception policy, found
`read_links` empirically broken (4); `SessionManager`/`ScratchManager`, mode caller-specified
pending tier 2, coverage caught a missing-`mkdir` bug (5); both validation tiers, found PRD
§9.1's fourth example message isn't implementable without workbook access — corrected, not
faked (6).

**Item 7 (`runner.py` — orchestration + audit logging, Spec §6.1/§6.2)** — built. Resolved
`copy`'s two-session dispatch and the `workbook`-field-stripping translation deferred since
items 3–4, plus a genuine open design question: an action's `ActionResult(status="error")`
does *not* stop the loop — resolved by what `if:` is actually for (PRD's own
`if: steps.x.status == 'success'` example only makes sense if a failed step doesn't abort the
run before later steps can check it). `RunResult.status` is `"error"` if any step failed, and
that (not loop completion) gates whether anything commits — only a *raised* exception aborts
the loop outright. Surfaced three real bugs, all fixed with regression tests: the audit log was
being written inside the directory `ScratchManager.cleanup()` deletes, so a successful run
deleted its own audit trail; a dict output key literally named `"values"` (`read_range`'s own
PRD §10.4 output shape) was shadowed by Python's real `dict.values()` method under Jinja2's
default attribute resolution, fixed generically (a custom Jinja `Environment` trying item
access first) rather than by renaming around the one collision; and `@file_action`/
`@com_action` were erasing every action's parameter types (`Callable[..., ActionResult]`),
silently defeating mypy at every call site — switched to `ParamSpec`, which then let us catch
and fix `read_metadata` quietly mishandling an unsupported `target` instead of rejecting it.
First genuine end-to-end suite, `tests/integration/test_run_workflow.py`; fixture workbooks
generated in code, not committed as static files (correction to the original `tests/data/`
plan, Spec §7).

**Crash-safety follow-up (user-requested before moving to item 8)** — found a fourth real bug,
the most significant one this session: **the scratch copy left behind after a crash didn't
actually contain any in-progress work.** openpyxl writes stay in memory until an explicit save,
and nothing was flushing them to the scratch file mid-run — so the "recovery artifact" PRD
§6.3.1 promises was, in practice, just an unchanged copy of the original. Fixed with
`SessionManager.checkpoint()`, called after every step in the loop, saving each dirty staged
session to its scratch file as it goes — found by writing the crash test properly (asserting
the scratch file's *content*, not just its existence) and watching it fail on the first
attempt. `TestCrashSafety` now covers: real file untouched, scratch copy genuinely has the
prior step's work, audit log survives, and — the strongest cross-platform proof sessions were
actually closed — a later valid run against the same file just works. "No orphaned Excel
process" stays untested until a COM backend exists (item 9) to spawn one.

**Item 8 (public API surface, Spec §6.3)** — built: `excel_runner/__init__.py` re-exports
(`run_workflow`, `RunResult`, `StepResult`, `Workflow`, `Step`, `WorkbookRef`, `list_actions`,
`ActionSpec`), and `list_actions()` itself (`discover_actions` wired to the real module — not a
second source of truth). Found a real gap: `ActionSpec` never actually had the `description`
field the original design's own words assumed ("name/docstring/param_schema") — without it,
`list_actions()` couldn't fulfill its one stated purpose (a future agent-tool wrapper needs a
description, not just a name and parameter shape). Fixed: `description` is now populated from
each action's docstring, first line only. Also caught, via a test asserting the real action
count: the "14 actions" repeated across items 4–7's notes was an off-by-one — it's 15. Corrected
in Spec §5.1/§6.3 and above; not rewriting git history for it.

All quality gates pass: 239/239 tests, ruff clean, mypy --strict clean (`excel_runner` +
`tests`), radon cc clean, vulture clean, **100% branch coverage** across all 6 modules.
`README.md` written and committed (`4eb5a37`) — every documented example verified end-to-end
against real workbooks, not just asserted.

**Design decision, 2026-08-19**: user flagged that without a way to halt a run early, avoiding a
failed lookup's downstream steps means repeating the same `if:` on every one of them. Agreed
design: a `stop` control-flow action, driven by the existing `if:` mechanism (no new per-step
field), with a distinct `StepResult(status="stopped")` for every step after it — kept separate
from `"skipped"` so the audit log can tell "this step's own condition skipped it" from "the run
ended before we got here." Written up in PRD §6.9 and Specification.md §4/§6.1/§8 (now build
order item 9, renumbering COM to 10 and deferred items to 11). A related but separate idea —
grouped `if:` blocks spanning multiple steps, to avoid repeating the same condition — was
deliberately **backlogged, not designed**: possible over-engineering, parked in PRD §12 until
real workflows show it's an actual recurring need.

**Item 9 (`stop` control-flow action, PRD §6.9/Spec §4/§6.1/§8)** — built via TDD. New
`@control_action` decorator (capability `"none"`) since `stop` takes no `session` and has no
`workbook:` field at all; joined `copy` in `_SCHEMA_EXEMPT_ACTIONS` for the same reason. New
`StepResult(status="stopped")`, distinct from `"skipped"`, for every step after a triggered
`stop` — kept `RunResult.step_results` at one entry per workflow step, and audit-logs the
stopped steps too. `stop` doesn't set `any_failed` on its own, so a deliberate early exit still
commits prior work while "not found → stop" still discards (the failed lookup already set
that). No new commit logic needed — only the loop's early-exit needed writing. Action count is
now **16** (was 15) — `list_actions()`'s count assertion caught every place that needed updating.
All quality gates pass: 248/248 tests, 100% branch coverage, ruff/mypy --strict/radon/vulture
clean.

**Next step:** Not yet decided — build order item 10 (COM phase, Windows-dependent) or item 11
(deferred items: `update_summary_table`, `aggregate`, `export_pdf`, AI-authoring inspection
actions). README's action reference still needs a `stop` entry.
**Notes:** `aggregate` and `update_summary_table`'s exact parameters are still explicitly
flagged as open in the PRD — don't block on them. Four things parked in PRD §12, none designed
or scheduled: grouped `if:` blocks; a "replay nice" desktop-comfort mode (visible Excel replay
of an already-finished real run, via xlwings — deliberately not a second execution backend);
openpyxl silently dropping charts it can't parse on any save (a real, evidence-backed risk to
the core save path, not just a feature idea — confirmed via openpyxl's own reader-warning
mechanism and a local round-trip test; verifying the real-Excel-chart case still needs a fixture
we don't have, same blocker as `read_links`); and a conversational agent-driven
spreadsheet-authoring product, explicitly flagged as likely a different product from this one.
Tracker below stays function/class-granular even though source files are consolidated — see
Spec §7.

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
| `SessionManager.checkpoint()` — `engine.py` §5.2 | ✅ | ✅ | ✅ | ✅ | ✅ |
| Public API surface (`__init__.py` re-exports, `list_actions()`, `ActionSpec.description`) — `runner.py` §6.3 | ✅ | ✅ | ✅ | ✅ | ✅ |
| Crash-safety (mid-run interruption) integration test (Spec §7) | ⏭️ | ⏭️ | ✅ | ⏭️ | ✅ |
| `stop` control-flow action (`actions.py`) + `"stopped"` status + loop early-exit — `runner.py` §6.1, PRD §6.9 | ✅ | ✅ | ✅ | ✅ | ✅ |

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
