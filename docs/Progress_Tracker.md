# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** In Progress
**Working on:** Build order items 1–4 (Spec §8) all done (item 4 for what's cleanly buildable
now — see below). Git repo initialized, four commits on `main`.

Item 2 surfaced a real design bug: PRD §10.1/Spec §2.2 originally said "render the whole YAML
file as one Jinja2 text pass, then parse it" — can't work for `{{ steps.* }}` (no step has run
at load time). Corrected to: parse YAML directly, resolve `env:`/`workbooks:` once at load time
(env-only), `Step.params`/`if_expr` left raw and resolved per-step during execution (env +
steps). Docs updated in place.

Item 3 (action registry + first 5 actions) surfaced four corrections, all recorded in
Spec §4/§5.1/§5.2: `ActionResult`/`WorkbookSession` moved to `core.py` (avoids a circular import
between `engine.py`'s registry and `actions.py`); no `workbook` param on action functions (the
not-yet-built runner resolves it into `session` before calling); actions call `backends.py`
directly rather than through a `session.<something>` indirection; `WorkbookSession` needed a
`path` field the original sketch missed.

Item 4 (remaining v1 file-backend actions) built 9 more: `copy`, `write_range`, `write_row`
(base + positional modes), `insert_range` (whole-row/column only), `set_column_width`,
`find_headers_row`, `find_row`, `find_column`, `find_columns`, `read_metadata` (properties/cells
sub-cases) — 14 actions built in total now. This batch established the error-handling policy
(structured `ActionResult(status="error")` for a normal "search found nothing" outcome vs. a
raised exception for a genuine authoring mistake — Spec §4) and surfaced real deferrals:
- `write_table`, `write_row`'s by-header mode, and `aggregate` all need step-output context
  (`source`/`headers_from` are step-id references, not templated values) that doesn't exist
  until `runner.py` threads it through — item 7, not blocking now.
- `read_links` — **empirically** (not just theoretically) downgraded: a spike showed openpyxl
  never creates an external-link relationship from a written formula, only reflects one already
  in a real-Excel-created file. Moved into the same deferred bucket as `write_links`. PRD §7/§8
  and §12 corrected — the earlier "resolved: reading looks solid" call was wrong, based on docs
  not a real test.
- `copy` needed two `WorkbookSession` params (source + target) — its YAML shape has two nested
  workbook refs, not one flat `workbook:` field. Built and tested; `runner.py` will need
  matching two-session wiring later.
- `read_metadata`'s `cells` sub-case needed a `sheet` param the original PRD §7 catalog didn't
  list.

All quality gates pass: 152/152 tests, ruff clean, mypy --strict clean (`excel_runner` +
`tests`), radon cc clean, vulture clean, **100% branch coverage** across all 4 modules.
**Next step:** Build order item 5 (Spec §8) — `engine.py`'s `SessionManager` (multi-workbook
lifecycle: lazy-open, promotion, close-all) and `ScratchManager` (scratch-copy execution model,
PRD §6.3.1). The `WorkbookSession` data shape is already built; `SessionManager` itself isn't.
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
| `WorkbookSession` / `SessionManager` — `engine.py` §5.2 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `ScratchManager` — `engine.py` §5.3 | ❌ | ❌ | ❌ | ❌ | ❌ |

## Phase 5 — Validation (Spec §5.4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Tier 1: static schema validation — `engine.py` §5.4 | ❌ | ❌ | ⏭️ | ❌ | ⏭️ |
| Tier 2: dry-run / step-graph validation — `engine.py` §5.4 | ❌ | ❌ | ❌ | ❌ | ❌ |

## Phase 6 — Runner, audit, public API (Spec §6)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| `AuditLogger` — `runner.py` §6.2 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `run_workflow` orchestration — `runner.py` §6.1 | ❌ | ❌ | ❌ | ❌ | ❌ |
| Public API surface (`__init__.py` re-exports) — `runner.py` §6.3 | ❌ | ❌ | ❌ | ❌ | ❌ |
| Crash-safety integration test (Spec §7) | ⏭️ | ⏭️ | ❌ | ⏭️ | ❌ |

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
