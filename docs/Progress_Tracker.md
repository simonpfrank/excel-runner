# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** In Progress
**Working on:** Build order items 1–6 (Spec §8) all done (items 4/5/6 for what's cleanly
buildable now). Git repo initialized, six commits on `main`.

**Earlier items, summarized** (full detail in git log / Spec §2–§5 correction notes):
item 1 — data model + errors in `core.py`. Item 2 — loading/templating; found the whole-file
"render then parse" plan couldn't work for `{{ steps.* }}`, corrected to parse-then-resolve-
per-field. Item 3 — action registry + first 5 actions; found `ActionResult`/`WorkbookSession`
needed to live in `core.py` to avoid a circular import, `workbook` gets stripped before an
action is called, actions call `backends.py` directly. Item 4 — 9 more actions (14 total);
established the `ActionResult(status="error")` vs. raised-exception policy; found `read_links`
empirically broken (openpyxl can't create the relationship it would need to read back);
`copy` needed a two-session signature. Item 5 — `SessionManager`/`ScratchManager`; mode is
caller-specified for now (tier 2 wasn't built yet); coverage caught a real missing-`mkdir` bug
on the read-only + `create_if_missing` path.

**Item 6 (both validation tiers, Spec §5.4)** — built, with one real gap found and corrected
rather than papered over: **PRD §9.1's fourth example message (checking a range against a
workbook's real defined names) isn't actually implementable in either tier** — it needs to
open the workbook, which contradicts both tiers' own "no workbook access" premise. The other
three examples are genuinely workbook-access-free and are what tier 1 builds: action-exists
(with a fuzzy "did you mean" suggestion), no-unrecognized-params, no-missing-required-params,
param-type-matches (handles `Literal`/`Union`/generic annotations, not just plain types), and
step-id-reference resolution (exists + defined earlier, with a fuzzy suggestion). `copy` is
exempt from the param-schema checks — its raw YAML shape (`source:`/`target:` dicts) still
doesn't match its Python signature, same reason as item 4's note, unresolved until the runner's
translation layer exists. Tier 2 (`plan()`) checks every referenced workbook is declared and
infers read/write mode via a small static write-actions lookup table (deliberately simple, not
deep analysis) — `copy` again gets the safe-fallback treatment (every workbook it touches
marked `read_write`, since source vs. target can't be told apart statically yet). PRD §9.1/§12
corrected to record the workbook-access gap as an open item.

All quality gates pass: 201/201 tests, ruff clean, mypy --strict clean (`excel_runner` +
`tests`), radon cc clean, vulture clean, **100% branch coverage** across all 4 modules — two
lines needed an honest `# pragma: no cover` (structurally unreachable via the current design,
documented why) rather than a contrived test.
**Next step:** Build order item 7 (Spec §8) — `runner.py` §6.1/§6.2: the orchestration loop
(`run_workflow`) and audit logging. This is also the point `copy`'s two-session wiring and the
`workbook`-field-stripping translation actually get implemented for real, and the first point
genuine end-to-end YAML-driven integration tests become possible (discussed with the user —
plan is to pause for integration testing once this item and item 8 land, before the COM phase).
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
