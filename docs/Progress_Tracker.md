# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** In Progress
**Working on:** Build order items 1 and 2 (Spec §8) both done — `core.py`'s data model, error
types, and loading/templating pipeline, all TDD, all green. Git repo initialized, initial
commits made.

Item 2 surfaced a real design bug during implementation: PRD §10.1/Spec §2.2 originally said
"render the whole YAML file as one Jinja2 text pass, then parse it" — this can't work for
`{{ steps.* }}` references since no step has run at load time. Corrected to: parse YAML
directly first, then resolve fields per-context — `env:`/`workbooks:` once at load time
(env-only), `Step.params`/`if_expr` left raw and resolved per-step during execution
(env + steps). PRD §10.1 and Spec §2.2 updated to match what's actually built, with the
correction explained in place rather than silently changed.

`resolve_value`/`evaluate_condition`/`load` implemented in `excel_runner/core.py`, plus a
custom `_Yaml12BoolLoader` (PyYAML SafeLoader subclass with the yes/no/on/off boolean resolver
removed, per PRD §7's quoting note). All quality gates pass: 53/53 tests, ruff clean,
mypy --strict clean, radon cc clean, vulture clean, **100% branch coverage** on `core.py`.
**Next step:** Build order item 3 (Spec §8) — action registry (`engine.py` §5.1) + a first
vertical slice of trivial actions in `actions.py` (`open`, `save`, `close`, `read_range`,
`write_cell`) to prove the discovery + capability-tag pattern end to end.
**Notes:** `aggregate` and `update_summary_table`'s exact parameters are still explicitly
flagged as open in the PRD — don't block on them, they're late in the build order (Spec §8
items 4 and 10). Tracker below stays function/class-granular even though source files are
consolidated — see Spec §7.

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
| Action registry (`ActionSpec`, `discover_actions`) — `engine.py` §5.1 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `open` action — `actions.py` | ❌ | ❌ | ❌ | ❌ | ❌ |
| `save` action — `actions.py` | ❌ | ❌ | ❌ | ❌ | ❌ |
| `close` action — `actions.py` | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_range` action — `actions.py` | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_cell` action — `actions.py` | ❌ | ❌ | ❌ | ❌ | ❌ |

## Phase 3 — Remaining v1 file-backend actions (Spec §3, §4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| File-backend primitives — `backends.py` §3 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `copy` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_range` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_row` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_table` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `insert_range` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `set_column_width` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `find_headers_row` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `find_row` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `find_column` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `find_columns` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `aggregate` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_metadata` action (properties/cells only) | ❌ | ❌ | ❌ | ❌ | ❌ |

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
| `aggregate` | Flagged for discussion, not resolved |
| `export_pdf` | Backlog |
| AI-authoring inspection actions (`list_sheets`, `describe_sheet`) | Planned, next phase after core engine works (PRD §9) |
| Instance-ownership tagging mechanism across crashes (PRD §6.2.1/§12) | Open question |
| Scratch-directory collision avoidance for concurrent runs (PRD §6.3.1/§12) | Open question |
| CLI / MCP wrapper | Deferred (PRD §3/§5) |
