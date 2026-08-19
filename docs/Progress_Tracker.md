# excel_runner — Progress Tracker

## Last Session (2026-08-19)
**Status:** In Progress
**Working on:** Build order item 1 (Spec §8) — `core.py` data model + error types, TDD.
Project skeleton created (`pyproject.toml`, `.venv` via `uv` — Homebrew's Python 3.14
`ensurepip` was broken, used `uv venv --python 3.12` instead). 21 unit tests written first
(red), then `excel_runner/core.py` implemented (green): `WorkbookRef`, `Step`, `Workflow`,
`ErrorDetail`, `ExcelRunnerError`, `ValidationError`, `ActionExecutionError`. All quality gates
pass: ruff clean, mypy --strict clean, radon cc clean (no C+), vulture clean (added
`vulture_whitelist.py` for the standard dataclass-field false-positive pattern), 100% branch
coverage on `core.py`.
**Next step:** Loading/templating pipeline (Spec §2.2) — `render`, `resolve_value`,
`evaluate_condition`, `load` — still in `core.py`. Needs `pyyaml` + `jinja2` added as real
(non-dev) dependencies.
**Notes:** No git repo initialized yet in this directory — flag if that should happen.
`aggregate` and `update_summary_table`'s exact parameters are still explicitly flagged as open
in the PRD — don't block on them, they're late in the build order (Spec §8 items 4 and 10).
Tracker below stays function/class-granular even though source files are consolidated — see
Spec §7.

## Status legend
❌ Not Done · 🟡 In Progress · ✅ Done — Results: ✅ Pass · ❌ Fail · ⏭️ N/A

## Phase 1 — `core.py`: data model, errors, templating (Spec §2)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Data model dataclasses (`WorkbookRef`, `Step`, `Workflow`) — §2.1 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Error types (`ErrorDetail`, exception classes) — §2.3 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Loading pipeline (`load`) — §2.2 | ❌ | ❌ | ⏭️ | ❌ | ⏭️ |
| Templating (`render`, `resolve_value`, `evaluate_condition`) — §2.2 | ❌ | ❌ | ⏭️ | ❌ | ⏭️ |

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
