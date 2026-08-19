# excel_runner — Technical Specification

Built from `docs/PRD.md` (the source of truth for *what* and *why* — this document is *how*).
Package name: `excel_runner`.

## 0. Sourcing policy

A prior tool this project supersedes may be **consulted** during implementation for
inspiration — some of its functions are genuinely worth a look for *how* to approach a problem.
What's not allowed:

- **No copying.** All code in this repo is written new. Not pasted, not closely paraphrased,
  and not structurally mirrored (module layout, class hierarchy, function boundaries) — the
  entire point of a fresh build is to not inherit that prior tool's structural issues
  (COM-everywhere, string-mini-language config, defensive fallback layers — see PRD §1 for what
  those were, described in first-principles terms rather than tied to any specific prior code).
- **No references in tracked material.** The prior tool is never named — not by name, not by
  path — anywhere in this repo's committed documentation, docstrings, code comments, or commit
  messages. `docs/PRD.md` and this document are already written that way, and this section is
  deliberately written the same way.
- **Any note that does need to reference it stays in a private, gitignored file** under
  `docs/research/` (already excluded — see `.gitignore`) — that's the one place to jot down a
  real-world behavior/gotcha it ran into, or "look at function Y for an idea," during
  implementation. If something learned there ends up shaping a decision in a tracked doc, the
  tracked doc gets the first-principles rationale, never a pointer back to that file or what it
  contains.

## 1. Package layout

Consolidated into 5 files by layer (not by concern-per-file) — deliberately reduced from a
finer-grained split, on the reasoning that file-count reduction here is a pure
human-navigability tradeoff (auto-discovery, capability tagging, and AI-tool-schema generation
all work identically regardless of which file a function lives in) and each of the 5 tells one
coherent story:

```
excel_runner/
  __init__.py       # re-exports the public surface from runner.py — see §6.3
  core.py            # data model, errors, loading & templating — the pure-logic layer (§2)
  backends.py         # openpyxl + xlwings primitives, owned-instance tracking (§3)
  actions.py           # all 24 action functions (§4)
  engine.py             # registry, session, scratch-copy, both validation tiers (§5)
  runner.py              # orchestration loop, audit logging, public API (§6)

tests/
  unit/            # test files stay granular even though source doesn't — see §7
  integration/      # real files/openpyxl, real Excel where COM is exercised
  data/              # fixture workbooks
```

No `cli.py` in v1 — the CLI/agent wrapper is explicitly deferred (PRD §3/§5); `runner.py`'s
public surface (§6.3) is built so that layer is thin whenever it's added.

## 2. Core layer — `core.py`

The pure-logic layer: no file I/O beyond reading the YAML source itself, no Excel, nothing that
needs mocking in a unit test. Three responsibilities live here because they're all "turn
config text into typed, validated Python objects" — genuinely one story, not an arbitrary
merge.

### 2.1 Data model

Plain dataclasses, not pydantic (project convention: pydantic only when strictly necessary;
here it isn't — the field-level error messages required by PRD §9.1 need custom wording per
case anyway, which a generic validator library doesn't buy us).

```python
@dataclass(frozen=True)
class WorkbookRef:
    name: str                       # logical name, the registry key
    file: str                       # path, may contain {{ env.* }} templating
    create_if_missing: bool = False
    template: str | None = None     # logical name of another WorkbookRef

@dataclass(frozen=True)
class Step:
    id: str
    action: str
    params: dict[str, Any]          # raw, schema-validated against the action's ActionSpec
    if_expr: str | None = None      # raw Jinja2 boolean expression, unevaluated at parse time

@dataclass(frozen=True)
class Workflow:
    env: dict[str, Any]
    workbooks: dict[str, WorkbookRef]
    steps: tuple[Step, ...]
```

`Step.params` stays a validated-but-generic `dict` rather than one dataclass per action —
24 near-duplicate param dataclasses would violate "avoid duplicate functionality" for no real
type-safety gain, since each action's registered function signature (via the registry in
`engine.py`, §5.1) already is the authoritative shape. Validation cross-checks `params` against
that signature at tier 1 (§5.4).

### 2.2 Loading and templating pipeline

Pipeline, per PRD §10.1 (render-then-parse):

1. Read the YAML file as raw text.
2. `render(raw_text, context) -> str` — render the **entire file** through Jinja2 first.
   `context` at this stage is just `{"env": <merged env dict>}` (§6.6 of the PRD's
   `env_overrides` merged over the file's own `env:` block) — step outputs don't exist yet, so
   any `{{ steps.* }}` expression in the file is *not* rendered here; it stays literal text for
   step 4.
3. `yaml.safe_load(...)` (with a YAML-1.2-core-schema loader, per PRD §7's quoting note, so
   `on`/`off`/`yes`/`no` are never silently coerced) the rendered text into a raw `dict`.
4. `build_workflow(raw_dict) -> Workflow` — structural parse into the §2.1 dataclasses.
   `Step.if_expr` and any `{{ steps.* }}` references inside `Step.params` are **not** resolved
   here — they're resolved per-step, during execution (§6.1), once the referenced step has
   actually run and its output exists.

Functions in this module:
- `load(path, env_overrides) -> Workflow` — the four steps above, wired together; the one
  function `runner.py` calls to go from a file path to a typed `Workflow`.
- `render(text: str, context: dict) -> str` — the file-level pre-parse render (step 2 above).
- `resolve_value(value: Any, context: dict) -> Any` — per-step resolution during execution;
  implements the "whole field is one `{{ }}` expression → native Python type, else stringify"
  rule from PRD §10.1. Recurses through nested dicts/lists in a step's `params` so a computed
  dict key (PRD §10.1's finding) is resolved before the dict itself is used.
- `evaluate_condition(if_expr: str, context: dict) -> bool` — for `Step.if_expr`.

### 2.3 Error types

```python
@dataclass(frozen=True)
class ErrorDetail:
    message: str            # plain English — PRD §6.8/§9.1's bar
    technical_reason: str   # original exception type + message, never shown as the headline
    field: str | None = None
    suggestion: str | None = None

class ExcelRunnerError(Exception): ...
class ValidationError(ExcelRunnerError): ...
class ActionExecutionError(ExcelRunnerError): ...
```

`engine.py`'s tier-1 validator (§5.4) builds `ValidationError` instances — PRD §9.1's four
example messages — via small, purpose-named functions (e.g. `_check_field_is_list`,
`_check_step_reference_exists`) rather than string-formatting inline at the call site, so each
is independently unit-testable against its own bad-input fixture.

## 3. Backends layer — `backends.py`

Everything that actually talks to openpyxl or xlwings/COM lives here — plain functions, not
classes, one function per primitive operation (`open_workbook`, `read_range`, `write_cell`,
`copy_range`, ...), so swapping or testing either side never requires touching action code
(PRD §6.1's backend-invisibility goal). Naming convention keeps the two sides unambiguous
within the single file: file-backend functions are unprefixed; COM-backend functions carry a
`com_` prefix (`com_open_workbook`, `com_recalculate`, `com_run_macro`, `com_refresh_links`,
`com_write_links`, `com_read_textbox`) — enforced by naming and a section-comment banner now
that a directory boundary no longer does it. If this file grows past a size where that
convention stops being enough to navigate by, split by concern then (not pre-split now, since
the real size is unknown before code exists).

### 3.1 Owned-instance tracking (PRD §6.2.1)

```python
class OwnedInstanceRegistry:
    def spawn(self) -> InstanceHandle: ...     # always a NEW xw.App, never xw.apps.active
    def close_owned(self) -> None: ...          # only ever acts on self._owned
```

Holds the run's own set of spawned `xw.App` PIDs. The `SessionManager` (`engine.py`, §5.2) asks
this registry for an App instance rather than ever calling `xw.App`/`xw.Book` directly — that
keeps "never touch an instance we didn't spawn" (PRD §6.2.1) enforced in one place instead of
scattered through action code. The cross-run "recognize an orphaned instance from a *previous*
crashed run" mechanism is explicitly not designed yet — PRD §12 open item — this class is the
seam where that logic will attach once designed, not a placeholder to guess at now.

## 4. Actions layer — `actions.py`

All 24 action functions (PRD §7's full catalog), each with the shape
`fn(session: WorkbookSession, **params) -> ActionResult`, where `ActionResult` is a small
dataclass (`output: dict`, `status: Literal["success","error"]`, `error: ErrorDetail | None`) —
this is what makes PRD §10.4's "output is always a keyed object" rule mechanical rather than a
convention each action has to remember to follow: `ActionResult` is the return type, full stop.

Actions never import `backends.py` functions directly by name scattered through the file's own
logic beyond calling them — they call `session.<something>` (from `engine.py`, §5.2) which in
turn delegates to `backends.py`, so backend choice stays centralized there rather than
duplicated per action.

This is the largest file in the package (an estimated 900–1200 lines across 24 functions) and
the one touched most often. Two things keep it navigable without adding source files:
- **Tests don't have to collapse to match.** `tests/unit/actions/test_read_range.py`,
  `test_write_row.py`, etc. stay one-file-per-action even though the source doesn't — most of
  the per-action navigability comes back on the test side, where it matters most for TDD (§7).
- Functions are ordered in the file to match PRD §7's table order (basic → data → structure →
  lookup → aggregate → links → COM), with a one-line comment banner per group, so the physical
  layout still mirrors the catalog even without file boundaries doing it.

## 5. Engine layer — `engine.py`

Everything that prepares and manages a run *before and during* action dispatch — registry
discovery, session/workbook lifecycle, the scratch-copy execution model, and both validation
tiers. Four sub-concerns, one file, because they're all "run-preparation and run-state," used
together by `runner.py` (§6) and nothing else.

### 5.1 Action registry

Mirrors a proven pattern (name checked at the pattern level, not copied — see §0) of scanning
for typed, docstringed functions and generating a schema from each one's signature:

```python
@dataclass(frozen=True)
class ActionSpec:
    name: str
    fn: Callable[..., ActionResult]
    capability: Literal["file", "com", "depends_on_param"]   # "depends_on_param": read_metadata
    param_schema: dict[str, Any]        # derived from fn's signature + docstring
```

`capability="depends_on_param"` is a **named, single exception**, not a general mechanism —
only `read_metadata` uses it (PRD §7: file for `properties`/`cells`, COM for `textboxes`).
`runner.py` checks for this literal case explicitly rather than building a generic
capability-resolution feature for one action.

`discover_actions()` runs once (import time or lazily, cached), introspecting `actions.py` to
build the module-level registry that `runner.py`'s public surface (§6.3) and the tier-1
validator (§5.4) both read from.

### 5.2 Session management

```python
@dataclass
class WorkbookSession:
    name: str
    backend: Literal["file", "com"]
    handle: Any                      # openpyxl Workbook | xlwings Book
    mode: Literal["read_only", "read_write"]
    scratch_path: Path | None        # None if opened read-only in place (no scratch copy)
    dirty: bool = False

class SessionManager:
    def get_or_open(self, name: str) -> WorkbookSession: ...
    def promote_to_com(self, name: str) -> WorkbookSession: ...   # file -> com, mid-run
    def commit_all(self) -> None: ...     # delegates to the scratch-copy manager, §5.3
    def close_all(self) -> None: ...      # always runs — see runner.py's try/finally, §6.1
```

`SessionManager` is the only thing that calls into `backends.py` (§3) — actions never do, they
call `session.<something>` instead, so backend choice (PRD §6.1's "never a user choice," and
here also "never an *action-code* choice") stays centralized in one place.

### 5.3 Scratch-copy execution model

Implements PRD §6.3.1 directly:

```python
class ScratchManager:
    def stage(self, workbook_refs: Iterable[WorkbookRef]) -> dict[str, Path]: ...
       # copies only workbooks the dry-run pass (§5.4) marked read-write;
       # read-only sources are opened in place, never staged
    def commit(self, name: str, scratch_path: Path, real_path: Path) -> None: ...
       # temp-path-then-os.replace, per workbook, atomic
    def commit_all(self) -> None: ...
    def cleanup(self, keep_on_failure: bool = True) -> None: ...
```

`SessionManager.commit_all()`/`close_all()` delegate here rather than duplicating the
atomic-write logic. `ScratchManager` has no knowledge of openpyxl/xlwings — it operates on
plain file paths, so both backends stage/commit through the same code path (PRD §6.3.1's "COM
steps operate on the scratch copy too").

### 5.4 Validation — two tiers

Matching PRD §9/§9.1, kept as clearly separated functions within the file since they run at
different times and check different things:

- **Tier 1 (static schema validation)** — runs immediately on the raw parsed `Workflow`, no
  workbook access. Checks: every `Step.action` exists in the registry (§5.1); every param
  name/type matches that action's signature; every `steps.<id>` reference in a `{{ }}`
  expression points at a step id that exists *and* appears earlier in the list; `range`/
  named-range syntax is at least well-formed. Produces the specific, corrective error format
  from PRD §9.1 — each check is a small, individually testable function, not one large
  validator.
- **Tier 2 (dry-run / step-graph validation)** — still no real workbook access, but now reasons
  over the whole step graph together: for each `WorkbookRef`, is it ever a `target`? (drives
  read-only vs. read-write inference, PRD §6.3) Does every workbook a step references appear in
  `workbooks:`? Produces an `ExecutionPlan` (per-workbook mode, which ones need scratch
  staging) that `runner.py` (§6.1) and `ScratchManager` (§5.3) consume directly — this *is* the
  "plain-English execution plan" PRD §9 promises an agent/user can sanity-check before a real
  run.

## 6. Runner layer — `runner.py`

The composition root, the audit trail, and the one public contract — grouped because together
they're "the layer that actually executes a run and is safe for other code to depend on."

### 6.1 Orchestration

```python
def run_workflow(path: str | Path, env_overrides: dict | None = None) -> RunResult:
```

One linear sequence (per project convention: a composition root doesn't need splitting just to
hit a line count — see AGENTS.md), wrapped in `try`/`finally` for the crash-safety guarantee
(PRD §6.3):

1. `core.load(path, env_overrides)` → `Workflow`
2. Tier-1 validation (§5.4) → raises with PRD §9.1-style errors on failure, before anything
   else runs
3. Tier-2 validation (§5.4) → `ExecutionPlan`
4. `session = SessionManager(); scratch = ScratchManager()`
5. `scratch.stage(...)` for every workbook the plan marked read-write
6. For each `Step` in order:
   a. if `step.if_expr` set, `core.evaluate_condition(...)` against accumulated step outputs —
      skip the step (record `status: "skipped"`) if false
   b. `core.resolve_value(...)` over `step.params`
   c. look up `ActionSpec` in the registry (§5.1), resolve capability (promoting the workbook's
      session if needed), call the action function
   d. record the `ActionResult` into the run's step-output context (for later `{{ }}`
      references) and into the audit log (§6.2)
7. On full success: `session.commit_all()` (→ `scratch.commit_all()`), then
   `scratch.cleanup(keep_on_failure=False)`
8. `finally`: `session.close_all()` unconditionally (file handles *and* any COM instances via
   `OwnedInstanceRegistry.close_owned()`, §3.1) — runs whether step 6 succeeded, raised, or the
   process was interrupted, satisfying PRD §6.3's hard requirement. On failure, scratch copies
   are deliberately **not** cleaned up (PRD §6.3.1) — left as the recovery artifact.

`RunResult` = `{status, step_results: list[StepResult], audit_log_path}`.

### 6.2 Audit logging

```python
class AuditLogger:
    def record_step(self, step: Step, result: ActionResult, started_at, ended_at) -> None: ...
```

One JSON object per line (JSONL) written to the scratch/run directory, per PRD §6.7. The
orchestration loop (§6.1) calls this once per step, success or failure, before deciding whether
to continue. Not a `logging`-module handler — deliberately a separate, structured artifact
(PRD §6.7 explains why).

### 6.3 Public API surface

The only symbols other Python code (or a future CLI/MCP wrapper, PRD §5) should import,
re-exported from the package's `__init__.py`:

```python
from excel_runner import run_workflow, RunResult, StepResult
from excel_runner import Workflow, Step, WorkbookRef   # for programmatic construction
from excel_runner import list_actions                   # -> tuple[ActionSpec, ...]
```

Everything else in the package tree is an implementation detail and may change without notice;
this surface is the versioned contract (PRD §3/§9/§12). `list_actions()` exposing `ActionSpec`
(name/docstring/param_schema) is what a later MCP/CLI wrapper would iterate over to generate
its own tool definitions — the "close to free" schema reuse PRD §6.1 promises, made concrete
here as one function.

## 7. Testing approach

- **`tests/unit/`** stays fine-grained even though `excel_runner/` doesn't — one test file per
  action (`tests/unit/actions/test_read_range.py`, ...), per class (`test_session_manager.py`,
  `test_scratch_manager.py`, ...), per concern within `core.py` (`test_templating.py`,
  `test_schema.py`, ...). Backends are mocked in action unit tests; `backends.py`'s own file
  side is unit-tested against a real in-memory openpyxl workbook (openpyxl needs no live Excel,
  so this isn't a "mock" in the disallowed sense — it's exercising a real dependency that's
  cheap and fast, same spirit as the project convention's "no mocks in integration tests," just
  applied at the unit level where that's affordable).
- **`tests/integration/`** — zero mocks, per project convention. File-backend actions run
  against real fixture workbooks in `tests/data/`. COM-backed actions/tests are marked
  `@pytest.mark.skipif` on platform/Excel-availability (PRD §4's "test what's testable on macOS
  now, finish on Windows later") — they skip cleanly, never mock the COM layer.
- A dedicated integration test deliberately crashes a run mid-step and asserts: no orphaned
  Excel process, no file lock on the original workbook, real files unmodified, scratch copies
  present — this is the concrete test for PRD §6.3/§6.3.1's crash-safety requirement, not just
  a design note.

## 8. Build order

Matches PRD §8's v1/later split, sequenced for TDD (pure logic first, I/O-heavy and
platform-dependent pieces last). Note the file each item lands in no longer maps 1:1 to the
increment — several increments land in the same file, built and tested one function/class at a
time within it.

1. `core.py` §2.1/§2.3 — data model + error types. Pure dataclasses, no I/O.
2. `core.py` §2.2 — loading/templating (render/resolve/condition), unit-testable with no files
   or workbooks at all.
3. `engine.py` §5.1 (registry) + a first vertical slice of trivial actions in `actions.py`
   (`open`, `save`, `close`, `read_range`, `write_cell`) to prove the discovery + capability-tag
   pattern end to end before building the other ~15 file-backend actions.
4. `backends.py`'s file-backend functions, filled out alongside the remaining v1 file-backend
   actions in `actions.py` (PRD §7/§8: `copy`, `write_range`, `write_row`, `write_table`,
   `insert_range`, `set_column_width`, `find_headers_row`, `find_row`, `find_column`,
   `find_columns`, `aggregate`, `read_links`, `read_metadata`'s file sub-case).
5. `engine.py` §5.2/§5.3 — session + scratch-copy layer, now that there are real actions to run
   through it.
6. `engine.py` §5.4 — both validation tiers.
7. `runner.py` §6.1/§6.2 — first end-to-end real run of a multi-step file-backend workflow.
8. `runner.py` §6.3 — the public surface, once there's a working engine underneath it to expose.
9. **Later phase (Windows-dependent COM work)**: `backends.py`'s COM-backend functions and
   `OwnedInstanceRegistry` (§3.1), and the COM actions in `actions.py` (`recalculate`,
   `run_macro`, `refresh_links`, `write_links`, `read_metadata`'s textbox sub-case).
10. **Deferred/flagged, per PRD**: `update_summary_table`'s real parameters, the `aggregate`
    discussion, `export_pdf`, the AI-authoring inspection actions (PRD §9: `list_sheets`,
    `describe_sheet`).

`docs/Progress_Tracker.md` tracks each item above against the project's standard Component /
Unit Tests / Code / Integration Tests / Results columns, at function/class granularity — not
collapsed to match the 5-file source layout.
