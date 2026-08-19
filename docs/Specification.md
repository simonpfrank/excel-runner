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

**Built and green.** No whole-file render step — corrected during implementation (see the
PRD §10.1 note on why). Resolution happens per field instead:

1. `load()` reads the YAML file as raw text and parses it directly with a custom
   `yaml.SafeLoader` subclass (`_Yaml12BoolLoader`) that removes PyYAML's YAML-1.1
   `yes`/`no`/`on`/`off` boolean resolver, keeping only `true`/`false` variants — `{{ }}` is
   just string content to a YAML parser, so there's no conflict with parsing first.
2. The merged `env` dict (file's own `env:` block, overridden by `env_overrides` — PRD §6.6)
   becomes the context for resolving `workbooks:` fields immediately, via `resolve_value` —
   e.g. a workbook's `file:` path gets its `{{ env.* }}` references resolved right away, since
   only `env` is available at load time and always will be for that field.
3. `Step.params` and `Step.if_expr` are built directly from the raw parsed step dict and left
   **completely unresolved** — any `{{ }}` they contain (including `{{ steps.* }}`) stays
   literal text until the step actually executes (§6.1), when both `env` and the
   accumulated step-output context exist.

Functions in this module:
- `load(path, env_overrides) -> Workflow` — the pipeline above; the one function `runner.py`
  calls to go from a file path to a typed `Workflow`.
- `resolve_value(value: Any, context: dict) -> Any` — the one resolution primitive, used both
  by `load()` (env-only context) and per-step during execution (env + steps context).
  Implements the "whole field is one `{{ }}` expression → native Python type, else stringify"
  rule from PRD §10.1, with a fast path that skips Jinja entirely for any string containing no
  `{{`/`{%`/`{#` at all. Recurses through dict keys *and* values and list items, so a computed
  dict key (PRD §10.1's finding) resolves naturally as part of resolving the dict it's in.
  Undefined references and syntax errors are caught and re-raised as `ValidationError` with a
  plain-English message and the original exception preserved as `technical_reason` (§2.3).
- `evaluate_condition(if_expr: str, context: dict) -> bool` — for `Step.if_expr`; accepts the
  expression with or without a surrounding `{{ }}` wrapper.

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

All 24 action functions (PRD §7's full catalog — **5 built so far**: `open`, `save`, `close`,
`read_range`, `write_cell`), each with the shape `fn(session: WorkbookSession, **params) ->
ActionResult`.

**Corrected during implementation — `ActionResult` and `WorkbookSession` are defined in
`core.py`, not here or in `engine.py`.** The original plan put `ActionResult` in this section
and `WorkbookSession` in §5.2, but that creates a circular import: `engine.py`'s registry must
import `actions.py` to discover its functions, and `actions.py`'s functions are typed against
`WorkbookSession`/`ActionResult` — if either type lived in `engine.py`, `actions.py` would need
to import `engine.py` right back. `core.py` depends on neither, and both already depend on it,
so that's where the shared types live now: `core.py` (data) → `{backends.py, actions.py}` →
`engine.py` → `runner.py`, a clean line, no cycle. `ActionResult` is still a small dataclass
(`status: Literal["success","error"]`, `output: dict`, `error: ErrorDetail | None`) — this is
still what makes PRD §10.4's "output is always a keyed object" rule mechanical rather than a
convention each action has to remember: `ActionResult` is the return type, full stop.

**Also corrected: no `workbook` parameter on any action function**, even though `workbook:` is
a required field on every step in the YAML (PRD §7/§11). The (not-yet-built) runner resolves
`workbook` into the `session` it passes in before calling the action — passing both would be
the same information twice. `workbook` still exists as a YAML field; it's just consumed before
the Python function is ever called, not forwarded to it.

**Also corrected: actions call `backends.py` functions directly, passing `session.handle`** —
not through a `session.<something>` indirection as originally sketched. Since an action's
capability tag already fixes which backend it will only ever run against, there's no runtime
branching to hide behind an indirection layer; the action just calls the matching (unprefixed
for file, `com_`-prefixed for COM) `backends.py` function. Backend choice still isn't the
action's decision — it's fixed once, by the capability tag, not decided per-call.

**Each action registers its capability via a decorator** (`@file_action`/`@com_action`, defined
in `core.py` alongside `ACTION_CAPABILITIES`, a plain `name -> capability` dict the decorators
populate) rather than by stamping an attribute onto the function object — keeps mypy --strict
clean (no dynamic-attribute `type: ignore` noise) and keeps registration trivially
introspectable for `engine.py`'s `discover_actions()` (§5.1).

**The 5 built actions have a deliberately reduced param surface vs. the full PRD §7 catalog**,
each documented in its own docstring: `open` omits `update_links` (no effect without a live
Excel session — COM, a later phase) and a `mode` override (depends on read/write inference that
tier-2 validation, §5.4, doesn't exist yet to override); `read_range` omits `as: formulas`
(depends on which `data_only` flag the workbook was opened with — a session-level decision, §5.4
again). These are scope boundaries for this increment, not permanent cuts — they get added back
once the machinery they depend on exists.

This is the largest file in the package (an estimated 900–1200 lines across 24 functions) and
the one touched most often. Two things keep it navigable without adding source files:
- **Tests don't have to collapse to match.** `tests/unit/actions/test_read_range.py`,
  `test_write_row.py`, etc. stay one-file-per-action even though the source doesn't — most of
  the per-action navigability comes back on the test side, where it matters most for TDD (§7).
- Functions are ordered in the file to match PRD §7's table order (basic → data → structure →
  lookup → aggregate → links → COM), with a one-line comment banner per group, so the physical
  layout still mirrors the catalog even without file boundaries doing it.

**14 built so far**: `open`, `save`, `close`, `copy`, `read_range`, `read_metadata`
(properties/cells sub-cases), `write_cell`, `write_range`, `write_row` (base + positional
modes), `insert_range` (whole-row/whole-column only), `set_column_width`, `find_headers_row`,
`find_row`, `find_column`, `find_columns`. All green, 100% branch coverage.

**Error-handling policy, established while building this batch**: an action returns
`ActionResult(status="error", error=ErrorDetail(...))` for an outcome that's a normal,
anticipated result of trying — a search that legitimately finds nothing (`find_row`,
`find_column`, `find_headers_row`), or a documented not-yet-supported case
(`insert_range`'s partial range). It *raises* `ActionExecutionError`/`ValidationError` for a
genuine authoring mistake the action can't proceed from at all (`write_row`'s positional mode
called without `start_column`; `read_metadata`'s `cells` target called without `sheet`/`cells`).
The line: would a workflow author reasonably want to branch on this with `if:` (structured
result), or is it just wrong and should stop the run (exception)? Deliberately not
wrapping every possible exception into a result — that's the "defensive fallback layers" PRD §1
names as a root cause to avoid, not a pattern to bring back under a different name.

**Deferred, with reasons found during implementation**:
- `write_table`, `write_row`'s by-header mode, and `aggregate` all need **step-output context**
  a standalone action call doesn't have — `write_table`'s `source: [step_ids]` and `aggregate`'s
  `source` are step-id references (not templated values `resolve_value` would catch), and
  `write_row`'s `headers_from` is the same shape. All three become buildable once `runner.py`
  threads accumulated step outputs through to action calls (build order item 7). `aggregate`
  was already flagged "discuss when we get to it" in PRD §7/§11.17 — this compounds that, it
  doesn't newly block it.
- `read_links` — **empirically downgraded**, not just theoretically uncertain. A quick spike
  during implementation: writing a formula string containing an external reference
  (`=[Source.xlsx]Sheet1!A1`) via openpyxl and reopening the file leaves
  `workbook._external_links` empty — openpyxl never creates the underlying relationship from a
  formula, it only reflects one already present in a file created by real Excel. Reading may
  still work *given* such a file, but verifying that needs a real Excel-generated fixture (or
  manual zip/XML surgery to fabricate one) — a real task, not a quick addition. Moved into the
  same deferred bucket as `write_links` (PRD §7/§8 updated to match). Any notes from a future
  spike go in the private, gitignored file per §0 — not named here.
- `copy` needed a genuinely different signature — two `WorkbookSession` params
  (`session`=source, `target`=target), since its YAML shape has two nested workbook refs
  (`source: {...}`, `target: {...}`) rather than one flat `workbook:` field. The action itself
  is built and tested; `runner.py` will need matching special-case wiring to resolve *two*
  sessions for this one action, not the usual one (build order item 7).
- `read_metadata`'s `cells` sub-case needed a `sheet` parameter the original PRD §7 catalog
  didn't list — reading specific cells needs to know which worksheet they're on, same as every
  other cell-addressing action. PRD §7 updated.

## 5. Engine layer — `engine.py`

Everything that prepares and manages a run *before and during* action dispatch — registry
discovery, session/workbook lifecycle, the scratch-copy execution model, and both validation
tiers. Four sub-concerns, one file, because they're all "run-preparation and run-state," used
together by `runner.py` (§6) and nothing else.

### 5.1 Action registry — **built**

Mirrors a proven pattern (name checked at the pattern level, not copied — see §0) of scanning
for typed, docstringed functions and generating a schema from each one's signature:

```python
@dataclass(frozen=True)
class ActionSpec:
    name: str
    fn: Callable[..., ActionResult]
    capability: Literal["file", "com", "depends_on_param"]   # "depends_on_param": read_metadata
    param_schema: dict[str, Any]        # derived from fn's signature (excludes `session`)
```

(`ActionResult` itself is imported from `core.py` — see §4's correction.)

`capability="depends_on_param"` is a **named, single exception**, not a general mechanism —
only `read_metadata` uses it (PRD §7: file for `properties`/`cells`, COM for `textboxes`), and
isn't built yet. `runner.py` (not built yet) will check for this literal case explicitly rather
than building a generic capability-resolution feature for one action.

`discover_actions(module)` scans a module with `inspect.getmembers`, keeping only functions
with an entry in `core.py`'s `ACTION_CAPABILITIES` dict (populated by the `@file_action`/
`@com_action` decorators, §4) — not every function in `actions.py`, just the tagged ones.
`param_schema` is derived from the function's signature, skipping `session` and marking any
parameter with no default as required.

### 5.2 Session management — **`WorkbookSession` data shape built; `SessionManager` not yet**

```python
@dataclass
class WorkbookSession:      # actually lives in core.py — see §4's correction
    name: str
    backend: Literal["file", "com"]
    handle: Any                      # openpyxl Workbook | xlwings Book
    path: str                        # added during implementation — see below
    mode: Literal["read_only", "read_write"]
    scratch_path: Path | None = None # None until the scratch-copy model (§5.3) is built
    dirty: bool = False

class SessionManager:               # not built yet
    def get_or_open(self, name: str) -> WorkbookSession: ...
    def promote_to_com(self, name: str) -> WorkbookSession: ...   # file -> com, mid-run
    def commit_all(self) -> None: ...     # delegates to the scratch-copy manager, §5.3
    def close_all(self) -> None: ...      # always runs — see runner.py's try/finally, §6.1
```

**`path: str` was missing from the original sketch and had to be added** — `save` needs to know
*where* to save to, and there's no way to derive that without the session carrying its own
current path. It's `path`, not reused-as-`scratch_path`, because `scratch_path` specifically
means "None until scratch-copy staging exists" (§5.3) — `path` is always concrete: the real file
path today, and wherever `scratch_path` points once §5.3 lands and starts routing writes there
instead.

`SessionManager` — the actual multi-workbook lifecycle tracker (lazy-open, promotion, close-all)
— isn't built yet; that's what makes actions.py's five built actions take a `WorkbookSession`
directly rather than a `SessionManager`. Once built, `SessionManager` is still the only thing
that calls into `backends.py` (§3) for the *lifecycle* operations (open/promote/close) — but per
§4's correction, actions call `backends.py` directly for their own per-action operations
(read/write/etc.), not through a session indirection.

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

1. `core.py` §2.1/§2.3 — data model + error types. Pure dataclasses, no I/O. **Done.**
2. `core.py` §2.2 — loading/templating (`load`, `resolve_value`, `evaluate_condition`),
   unit-testable with no real workbooks (temp YAML files only). **Done.**
3. `engine.py` §5.1 (registry) + a first vertical slice of trivial actions in `actions.py`
   (`open`, `save`, `close`, `read_range`, `write_cell`), plus `backends.py`'s first 5
   file-backend primitives, to prove the discovery + capability-tag pattern end to end before
   building the other ~15 file-backend actions. **Done** — surfaced the corrections recorded in
   §4/§5.1/§5.2 (shared types moved to `core.py`, `workbook` param dropped from action
   signatures, actions call `backends.py` directly, `WorkbookSession` needed a `path` field).
4. `backends.py`'s remaining file-backend functions, filled out alongside the remaining v1
   file-backend actions in `actions.py`. **Done, for what's cleanly buildable now** — `copy`,
   `write_range`, `write_row` (base + positional modes), `insert_range` (whole-row/column),
   `set_column_width`, `find_headers_row`, `find_row`, `find_column`, `find_columns`,
   `read_metadata`'s file sub-case. `write_table`, `aggregate`, `write_row`'s by-header mode,
   and `read_links` deferred with concrete reasons — see §4's "Deferred, with reasons found
   during implementation."
5. `engine.py` §5.2/§5.3 — the `SessionManager` class (multi-workbook lifecycle: lazy-open,
   promotion, close-all — the `WorkbookSession` data shape itself is already built, §5.2) and
   scratch-copy layer, now that there are real actions to run through it.
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
