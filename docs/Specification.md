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
  (always-live-Excel, string-mini-language config, defensive fallback layers — see PRD §1 for what
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
  actions.py           # all action functions (§4)
  engine.py             # registry, session, scratch-copy, both validation tiers (§5)
  runner.py              # orchestration loop, audit logging, public API (§6)
  cli.py                  # thin argument-parsing/JSON-formatting wrapper over runner.py (§6.4)
  __main__.py              # `python -m excel_runner` entry point, delegates to cli.py

tests/
  unit/            # test files stay granular even though source doesn't — see §7
  integration/      # real files/openpyxl, real Excel where xlwings is exercised
  data/              # fixture workbooks
```

**Correction**: `cli.py` was originally deferred out of v1 ("no CLI/agent wrapper" — PRD §3/§5),
but was built once a real, concrete driving need showed up (invoking a workflow from a 3rd party workflow system as an external process) — see §6.4. It stays thin by design, exactly as this section
originally anticipated: argument parsing and JSON result formatting only, no logic of its own.

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

**A real bug found while building `runner.py` (§6.1): a dict field literally named `"values"`
(exactly `read_range`'s own output key, PRD §10.4) was shadowed by Python dict's real
`.values()` method.** Jinja2's default attribute resolution tries `getattr(obj, name)` before
`obj[name]` — and every dict has real `.keys()`/`.values()`/`.items()`/etc. methods — so
`{{ steps.x.output.values }}` returned the bound method, not the output dict's `"values"`
entry. Fixed generically, not by renaming around the one collision: `_ENV` is now a
`_DictItemFirstEnvironment` subclass overriding `getattr` to try item access first, falling
back to real attribute access only when the object isn't subscriptable. Correct here — not a
hack — because every value flowing through these templates (`env`, `steps`, an action's
output) is a plain dict or list, never an object whose real attributes should take priority
over its data.

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

Everything that actually talks to openpyxl or xlwings (live Excel) lives here — plain functions,
not classes, one function per primitive operation (`open_workbook`, `read_range`, `write_cell`,
`copy_range`, ...), so swapping or testing either side never requires touching action code
(PRD §6.1's backend-invisibility goal). Naming convention keeps the tiers unambiguous within the
single file — corrected from an earlier two-tier "COM" catch-all that was inaccurate on macOS
(no COM exists there, only Apple Events; xlwings abstracts the difference, PRD §4):
- File-backend functions (openpyxl) are unprefixed.
- `xlw_`-prefixed functions use xlwings' portable, cross-platform API — the default live-Excel
  case.
- `com_`-prefixed functions are the genuine exception: xlwings' `.api` escape hatch for
  something only the raw, Windows-only COM object can do (PRD §7's `recalculate`
  full/full_rebuild modes are the known example) — not the default, only where actually needed.

Enforced by naming and a section-comment banner now that a directory boundary no longer does it.
Which specific future primitives land in which tier isn't fully decided until each is actually
built (most will be `xlw_`; `recalculate` is the known case needing both, split by `mode`) — not
asserted as a fixed list here. If this file grows past a size where the naming convention stops
being enough to navigate by, split by concern then (not pre-split now, since the real size is
unknown before code exists).

### 3.1 Owned-instance tracking (PRD §6.2.1) — **built**

```python
class OwnedInstanceRegistry:
    def spawn(self, visible: bool = False) -> xw.App: ...  # always a NEW xw.App, never xw.apps.active
    def close_owned(self) -> None: ...                      # only ever acts on self._owned
    pids: tuple[int, ...]                                    # property — the audit-trail-ready PID set
```

Holds the run's own set of spawned `xw.App` instances, keyed by PID (`dict[int, xw.App]`, not
just a bare PID set — the actual object reference is what `close_owned()` needs to act on). The
(future, not-yet-built) `SessionManager` xlwings promotion will ask this registry for an App
instance rather than ever calling `xw.App`/`xw.Book` directly — that keeps "never touch an
instance we didn't spawn" (PRD §6.2.1) enforced in one place instead of scattered through action
code. `close_owned()` mirrors `SessionManager.close_all()`'s `ExceptionGroup` pattern exactly:
every owned instance gets a quit attempt regardless of an earlier one failing. The cross-run
"recognize an orphaned instance from a *previous* crashed run" mechanism is still explicitly not
designed — PRD §12 open item — this class is the seam where that logic will attach once
designed, not a placeholder to guess at now.

**Tested for real against a locally-spawned Excel instance, no mocks** (`tests/unit/
test_owned_instance_registry.py`), gated by a `requires_excel` skip marker (`tests/unit/
conftest.py`) so the suite degrades cleanly on a machine without Excel rather than failing.
Found two genuine, empirical facts about macOS Excel automation while writing these tests —
neither is a bug in this class, both are documented here so later xlwings work doesn't rediscover
them the hard way:
- **`app.quit()` is asynchronous.** The call returns before the underlying process has actually
  terminated (~0.5s observed locally). A test — or any future code — checking "is this PID gone
  yet" immediately after `quit()` returns is racing it, not verifying it; needs a short poll.
- **Quitting an already-dead instance doesn't fail predictably.** It raises `-600 Application
  isn't running` when it's the *only* Excel process on the machine, but silently no-ops when
  another owned instance is still alive — confirmed this does **not** cross-target and
  accidentally kill the other live instance (PRD §6.2.1's core safety concern holds), it's just
  an inert no-op. Not something `close_owned()`'s tests can rely on for a real double-quit
  failure case — mirrors `SessionManager`'s own precedent (openpyxl's `close()` is also a no-op
  on a second call), so the "one failure doesn't block the rest" test uses a fake exploding
  stand-in instead of a genuine double-quit, same justification.

Also confirmed empirically (the reason build order item 10 can even start on macOS): spawning a
dedicated App, opening/adding a workbook, and reading/writing cell values all work reliably here
via Apple Events. `save()` specifically does not — `Parameter error (-50)` on
`save_workbook_as`, reproduced consistently (headless and visible, simple and complex paths, no
dialog involved) and matches multiple open xlwings GitHub issues, not something specific to this
machine. Doesn't block v1 (PRD §4 already treats Windows as the real target and macOS as
"test what's testable now") — it does mean anything routing through xlwings' `save()` (and possibly
`recalculate`/`run_macro`/`refresh_links`/`write_links` — not yet individually checked) needs its
tests gated behind Windows access, per this section's skip-don't-mock convention, not written off
as broken.

### 3.2 Live-Excel hang safety — process isolation (PRD §6.2.3) — **not yet built**

Every live-Excel call (`spawn`, `xlw_open_workbook`, `xlw_save_workbook`, and every not-yet-built
one — `recalculate`, `run_macro`, `refresh_links`, `write_links`) is a potential hang point, not
just the already-fixed `close_owned()` case (§3.1's note). A Python thread can't forcibly
interrupt a blocked COM call cleanly; a process boundary can.

```python
# backends.py — new
@dataclass(frozen=True)
class TimeoutResult:
    """Outcome of a timeout-guarded live-Excel call.

    Args:
        completed: True if the call returned within the timeout (or no timeout was given).
        value: The call's return value, if completed.
        killed_pid: The Excel PID that was force-killed, if the call timed out.
    """
    completed: bool
    value: Any = None
    killed_pid: int | None = None


def run_with_timeout(
    fn: Callable[[], Any], timeout: float | None, owned_pid_file: Path
) -> TimeoutResult:
    """Run `fn` in an isolated worker process, enforcing `timeout` if given.

    Args:
        fn: The blocking live-Excel call to run (e.g. a closure wrapping `spawn()` then
            `app.calculate()`).
        timeout: Seconds to wait, or None to wait indefinitely (PRD §6.2.4's decision — no
            default timeout is ever imposed).
        owned_pid_file: Path the worker writes its spawned Excel PID to immediately after
            spawning — written *before* the risky blocking call, so the parent can still find
            and kill the right (and only the right) Excel process even if the worker itself
            becomes unresponsive.

    Returns:
        A `TimeoutResult`. `completed=False` means the worker (and its owned Excel PID, read
        back from `owned_pid_file`) were force-killed after `timeout` elapsed.
    """
```

Uses `multiprocessing.Process` (a genuine OS process, not a thread) for the worker; the parent
calls `process.join(timeout)` and, if the process is still alive afterward, reads
`owned_pid_file` to find the Excel PID to kill (via `OwnedInstanceRegistry`-style lookup, §3.1),
terminates the Excel process, then terminates the worker process itself. A timed-out call always
surfaces as `ActionExecutionError` (never a recoverable `ActionResult(status="error")`) —
workbook state after a forced kill can't be trusted (PRD §6.2.3/§6.8).

**Not yet built** — needs a live Excel instance on Windows to develop and verify the
kill-sequencing for real, not just reasoned about. Build order item 12 (below).

### 3.3 Configurable timeouts + signal summarization for `recalculate`/`run_macro` (PRD §6.2.4) — **not yet built**

```python
# backends.py — new, once recalculate/run_macro land
@dataclass(frozen=True)
class CalculationWaitSummary:
    """Compact audit-log-ready summary of a recalculate/run_macro wait (PRD §6.2.4).

    Args:
        state_counts: Histogram of every distinct signal value observed (e.g.
            {"xlPending": 3, "xlCalculating": 118, "xlDone": 1}) — empty if no signal exists
            at all (the `run_macro` case).
        last_state: The final observed value, or None if no signal was available.
        poll_count: Total number of liveness polls taken.
        elapsed_seconds: Total wall-clock time waited.
        outcome: "completed" | "timed_out" | "no_signal_available".
    """
    state_counts: dict[str, int]
    last_state: str | None
    poll_count: int
    elapsed_seconds: float
    outcome: Literal["completed", "timed_out", "no_signal_available"]
```

`recalculate(timeout: float | None = None)` and `run_macro(timeout: float | None = None)` both
gain an optional `timeout` param, defaulting to unbounded (PRD §6.2.4 — never a short default,
since legitimate real-world calculations against plugin-formula workbooks can take hours). Built
on top of §3.2's `run_with_timeout` for the actual hang-safety/kill mechanism; a **separate**
watchdog connection (not the same blocked call) is what would poll `Application.CalculationState`/
`Application.Ready` for `recalculate` specifically, producing the `CalculationWaitSummary` above
— per PRD §6.2.4's research finding, polling from the *same* call that triggered the calculation
is unreliable (can stay `xlPending` indefinitely). `run_macro` has no equivalent signal at all;
its summary is always `state_counts={}`, `last_state=None`,
`outcome="no_signal_available"` on completion, or `"timed_out"` if `timeout` elapsed.

The poll-interval/signal-capture logic must live in one small, isolated function (not scattered
across `recalculate`/`run_macro`'s own bodies) — expected to be tuned as real behavior is
observed via soak testing (PRD §6.2.4), not treated as a one-time design to leave alone.
**Not yet built** — needs empirical verification against real Excel first (build order item 12).

### 3.4 Refreshing a linked consumer workbook against a not-yet-committed source (PRD §6.3.2) — **not yet built, Option 2 decided**

Scenario (PRD §6.3.2): workbook A is modified via file-backend actions this run (scratch-only,
real file untouched per §5.3); a separate workbook B has external references to A — classic
cell-reference links, and/or a data connection/Power Query, and/or simply needs recalculating —
and must reflect A's *new* values before B is itself used. **Decided: redirect-then-restore**
(not an early checkpoint-commit of A) — preserves the "real files untouched until final
success" invariant (§5.3) exactly.

```python
# backends.py — new, needs write_links (com-tier, .api's ChangeLink/LinkSources) built first
def redirect_external_links(workbook: Any, old_target: str, new_target: str) -> None: ...
def restore_external_links(workbook: Any, redirected: dict[str, str]) -> None: ...
```

Sequence for a step that refreshes B against A: (1) `redirect_external_links(B, A.real_path,
A.scratch_path)` — rewrite B's references to point at A's scratch copy; (2) refresh B normally
(classic links via `refresh_links`, connections via Excel's own query refresh, or plain
`recalculate`); (3) `restore_external_links(B, ...)` — rewrite back to `A.real_path` *before* B
itself is committed, so B's real file never ends up pointing at a scratch path.

**Scope, per PRD §6.3.2**: classic cell-reference links are real, buildable work (xlwings' `.api`
`ChangeLink`/`LinkSources`). **Power Query/data-connection redirection is a narrower, harder
case** — a query's source lives in its M-code/connection string, not a simple external-link
table; v1 scope is likely limited to connections whose source is a plain file-path parameter,
with anything more complex documented as a named limitation, not silently attempted. Plain
`recalculate` needs no redirection at all — it's a separate, ordered operation from refreshing
links/connections (PRD §6.3.2's third bullet).

**Not yet built** — depends on `write_links`/`refresh_links` (build order item 10) existing
first; the redirect/restore wrapper is build order item 13 (below).

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
for file, `xlw_`-prefixed for the portable xlwings API, `com_`-prefixed for the raw-COM
exception) `backends.py` function. Backend choice still isn't the action's decision — it's fixed
once, by the capability tag, not decided per-call.

**Each action registers its capability via a decorator** (`@file_action`/`@xlw_action`/
`@com_action`, defined in `core.py` alongside `ACTION_CAPABILITIES`, a plain `name -> capability`
dict the decorators populate) rather than by stamping an attribute onto the function object —
keeps mypy --strict clean (no dynamic-attribute `type: ignore` noise) and keeps registration
trivially introspectable for `engine.py`'s `discover_actions()` (§5.1).

**The 5 built actions have a deliberately reduced param surface vs. the full PRD §7 catalog**,
each documented in its own docstring: `open` omits `update_links` (no effect without a live
Excel session — xlwings, a later phase) and a `mode` override (depends on read/write inference that
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
  lookup → aggregate → links → live-Excel), with a one-line comment banner per group, so the physical
  layout still mirrors the catalog even without file boundaries doing it.

**15 built so far** (corrected — earlier notes across items 4–7 said 14, an off-by-one
miscount not caught until `list_actions()` (§6.3) asserted the real count directly): `open`,
`save`, `close`, `copy`, `read_range`, `read_metadata` (properties/cells sub-cases),
`write_cell`, `write_range`, `write_row` (base + positional modes), `insert_range`
(whole-row/whole-column only), `set_column_width`, `find_headers_row`, `find_row`,
`find_column`, `find_columns`. All green, 100% branch coverage.

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

**`stop` — control-flow action — built (PRD §6.9).** **16 actions now** (was 15). Unlike every
other action, `stop` takes no `session` parameter and has no `workbook:` field at all — it's
pure control flow inside `runner.py`'s loop (§6.1), not a backend call, registered via a new
`@control_action` decorator (capability `"none"`, a fourth value alongside `"file"`/`"com"`/
`"depends_on_param"`). It joins `copy` in `_SCHEMA_EXEMPT_ACTIONS` (§5.4) since neither's YAML
shape is the generic "flat `workbook:` + params" the standard schema check handles.

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

**A real gap found while adding a regression test in `runner.py`'s build: `@file_action`/
`@com_action` were typed as `Callable[..., ActionResult] -> Callable[..., ActionResult]`,
which erases every decorated action's actual parameter types to "accepts anything" — silently
defeating mypy --strict at every call site, not just inside the action itself.** A call like
`read_metadata(session=s, target="textboxes")` — invalid, `target` is
`Literal["properties", "cells"]` — type-checked clean under the old decorator typing. Fixed
with `ParamSpec` (`Callable[P, ActionResult] -> Callable[P, ActionResult]`), which preserves
the exact signature through the decorator; mypy now genuinely catches a bad call to any action.
One caveat this doesn't and can't fix: the runner's actual dispatch calls every action via
`fn(session=session, **kwargs)` with `kwargs` built from a dynamically-typed dict — no amount
of decorator typing can check that path, which is exactly why the actions themselves still
need their own runtime guards for genuinely invalid input (e.g. `read_metadata`'s explicit
rejection of an unsupported `target`, found via this same investigation and fixed in
`actions.py`).

`capability="depends_on_param"` is a **named, single exception**, not a general mechanism —
only `read_metadata` uses it (PRD §7: file for `properties`/`cells`, xlw for `textboxes`), and
isn't built yet. `runner.py` (not built yet) will check for this literal case explicitly rather
than building a generic capability-resolution feature for one action.

`discover_actions(module)` scans a module with `inspect.getmembers`, keeping only functions
with an entry in `core.py`'s `ACTION_CAPABILITIES` dict (populated by the `@file_action`/
`@com_action` decorators, §4) — not every function in `actions.py`, just the tagged ones.
`param_schema` is derived from the function's signature, skipping `session` and marking any
parameter with no default as required.

### 5.2 Session management — **built** (except bidirectional backend switching, see below)

```python
@dataclass
class WorkbookSession:      # lives in core.py — see §4's correction
    name: str
    backend: Literal["file", "xlw"]
    handle: Any                      # openpyxl Workbook | xlwings Book
    path: str                        # added during implementation — see below
    mode: Literal["read_only", "read_write"]
    scratch_path: Path | None = None
    dirty: bool = False

# module-level, engine.py — built (PRD sec 6.2.2)
def _needed_backend(
    capability: Literal["file", "xlw", "com", "depends_on_param", "none"],
) -> Literal["file", "xlw"]: ...   # raises ActionExecutionError for the last two — see below

class SessionManager:
    def get_or_open(
        self, name: str, mode: Literal["read_only","read_write"] = "read_write",
        capability: Literal["file", "xlw", "com", "depends_on_param", "none"] = "file",
    ) -> WorkbookSession: ...     # built — raises on a capability/backend mismatch, see below
    def checkpoint(self) -> None: ...     # save every dirty staged session's scratch file
    def commit_all(self) -> None: ...     # save-all (via checkpoint's helper) + ScratchManager commit, §5.3
    def close_all(self) -> None: ...      # always runs — see runner.py's try/finally, §6.1
    # _switch_backend (the actual save/close/reopen dance, PRD sec 6.2.2) is NOT built yet —
    # needs the live-Excel phase's remaining pieces (§8 item 10) and a live Excel instance to
    # verify for real. No stub for it; it's simply absent until then.
```

**`checkpoint()` was added after the fact, found via a failing crash-safety integration test,
not designed up front.** openpyxl writes stay in memory until an explicit save — nothing
flushes them to the scratch file on disk mid-run on its own — so without this, a crash after
several successful steps left the scratch copy no more informative than the untouched original:
none of the in-progress work was actually on disk yet. `runner.py`'s step loop (§6.1) now calls
`checkpoint()` after every step, not only at the very end, so the recovery artifact (PRD
§6.3.1) genuinely contains everything that succeeded before a crash. Both `checkpoint()` and
`commit_all()` share one private helper that saves every staged session where `dirty` is true
and then clears the flag — `commit_all()` keeps its own save pass too even though per-step
checkpointing usually makes it a no-op by the time a run reaches there, as the final safety net
at the actual commit boundary, not removed just because it's often redundant.

**`path: str` was missing from the original sketch and had to be added** — `save` needs to know
*where* to save to, and there's no way to derive that without the session carrying its own
current path. It's `path`, not reused-as-`scratch_path`, because `scratch_path` specifically
means "was this session staged" — `path` is always concrete: the scratch path for a staged
(read-write) session, the real path for a read-only one.

**Mode is caller-specified, not statically inferred — this is deliberate for now, not a
shortcut.** PRD §6.3's "infer read-only vs. read-write from whether a workbook is ever a
target" is tier-2 validation's job (§5.4), which comes *after* session management in the build
order. Rather than block session management on validation existing, `get_or_open`'s `mode`
param is the seam that inference will feed into later — `mode="read_write"` today just means
"the caller is telling me this workbook will be written to," which is exactly the condition
PRD §6.3.1 stages against, whether a human or a validation pass decided it.

**Bidirectional backend switching (PRD §6.2.2) — capability threading is built; the actual
switch is not.** `get_or_open` gained a `capability` param (from the dispatching action's
`ActionSpec.capability`, §5.1) alongside `mode`, threaded through from `runner.py`'s `_dispatch`
(§6.1) — it already had the registry lookup to find `fn`, so passing `.capability` too was the
only change needed on the dispatch side, no new information threaded in from elsewhere. A new
module-level `_needed_backend(capability)` maps capability to the backend it needs (`file` →
`file`; `xlw`/`com` → `xlw`, since `com` reaches deeper via xlwings' `.api` on an xlw-backed
session rather than needing a distinct backend state). `get_or_open` now checks every returned
session (newly-opened or cached) against this: **if the session's current backend doesn't match
what the capability needs, it raises a clear `ActionExecutionError` rather than silently
returning the wrong backend or switching** — the actual switch (save-if-dirty → close → reopen
on the other side, PRD §6.2.2's `_switch_backend`) isn't built yet, so this is the honest
boundary of what's supported today, not a stub. Every action built so far is `file`-capability,
so this boundary is never hit in current real usage — proven by the full existing test/
integration suite passing unchanged with `capability` defaulting to `"file"` everywhere it
isn't explicitly passed.

**Still to build**: `_switch_backend` itself (save-then-close-then-reopen, mirroring the
sequence already used by `checkpoint()`/`close_all()`, just also triggered mid-run at the exact
point a switch is needed) and its `OwnedInstanceRegistry` (§3.1) integration — `SessionManager`
will hold one registry per run, spawning its single shared App lazily on the first `xlw`/`com`
request, not one App per workbook. `close_all()` (below) will need to additionally call
`self._xlw_registry.close_owned()` once this exists. Needs a live Excel instance to build and
verify for real — paused per the user's request (2026-08-20), see `docs/Progress_Tracker.md`.

**`create_if_missing`/`template` resolution lives in `SessionManager._create()`**, called from
both the read-write and read-only open paths when the target file doesn't exist yet. `template`
is a *logical name* (another entry in the `workbooks:` registry), resolved to that entry's
`file` path before delegating to `backends.create_workbook()`.

**A real bug, caught by a coverage gap, not by intuition**: the read-only-plus-`create_if_missing`
combination (unusual — why read something you just created blank? — but not forbidden) failed
because `_create()` never ensured its target's parent directory existed. The read-write path
got this for free from `ScratchManager.stage()`'s own `mkdir`; the read-only path didn't have
an equivalent. Found by chasing a coverage report down to an untested branch, not by reasoning
about it up front — exactly the kind of thing "90%+ branch coverage" is supposed to catch.

**Error type choice**: unknown workbook name, and missing file without `create_if_missing`,
both raise `ActionExecutionError` (not `ValidationError`) — both only surface once
`SessionManager` actually tries to touch the filesystem, which happens during a run, not during
static (tier-1) validation. Once tier-2 validation (§5.4) exists, the *unknown name* case should
be caught earlier as a `ValidationError` before any workbook is touched at all — this exception
is the fallback for whatever tier-2 doesn't catch, not the primary catch of that particular
mistake.

**`close_all()` aggregates failures via `ExceptionGroup` rather than stopping at the first
one** — every session gets a close attempt regardless of whether an earlier one failed (PRD
§6.3's crash-safety requirement), but every failure is still surfaced afterward, not silently
swallowed. This is a deliberate exception to "avoid defensive fallback layers" (PRD §1): the
distinction is inventing fallback *behavior* for bad input (the anti-pattern) vs. guaranteeing
cleanup still runs and still reports what went wrong (the actual requirement).

### 5.3 Scratch-copy execution model — **built** (2026-08-21, supersedes the original tempfile-based design)

Implements PRD §6.3.1, revised by §6.3.3/§6.3.4's decisions:

```python
class ScratchManager:
    def __init__(self, working_dir: Path) -> None: ...
       # working_dir: <base>/excel_runner_runs/<yaml_stem>/ — see §6.1's note on where <base>
       # comes from. scratch/ workbook copies live at working_dir/scratch/.
    def stage(self, name: str, real_path: Path, writes: bool = True) -> Path: ...
       # copies real_path into working_dir/scratch/ if it exists; if not (create_if_missing
       # case), just reserves the scratch path for the caller to create the workbook at
       # directly. Called for BOTH read-write AND read-only sessions now — see correction below.
       # writes=False (read-only) means commit_all() skips this workbook entirely.
    def commit(self, name: str) -> None: ...     # one workbook's rename-based commit, see below
    def commit_all(self) -> None: ...
       # raises ActionExecutionError on any workbook's commit failure — see rollback design below
```

**Correction (PRD §6.2.3): read-only sessions are now staged too, not opened directly against
the real path.** Originally "open read-only in place… faster, safer" (§6.3's original wording) —
revised after research confirmed openpyxl's `read_only=True` keeps a genuine OS file handle open
until `.close()`/process exit, and while a genuine crash self-heals (OS releases handles on
process teardown), a *hang* does not — the file stays locked for as long as the hang persists.
Staging read-only opens too doesn't prevent the hang, but changes its blast radius from "a
shared/production file is locked, someone has to notice and intervene" to "an orphaned copy in
our own working_dir, harmless." `SessionManager._open_read_only` now calls
`scratch.stage(name, real_path, writes=False)` exactly like `_open_read_write` does (which uses
the `writes=True` default); the only difference is `writes=False` means `commit_all()` skips it
entirely — nothing about a read-only session's content ever changes, so there's nothing to
commit back.

**Correction (PRD §6.3.4): `working_dir` replaces the old `tempfile.mkdtemp()`-based scratch
dir entirely** — a fixed, predictable location (`<base>/excel_runner_runs/<yaml_stem>/`) instead
of a random per-run temp path, so external tooling can construct the path itself from just the
yaml's filename, without reading any output field. See §6.1 for exactly how `<base>` is resolved
(cwd default, or the `--working-dir` CLI flag — see §6.4's correction: the originally-sketched
`working_dir:` YAML field was **not built**, kept out of scope since there was no clear use case
for it beyond the CLI flag, which already covers the real driving need). Re-running the same
yaml overwrites the previous run's `working_dir` contents automatically — `stage()`'s existing
`shutil.copy2` already overwrites in place; `AuditLogger` (§6.2) now opens `audit.jsonl` in
truncate mode at the start of each run, not append, so a new run's audit trail never mixes with
a previous run's leftover records.

**`cleanup()` is removed entirely** — nothing in `working_dir` is ever deleted automatically now,
success or failure (PRD §6.3.4). Safe because a re-run of the same yaml just overwrites its own
fixed folder rather than accumulating; there is no unbounded growth to guard against. The only
things ever deleted are the transient `.bak`/`.tmp` files the commit mechanism below creates and
removes itself within a single run.

**Rename-based commit, with per-file rollback on a later failure (PRD §6.3.3) — replaces the
original "temp-path-then-Path.replace" description**:

Per-workbook commit sequence (`ScratchManager.commit(name)`, called directly for one workbook
or looped over by `commit_all()` for the whole batch):
1. `shutil.copy2(scratch_path, tmp_path)` — prepare the new content off to the side first,
   *before touching `real_path` at all*. The only step that can fail for disk-space/permission
   reasons unrelated to `real_path` being locked; if it fails, `real_path` is completely
   untouched, no rollback needed.
2. If `real_path` exists: `real_path.rename(bak_path)` — an instant, zero-copy move. This *is*
   "keeping the original" (PRD's phrasing) — it was already sitting there untouched pre-commit,
   so preserving it costs nothing extra.
3. `tmp_path.rename(real_path)` — installs the new content.

`commit_all()` attempts every write-intent staged workbook's commit in turn (no separate
upfront precheck pass — simplest option discussed, PRD §6.3.3). If workbook *N*'s `commit()`
raises `OSError`, `commit_all()` rolls back every workbook that already committed successfully
in this call, in reverse order, via `_rollback()`: `real_path.unlink(missing_ok=True)` then (if
a `.bak` exists — a brand-new `create_if_missing` workbook that had no prior real file won't
have one) `bak_path.rename(real_path)`. **Records, per file, whether that rollback itself
succeeded** — if it fails too, that workbook's name is named explicitly (`MANUAL INTERVENTION
NEEDED for: ...`) in the raised `ActionExecutionError`'s message, its `.bak` is deliberately
*not* deleted (so the original content is still recoverable from disk, not just described in an
error message). On full success (every workbook committed), every `.bak` created during the
call is deleted.

**Simplification found during implementation**: the originally-sketched `CommitFailure`
dataclass (a structured payload attached to the raised error) wasn't built — `ErrorDetail`'s
existing `message`/`technical_reason` fields already carry everything needed (which workbook
failed, which others were rolled back, which need a human) as clear, human-readable text; adding
a whole new public type for this would be extra surface for no real gain, given `ErrorDetail`'s
`message` is already meant to be the actionable, plain-English summary (§6.8).

**A partial-but-fully-rolled-back commit failure still makes `RunResult.status == "error"`**
(PRD §6.3.3) — the run didn't fully complete what it promised, regardless of how cleanly the
failure was contained.

### 5.4 Validation — two tiers — **built**

Matching PRD §9/§9.1, kept as clearly separated functions within the file since they run at
different times and check different things.

**Correction found while building tier 1: range/named-range syntax checking against a
workbook's real defined names is not implementable here** — it needs to open the workbook,
which contradicts this tier's own "no workbook access" premise. Not built; PRD §9.1/§12
corrected to record this as an open item rather than a solved one. What tier 1 actually checks,
via `validate_static(workflow, registry)`, five small functions run in a fixed order (each one
individually testable, and assuming every earlier one in the list already passed):

1. **Action exists** in the registry — `difflib.get_close_matches` suggests a fix for typos.
2. **No unrecognized params** — every key in `step.params` must be either a real property of
   that action's `param_schema`, or the one universally-implicit field, `workbook` (present on
   almost every action's raw YAML but absent from its Python signature, per §4's correction —
   hardcoded as a named exception here, not a general mechanism).
3. **No missing required params** — same allowance for the implicit `workbook` field.
4. **Param types match** — `_matches_type()` handles plain types, `Literal[...]` (membership
   check), `X | Y` unions (any branch matches), and generic aliases like `list[...]` (checks
   only the origin — "is this a list at all", not element-by-element). Produces PRD §9.1-style
   messages, including the "wrap it in [ ]" suggestion specifically when a list was expected
   and a string was given.
5. **Step references resolve** — every `steps.<id>` found anywhere in a step's `params` or
   `if_expr` (recursive scan, regex `steps\.([A-Za-z_]\w*)`) must name a step id that both
   exists and appears *earlier* in the list; a fuzzy-matched suggestion is offered when it
   doesn't exist at all.

**`copy` is exempt from checks 2–4** (`_SCHEMA_EXEMPT_ACTIONS`) — its raw YAML shape (`source:`/
`target:` dicts) doesn't match its Python signature (§4's two-session correction), so there's no
schema to check it against yet; that's the runner's translation-layer job (build order item 7),
not validation's, until that layer exists.

**Tier 2 (dry-run / step-graph), via `plan(workflow, registry) -> ExecutionPlan`** — still no
workbook access, reasons over the whole step list together:
- Checks every workbook name referenced anywhere in any step's params (a generic recursive
  walk for any key literally named `workbook`, handling both a flat field and `copy`'s nested
  dicts with the same code) actually appears in `workbooks:`.
- Infers each workbook's mode: `read_write` iff some step's action has `writes=True` in the
  registry (looked up as `registry[step.action].writes`), else `read_only`. **Correction**:
  originally a hardcoded `_WRITE_ACTIONS` set living here in `engine.py`, disconnected from
  each action's own definition — a second place to remember to update whenever a new write
  action was added, and easy to forget (found while adding `create_sheet`/`rename_sheet`/
  `delete_sheet`). Now `writes` is declared on the *same* `@file_action`/`@xlw_action`/
  `@com_action` decorator that registers the action's capability
  (`@file_action(writes=True)`), so `ActionSpec.writes` is derived straight from the registry
  — one source of truth, right next to the function, not a list to keep in sync by hand.
  `copy` is a documented exception: since its
  source vs. target can't be told apart statically yet (same reason as above), *every* workbook
  it references gets marked `read_write` — the safe over-provisioning PRD §6.3 itself prescribes
  ("if a target reference can't be resolved statically, default to read-write") rather than
  risking one being silently left read-only.

This *is* the "plain-English execution plan" PRD §9 promises an agent/user can sanity-check
before a real run — `SessionManager.get_or_open`'s `mode` param (§5.2) is the seam this plugs
into once `runner.py` exists to wire them together.

## 6. Runner layer — `runner.py`

The composition root, the audit trail, and the one public contract — grouped because together
they're "the layer that actually executes a run and is safe for other code to depend on."

### 6.1 Orchestration — **built** (2026-08-21, working_dir resolution)

```python
def run_workflow(
    path: str | Path, env_overrides: dict | None = None, working_dir: str | Path | None = None,
) -> RunResult:
```

One linear sequence (per project convention: a composition root doesn't need splitting just to
hit a line count — see AGENTS.md), wrapped in `try`/`finally` for the crash-safety guarantee
(PRD §6.3). Actual steps, corrected against what got built:

1. `core.load(path, env_overrides)` → `Workflow`.
2. Tier-1 validation (§5.4) → raises before anything is opened.
3. Tier-2 validation (§5.4) → `ExecutionPlan`.
4. **Resolve `working_dir` (PRD §6.3.4 — replaces `tempfile.mkdtemp()`)**:
   `base = Path(working_dir) if working_dir is not None else Path.cwd()`; `run_dir = base /
   "excel_runner_runs" / Path(path).stem`. `working_dir` is fed by the CLI's `--working-dir`
   flag (§6.4) — **the originally-sketched `working_dir:` YAML field was not built**, kept out
   of scope; no clear use case for it beyond the CLI flag actually showed up, so it wasn't
   added speculatively. `scratch = ScratchManager(run_dir)`; `session_manager =
   SessionManager(workflow.workbooks, scratch)`. **No separate upfront staging loop** —
   staging happens lazily, inside `SessionManager.get_or_open`, the first time each workbook is
   actually referenced (now true for read-only sessions too, §5.3's correction). The runner's
   only job here is to pass the right `mode` (from `plan.modes[name]`) into each `get_or_open`
   call.
5. For each `Step` in order: if `if_expr` is set and evaluates false, record a `"skipped"`
   `StepResult` and move on (logged at `INFO`, §6.2.1). Otherwise: `logger.info(...)` that the
   step is starting, then dispatch — every action except `copy` resolves its `workbook` field
   into one session and calls `registry[step.action].fn(session=session, **remaining_kwargs)`;
   `copy` is dispatched separately (`_dispatch_copy`), resolving *both* `source.workbook` and
   `target.workbook` into two sessions, exactly the special-case wiring §4/§5.1 flagged as
   still owed. `_dispatch` itself logs the resolved params at `DEBUG` right after
   `core.resolve_value` runs (§6.2.1) — not duplicated in the main loop, to avoid resolving the
   same params twice per step just for a debug line. Every step gets an audit record (§6.2)
   regardless of outcome, an `INFO`/`ERROR` console record on completion (§6.2.1), and — added
   after a failing crash-safety test, not part of the original design — every step also gets a
   `session_manager.checkpoint()` call (§5.2), persisting whatever it just wrote to the scratch
   file before moving on.
6. **An action returning `ActionResult(status="error")` does not stop the loop** — this was an
   open design question, resolved by what `if:` conditions are actually for: PRD's own
   `if: "{{ steps.refresh.status == 'success' }}"` example only makes sense if a failed step
   doesn't abort the run before later steps get a chance to check its status. So the loop
   continues, but `RunResult.status` is `"error"` if *any* step failed, and that governs
   whether anything gets committed — not whether the loop finished. A *raised* exception
   (`ActionExecutionError`, `ValidationError`) is the only thing that actually aborts the loop,
   consistent with the error-handling policy from §4: a raised exception means a genuine
   mistake, not a normal "didn't work" outcome.
7. If no step failed: `session_manager.commit_all()` (saves every staged session, then
   `scratch.commit_all()` moves each scratch file to its real path — see §5.3's rename-based
   commit/rollback design). **No `scratch.cleanup()` call anymore** (§5.3/§6.3.4 — nothing in
   `working_dir` is deleted automatically, success or failure).
8. `finally: session_manager.close_all()` — unconditionally, whether the loop finished, a step
   failed, or an exception propagated. `working_dir`'s contents (scratch copies + `audit.jsonl`)
   are simply left in place either way now — not conditionally cleaned up.


**`stop` — built (PRD §6.9)**: a step whose `action` is `stop` and whose `if:` is true (or
absent) ends the loop immediately after being recorded — the one *action* that does stop the
loop, deliberately distinct from #6 above (a normal `status: "error"` result, which doesn't).
Every step after it gets `StepResult(status="stopped")` instead of being dispatched at all, so
`RunResult.step_results` keeps its "one entry per workflow step" contract, and the audit log can
distinguish "this step's own `if:` said don't run me" from "the run ended before we got here."
`stop` doesn't set `any_failed` itself — whether the run commits is still governed purely by
whether any *earlier* step returned `status: "error"` (#6 above), unchanged. Implemented as a
small addition to the per-step loop (`enumerate` for the index, a nested loop over the remaining
steps on trigger) — not a new abstraction, per the "composition root doesn't need splitting"
convention (§6.1's intro note).

`OwnedInstanceRegistry.close_owned()` isn't wired in here — `SessionManager` doesn't promote
sessions to xlwings yet (§3.1, build order item 10), so there's nothing to close on that front.

**A real bug found while writing the first integration test**: the audit log was originally
written inside the same directory `ScratchManager.cleanup()` deletes — so a *successful* run
was deleting its own audit trail. Fixed by splitting the run directory into `run_dir/scratch/`
(what `cleanup()` touches) and `run_dir/audit.jsonl` (what it doesn't).

### 6.2 Audit logging — **built**

```python
class AuditLogger:
    def record_step(self, step: Step, result: StepResult, started_at, ended_at) -> None: ...
```

Takes a `StepResult`, not an `ActionResult` as first sketched — `StepResult` is the superset
that also covers `"skipped"`, which never produces an `ActionResult` at all (the action never
runs). One JSON object per line (JSONL), written to `working_dir/audit.jsonl` (§6.1's bug fix
above for why not inside `scratch/`; `working_dir` itself per §5.3/§6.3.4's redesign). Opened in
truncate mode at the start of each run, not append — a re-run against the same `working_dir`
must never mix a new run's records with a previous run's leftovers (§5.3's correction). Logs the
step's *raw* params, not resolved ones — resolved values aren't available uniformly for a
skipped step, and raw params are simpler and always available; a minor, deliberate deviation
from the original "resolved parameters" phrasing. Not a `logging`-module handler — deliberately
a separate, structured artifact (PRD §6.7 explains why).

### 6.2.1 Console/application logging (PRD §6.7.1) — **built** (2026-08-21)

A second, distinct output from the audit log above: real-time narration via stdlib `logging`,
for a human (or the user's own workflow tool) watching a run as it happens. Only `runner.py`
has a module-level `logger = logging.getLogger(__name__)` so far — `_dispatch`'s `DEBUG` line
is the only per-step detail logged today (no handler/formatter configuration in library code —
standard Python practice, never hijack whatever logging setup the importing application
already has). Per PRD §6.7.1:
- `INFO`: every step's start (`'Step "%s" (%s): starting'`), a skip
  (`'Step "%s" (%s): skipped (if: was false)'`), or a non-error completion
  (`'Step "%s" (%s): %s'`, the action's status).
- `DEBUG`: resolved params, logged once in `_dispatch` right after `core.resolve_value` runs —
  not duplicated in the main loop, to avoid resolving the same params twice per step.
- `ERROR`: a failed step (`'Step "%s" (%s): failed — %s'`, the error's plain-English message) —
  self-sufficient for a human reading just the console, not merely pointing at the audit log.

The CLI (§6.4) is the only place that sets a severity threshold (`--logging-level`) — it does
not attach handlers either; stream/format configuration is explicitly out of scope for
excel_runner entirely (PRD §6.7.1's correction after initial over-design).

### 6.3 Public API surface — **built**

The only symbols other Python code (or a future CLI/MCP wrapper, PRD §5) should import,
re-exported from the package's `__init__.py`:

```python
from excel_runner import run_workflow, RunResult, StepResult
from excel_runner import Workflow, Step, WorkbookRef   # for programmatic construction
from excel_runner import list_actions, ActionSpec        # list_actions() -> tuple[ActionSpec, ...]
```

Everything else in the package tree is an implementation detail and may change without notice;
this surface is the versioned contract (PRD §3/§9/§12). **`ActionSpec` itself had to join the
re-exports** — not in the original sketch, but `list_actions()`'s return value is a tuple of
`ActionSpec` instances, so calling code needs the type itself to work with the result
meaningfully (e.g. `isinstance` checks), not just the function.

**A real gap found while building this: `ActionSpec` had no `description` field at all.** The
original plan's own words — "`list_actions()` exposing `ActionSpec` (name/**docstring**/
param_schema)" — assumed a field that was never actually added back in §5.1. Without it,
`list_actions()` couldn't fulfill its one stated purpose (PRD §6.1's "close to free" schema
generation for a future agent-tool wrapper) — a tool definition needs a description, not just a
name and parameter shape. Fixed: `ActionSpec.description` is now populated from each action's
docstring, first line only (`inspect.getdoc(fn)`, then split on the first newline) — every
action already had a clean one-line summary to start with, so no action docstrings needed
rewriting to support this.

`list_actions()` itself is `tuple(discover_actions(actions_module).values())` — deliberately
just `discover_actions` (§5.1) wired to the real `actions` module, not a second source of truth
that could drift from it.

### 6.4 CLI — `cli.py` — **built** (2026-08-21)

```python
def main(argv: list[str] | None = None) -> int:
```

Thin wrapper over `run_workflow()` — argument parsing and JSON result formatting only, no logic
of its own (§1's correction). Args: `workflow` (positional path), `--env KEY=VALUE`
(repeatable), and (added 2026-08-21, PRD §6.3.4/§6.7.1):

```python
parser.add_argument("--working-dir", default=None, help="Base directory for this run's "
    "working_dir (excel_runner_runs/<yaml_stem>/ is always appended). Defaults to cwd.")
parser.add_argument("--logging-level", help="DEBUG,INFO,WARNING,ERROR", default="INFO")
```

`--working-dir`'s value is passed straight through as `run_workflow(..., working_dir=...)`
(§6.1). `--logging-level` calls `logging.getLogger("excel_runner").setLevel(...)` before
running the workflow — setting the level on the package's parent logger, not each module's own
`__name__`-based logger, so it propagates down to every child logger (`excel_runner.runner`,
etc.) that doesn't set its own explicit level. No handler or formatter configuration (§6.2.1's
decided scope boundary).

## 7. Testing approach

- **`tests/unit/`** stays fine-grained even though `excel_runner/` doesn't — one test file per
  action (`tests/unit/actions/test_read_range.py`, ...), per class (`test_session_manager.py`,
  `test_scratch.py`, ...), per concern within `core.py` (`test_templating.py`,
  `test_schema.py`, ...). **Correction: nothing is ever mocked, not even in unit tests** — the
  original plan said action unit tests would mock `backends.py`; in practice every action test
  uses a real `WorkbookSession` wrapping a real openpyxl workbook (via a shared `conftest.py`
  fixture), same as `backends.py`'s own tests. openpyxl needs no live Excel, so this is a real
  dependency exercised cheaply, not a mock — turned out to be just as easy as mocking would
  have been, with no risk of the mock drifting from real behavior.
- **`tests/integration/`** — zero mocks, per project convention, now built (build order item
  7): `tests/integration/test_run_workflow.py` runs real `workflow.yaml` text through
  `run_workflow()` against real openpyxl workbooks, exactly the shape the user and Claude
  agreed on — a runnable YAML fixture *is* the integration test, more realistic than testing
  components in isolation. **Correction: fixture workbooks are generated in code
  (openpyxl, written to `tmp_path`), not committed as static binary files in `tests/data/`** as
  first sketched — more reviewable (visible in a diff, no binary blobs), easier to vary per
  test, and every other test in this codebase already does it this way. `tests/data/` stays
  empty/unused unless a real need for a static fixture shows up. xlwings-backed actions/tests
  will be marked `@pytest.mark.skipif` on platform/Excel-availability (PRD §4's "test what's
  testable on macOS now, finish on Windows later") — the pattern's already in place
  (`tests/unit/conftest.py`'s `requires_excel`, used by `OwnedInstanceRegistry`'s tests,
  build order item 10) even though no *action* uses it yet.
- **Crash-safety test — built, as a follow-up to item 7**: `TestCrashSafety` in
  `tests/integration/test_run_workflow.py` deliberately triggers a raised exception mid-run
  (`write_row`'s positional mode without `start_column` — a real, already-covered way to
  produce one) and asserts: the real file is completely untouched; the scratch copy survives
  and actually contains the prior step's write (not just whatever existed at staging time —
  see §5.2/§6.1's `checkpoint()` correction, found by this very test failing on the first
  attempt); the audit log survives too; and — the strongest cross-platform evidence that
  sessions were genuinely closed — a second, valid run against the same real file afterward
  just works. **"No orphaned Excel process" is explicitly not tested** — no *action* spawns
  Excel yet (`OwnedInstanceRegistry` itself does, and is tested for that directly, §3.1 — but
  nothing wires it into a real run); that part of PRD §6.3's requirement gets a real
  `run_workflow()`-level test once build order item 10's actions exist. Directly detecting
  an OS-level file lock was deliberately not attempted either — meaningful mainly on Windows
  (PRD §4), not reliably testable on macOS, so the "does a later run succeed" check stands in
  for it as the behavior that actually matters.

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
5. `engine.py` §5.2/§5.3 — `SessionManager` (lazy-open, `create_if_missing`/`template`,
   close-all) and `ScratchManager` (scratch-copy staging/atomic commit). **Done**, except
   bidirectional backend switching (PRD §6.2.2 — needs the live-Excel phase, item 10) and mode
   inference (needs §5.4, item 6) — mode is caller-specified for now, the seam validation will
   feed into later. See §5.2's notes for the real bug this surfaced (missing parent-directory
   creation on one code path) and the `ExceptionGroup`-based crash-safety design in `close_all()`.
6. `engine.py` §5.4 — both validation tiers. **Done**, except checking a range against a
   workbook's real defined names (PRD §9.1's fourth example) — not implementable in either
   tier as designed (needs workbook access, both tiers are explicitly workbook-access-free),
   carried to PRD §12 as an open item.
7. `runner.py` §6.1/§6.2 — first end-to-end real run of a multi-step file-backend workflow.
   **Done.** Resolved the deferred `copy` two-session wiring and `workbook`-field-stripping
   translation flagged since item 3/4. Surfaced two real bugs (§6.1/§2.2's notes: the audit
   log being deleted by its own success-path cleanup; a dict output key named `"values"`
   shadowed by Python's real `dict.values()` method under Jinja2's default attribute
   resolution) and one real typing gap (§5.1's note: the capability decorators erasing every
   action's parameter types, switched to `ParamSpec`). First genuine
   `tests/integration/test_run_workflow.py` written — see §7's corrections to the original
   testing-approach sketch. The dedicated mid-run-crash integration test (§7) is also **done**
   — it found and fixed a real gap in the scratch-copy model itself (`SessionManager.checkpoint()`,
   §5.2), not just a missing test.
8. `runner.py` §6.3 — the public surface. **Done** — `excel_runner/__init__.py` re-exports,
   `list_actions()`, and `ActionSpec.description` (a real gap the original sketch's own words
   implied but never actually added — see §6.3's note). This is also where the "14 actions"
   count repeated across items 4–7's notes turned out to be an off-by-one for "15" —
   `list_actions()` asserting the real count directly is what caught it.
9. `runner.py` §6.1 — the `stop` control-flow action (PRD §6.9): schema-exempt registration
   (`stop` joins `copy` in `_SCHEMA_EXEMPT_ACTIONS`, §5.4), the new `"stopped"` `StepResult`
   status, and the runner loop's early-exit handling. Pure logic, no I/O — fits before the
   platform-dependent phases below, same rationale as the rest of this build order. **Done.**
10. **xlwings / live-Excel phase**: `backends.py`'s `OwnedInstanceRegistry` (§3.1) — **done**,
    tested for real on macOS (Excel is installed here). `xlw_open_workbook`, `xlw_close_workbook`,
    `xlw_save_workbook` — **done**, tested for real (open/close) or gated behind
    `requires_working_xlwings_save` (save — `tests/unit/conftest.py`, since `save()` is confirmed
    broken via xlwings on this Mac's Excel build, §3.1's note; real write-path verification needs
    the Windows environment, PRD §4/§12). `backends.py` sits at 99% branch coverage as a result —
    the one uncovered line is `xlw_save_workbook`'s `book.save()` call, only exercised on
    Windows; not forced to 100% with a `# pragma: no cover` since the line genuinely is
    reachable, just not on this test runner — expected to close once tested on Windows, not a
    permanent gap. Remaining live-Excel actions in `actions.py` (`recalculate`, `run_macro`,
    `refresh_links`, `write_links`, `read_metadata`'s textbox sub-case) — **not yet built**.
11. **Deferred/flagged, per PRD**: `update_summary_table`'s real parameters, the `aggregate`
    discussion, `export_pdf`, the AI-authoring inspection actions (PRD §9: `list_sheets`,
    `describe_sheet`).
12. **Crash/lock-safety hardening (PRD §6.2.3/§6.3.3/§6.3.4/§6.7.1)** — **done** (2026-08-21):
    - `working_dir` relocation (§5.3/§6.1): replaces `tempfile.mkdtemp()` with the fixed
      `<base>/excel_runner_runs/<yaml_stem>/` path; `AuditLogger` truncates instead of
      appending; `ScratchManager.cleanup()` removed entirely.
    - Read-only sessions now staged too (§5.3's correction to §6.3's original wording), via a
      new `stage(..., writes: bool)` flag so `commit_all()` knows to skip them.
    - Rename-based commit with per-file rollback on a later workbook's commit failure (§5.3) —
      the `.bak`/`.tmp` rename sequence; the originally-sketched `CommitFailure` dataclass
      wasn't built (§5.3's note — `ErrorDetail`'s existing fields already carry it clearly).
    - Console/application logging via stdlib `logging` (§6.2.1) plus the CLI's
      `--working-dir`/`--logging-level` flags (§6.4). The `working_dir:` YAML field sketched in
      the PRD wasn't built — no clear use case beyond the CLI flag showed up.
    - Every integration test that calls `run_workflow()` now passes `working_dir=str(tmp_path)`
      explicitly — found necessary because the *default* (cwd) is the real repo directory
      during a pytest run, which would otherwise litter the actual project with
      `excel_runner_runs/` test artifacts on every test run. `.gitignore` also covers it as a
      safety net for real, manual CLI usage inside the repo.
    - Pure logic + filesystem operations, no live Excel needed — fits before item 13/14 below,
      same "platform-independent work first" rationale as the rest of this build order.
13. **Live-Excel hang safety + configurable timeouts (PRD §6.2.3/§6.2.4)** — **not yet built**,
    needs a live Windows Excel instance to develop and verify for real (§3.2/§3.3):
    `run_with_timeout`'s process-isolation mechanism, then `recalculate`/`run_macro`'s
    `timeout` param and `CalculationWaitSummary` audit-log summarization built on top of it.
    Soak-testing real client workbooks (desktop + 3rd party workflow system) to establish empirical reliability is a
    validation activity for this item, not a separate build step.
14. **Linked-consumer-workbook refresh (PRD §6.3.2)** — **not yet built**, depends on item 10's
    `write_links`/`refresh_links` existing first (§3.4): `redirect_external_links`/
    `restore_external_links` for classic cell-reference links; Power Query/data-connection
    support scoped narrowly (plain file-path-parameter sources only) or documented as a named
    limitation.

`docs/Progress_Tracker.md` tracks each item above against the project's standard Component /
Unit Tests / Code / Integration Tests / Results columns, at function/class granularity — not
collapsed to match the 5-file source layout.
