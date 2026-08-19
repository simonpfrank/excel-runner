# excel_runner — Product Requirements Document

Status: **Draft — core action catalog and syntax mostly decided.** A couple of items
(`aggregate`, `update_summary_table`'s exact parameters) are explicitly flagged for later
discussion rather than settled now.

## 1. Problem statement

Excel automation work is currently built as one-off scripts per project, with no reusable
abstraction — every project reinvents file access (openpyxl/pandas) and live-Excel automation
(xlwings) from scratch. This is **not** an RPA tool (no click-simulation/UI automation) — it's a
declarative, code-based automation layer for Excel workbook operations, closer in spirit to
Ansible/GitHub Actions than to UiPath/Power Automate Desktop.

Design principles established up front, driving every decision below:

1. **Only pay for a live Excel COM session where the work genuinely requires it** —
   recalculation, VBA macros, live external-link refresh. Everything else (reads, writes,
   copies, structural edits) runs against the file directly via openpyxl: fast, headless, no
   live Excel process needed.
2. **Every field in the config means one thing, always, and is a real structured value** — a
   YAML list, a mapping, a typed scalar — never a semicolon/colon-packed string requiring a
   hand-rolled parser.
3. **Fail fast with a specific, actionable validation error** before a real workbook is
   touched, rather than silently falling back to guessed/defaulted behavior at runtime.

## 2. Goals

- A declarative, YAML-based workflow runner for Excel automation (open/read/write/save/macro
  operations across one or more workbooks), authored in a YAML dialect hung on a widely-known
  convention (§10 — Jinja2).
- An action catalog that is easy for a human (or an LLM) to author correctly on the first try —
  structured, self-describing fields; no encoded mini-languages; fail-fast validation instead
  of silent fallback behavior.
- An architecture that transparently uses a fast file-only backend (openpyxl) for actions that
  don't need it, and only pays the cost of a live Excel COM session (via xlwings) for actions
  that genuinely require Excel's calc engine or VBA — the user never chooses this explicitly.
- **A clean, directly-importable Python library**, usable from a plain script in another
  project, not only via a YAML file on disk — see §3 and §9.
- **An AI-agent-authoring experience that is as important as the engine working correctly** —
  a user describes what they want, an agent inspects the real workbook(s), and writes/expands
  a complete, runnable workflow from that description. See §9.
- Operational safety: a run that crashes partway through must never leave an Excel process
  orphaned or a workbook file locked (§6.3), and must never touch an Excel instance it didn't
  create itself — other workflows or a real user's own Excel session may be running
  concurrently on the same machine (§6.2.1).

## 3. Non-goals (for now)

- Not an RPA / UI-automation tool.
- The single-action CLI and any agent-facing *interface plumbing* (specific CLI syntax, MCP
  server wiring) are out of scope for v1 as concrete deliverables. But two things that enable
  them are **not** deferred — they're v1 architectural requirements:
  1. The engine must be a clean, directly-importable Python library from the start (usable from
     a plain script in another project — `from excel_runner import ...` — not only via a YAML
     file on disk). This is also the substrate an MCP server or CLI wrapper would sit on top of
     later, so building it in from day one avoids a rewrite (see the API-surface item in §12).
  2. See §9 — the AI-agent-authoring experience is a first-class v1-adjacent goal, not a
     "later" item, even though the specific wrapper mechanism is undecided.
- Not every action in the catalog needs to exist in v1 — see §8 for what's in v1 vs. later.

## 4. Target platform

Primary real-world target is a **Windows environment with a live Excel install** (this is
where the tool will actually be used day to day). Development happens on macOS.

Decision: use **xlwings as the COM abstraction on both platforms** (it handles the Windows/COM
vs. Mac/AppleScript difference internally) rather than raw `win32com`. Consequence: on macOS,
some COM-backed actions may behave differently or not be fully supported — this is an accepted
limitation of the dev environment, not a v1 blocker, since Windows is the real target. File-only
(openpyxl) actions are unaffected by platform. Testing strategy: test what's testable on macOS
now, finish COM-path integration testing on Windows later — accepted (see §12).

## 5. Interfaces

- **v1**: a YAML workflow runner — parse, validate, and execute a `workflow.yaml` file made of
  `workbooks:`, `env:`, and `steps:`, using Jinja2 templating (§10.1) for cross-step references
  and conditions.
- **v1 (architecture, not a shipped wrapper)**: the engine must be directly importable and
  usable from a plain Python script — not just runnable via a YAML file — see §3 and §9.
- **Deferred as a concrete deliverable**: the specific agent-facing wrapper (CLI with JSON
  output, or an MCP server — undecided). Because the engine is a clean importable library
  (above), either wrapper is a thin layer once the engine exists, not a redesign.

## 6. Architecture decisions

### 6.1 Actions as separate, auto-discovered files

Modeled on the tool-registry pattern already proven in `../agent-harness`
(`agent_harness/tools.py`): plain, typed Python functions with Google-style docstrings, one
per file, auto-discovered from a directory (mirroring `discover_tools()` /
`generate_schema()`). No pydantic, no class hierarchy for actions themselves.

Adaptation vs. agent-harness's stateless tools: an excel_runner action is not a one-shot
string-in/string-out call. Each action function takes `(session, **typed_params)` where
`session` is the current `WorkbookSession`, and carries a **capability tag** (`file` or `com`)
so the runner knows whether it can stay on the openpyxl backend or must promote the workbook
to a live xlwings session. **Backend selection is never exposed to the author** — the same
action name (`open`, `copy`, ...) is used regardless of which backend actually executes it.
One exception, noted where it applies (§7): `read_metadata`'s capability depends on its own
`target:` parameter, not a fixed per-action tag.

This also means: once the registry exists, generating an Anthropic-style tool schema
(name/description/input_schema) from each action is close to free — directly sets up the
deferred agent-tool wrapper without having to build that machinery twice.

### 6.2 Session model

A flat `SessionManager`: dict of logical workbook name (from the `workbooks:` registry) →
`WorkbookSession` (current backend handle, which backend is active, dirty flag). Steps address
cells as `(workbook_name, sheet_name, range)`. No additional "context" abstraction on top of
this.

A workbook that doesn't exist yet is created automatically the first time it's referenced, if
its registry entry sets `create_if_missing: true` (optionally from a `template:` workbook) —
this is a registry-entry concern, not a separate step action; see §11 item 10.

### 6.2.1 xlwings instance ownership

xlwings — and the underlying Excel COM API — will happily attach to whatever Excel instance is
already running on the machine, rather than always spawning an isolated instance of its own.
On a machine where more than one automation run (or a real user's own Excel session) can be
active at the same time, this is dangerous: a run that isn't careful about this can end up
operating on, or worse closing, a workbook or process that belongs to something else entirely.

**Decision: every run that needs a COM session creates its own dedicated Excel App instance,
and only ever tracks and acts on instances it created itself.**

- On first promotion to COM (§6.1), the runner spawns a new `xw.App` explicitly — never an
  unqualified lookup (`xw.apps.active`, or a bare `xw.Book("name.xlsx")` that could bind to any
  already-open workbook of that name anywhere on the system) that might attach to an instance
  it doesn't own — and records that instance's process ID against the run.
- The `SessionManager` (§6.2) tracks **owned App instances**, not just workbook handles — a run
  may open several workbooks, but they must all live inside App instance(s) the run itself
  spawned.
- Cleanup (§6.3's crash-safety guarantee, §6.3.1's commit/cleanup) only ever closes or kills
  App instances/processes in the run's own owned set. It must never enumerate and act on
  "every EXCEL.EXE process" or "every open workbook with this filename" system-wide — that is
  exactly the failure mode being designed against.
- This also constrains recovery from a *previous* run that crashed without cleaning up: a
  leftover orphaned Excel process can only be safely treated as "ours" if it's identifiable as
  such (e.g. its PID was recorded in that run's audit log/lock file) — blindly killing any
  idle-looking EXCEL.EXE found on the machine is not acceptable, since it might belong to a
  concurrent run or a real user. Exact tagging/recognition mechanism is an open question for
  spec phase — see §12.
- Relates to §6.3.1's flagged-but-unaddressed "two runs target the same workbook concurrently"
  cost — this section explains why that matters, without yet fully solving it.

### 6.3 Implicit workbook lifecycle

- Workbooks are **lazily opened** on first reference by any step — no mandatory `open` step.
  If an explicit `open` step is used, an optional `mode: read_only|read_write` lets the author
  override the inferred mode; the implicit (no explicit `open` step) path always uses the
  inferred mode below.
- Workbooks are **saved and closed automatically** at the end of a successful run — this also
  covers the case where an author explicitly opens a workbook and simply forgets to close it;
  the runner still closes it.
- Explicit `open`/`save`/`close` actions remain available for manual control (e.g. save
  mid-run, close early to release a lock) but are optional.
- **Crash safety is a hard requirement, not a nice-to-have.** The runner must guarantee that
  every workbook it has opened — file-backend or COM — is properly closed and every COM
  process/file lock released, even when a step raises, the process is interrupted, or an
  unhandled exception occurs mid-run. Mechanism: the run's session lifecycle is wrapped in a
  `try`/`finally` (or equivalent context-manager) at the top level so cleanup always executes
  regardless of how the run ends. The "does a crash save partial progress or not" question is
  resolved by §6.3.1's scratch-copy model, not answered here — see there.
- A **static pre-pass** over the parsed step list (before execution) determines, per workbook,
  whether it's ever used as a step's target. If never — open read-only (faster, safer). If a
  target reference can't be resolved statically (e.g. it comes from a templated expression),
  default to read-write and say so — never silently guess.

### 6.3.1 Scratch-copy execution model

**Any workbook a run will write to is copied to a scratch location first; all work happens on
the scratch copy; the real file is only touched once, at the very end, by an atomic commit —
and only if the whole run succeeded.** This makes crash-safety close to automatic rather than
something to carefully engineer step by step:

- **What gets copied**: only workbooks the static pre-pass (§6.3) determines will be written to
  — a pure-source workbook that's never a `target` doesn't need a scratch copy, it can be
  opened read-only in place.
- **COM-backed steps operate on the scratch copy too**, for the same reason — this also
  protects the real file from an orphaned/crashed Excel session, not just from a Python
  exception.
- **Commit is atomic per workbook**: write the finished scratch copy to a temp path next to the
  real target and rename over it (`os.replace` or equivalent), never overwrite the original in
  place — so a crash *during* the commit itself still can't leave the real file half-written.
- **On any failure, the real files are never touched at all** — this is what actually resolves
  the "save partial progress or not" question from §6.3: the answer is neither, because the
  real file was never in the loop until a full, successful commit. The scratch copies are left
  in place (not cleaned up) on failure, so they — plus the audit log (§6.7) — are the
  recovery/debugging artifact, which is safer than trying to reconstruct state from a log alone.
  **Correction found via a crash-safety integration test: this requires periodic checkpointing,
  not just "leave the scratch copy alone."** openpyxl writes stay in memory until an explicit
  save — nothing writes them to the scratch file on disk mid-run on its own — so without an
  explicit fix, a crash after several successful steps would leave a scratch copy no more
  informative than the untouched original: none of the in-memory progress would actually be on
  disk. Fixed by saving every dirty, staged workbook to its scratch file after *each* step
  (Specification.md §5.2/§6.1), not only at the very end — so the recovery artifact actually
  contains everything that succeeded before the crash, not just what existed at staging time.
- **On success**, scratch copies are cleaned up by default after commit; keeping them (for
  inspection, or to visualize session state during development) should be an option, not the
  default.
- Side benefit: never opening the real file for writing during the run avoids most file-locking
  contention on the original entirely — a run and a human both having a workbook open no longer
  conflicts the way it would if the run worked on the real file directly.

Cost to be honest about: extra I/O for the copy-out/commit-back, and scratch-directory
management (naming, collision-avoidance if two runs target the same workbook concurrently —
open question, not addressed yet). Likely negligible next to Excel/COM overhead itself, but
worth flagging if a very large workbook makes the copy step itself slow.

### 6.4 No per-action "step reference" field

Every step gets an `id`, and its output is automatically addressable via the templating syntax
from §10 (e.g. `{{ steps.<id>.output }}`). No action ever needs a separate "store this result"
parameter — the step `id` + templating mechanism is the one and only way results flow between
steps.

### 6.5 Action syntax design principles

1. **One field, one meaning.** A field must mean the same thing on every step that has it — a
   field name is never repurposed to mean something different depending on the action.
2. **Split an action rather than overload a field or a param's format.** If two usages of one
   action need meaningfully different parameters, they become two actions with clear names,
   not one action that branches on the shape of a string.
3. **Real YAML structures, not encoded strings — DECIDED (§10.2)** — lists are YAML lists,
   mappings are YAML mappings, conditions are structured objects (`{column: O, equals: FAIL}`),
   never semicolon/colon-packed strings requiring a hand-rolled parser. Validation (§9.1) is
   the ergonomics safety net, not a looser syntax.
4. **Formulas need no dedicated action.** Any action that writes a cell value accepts a string;
   if it starts with `=`, openpyxl stores it as a formula automatically. Caveat to document
   prominently: openpyxl cannot evaluate formulas — reading a freshly-written formula back
   returns `None`/stale cached value until a COM `recalculate` step runs.

### 6.6 Runtime parameterization

The importable Python API (§3, §9) needs to accept overrides for the YAML's own `env:` block at
call time — e.g. `run_workflow(path, env_overrides={"input_folder": "..."})` — merged over
whatever the file itself declares. This is a v1 requirement on the core API's signature, not
tied to any specific caller: the real-world use case is invoking a run with arguments supplied
by whatever external process launches it, without editing the YAML file per invocation. Which
specific outer tooling does the launching is out of scope for this PRD.

### 6.7 Audit logging

A structured, on-disk audit log — one JSON record per step (step id, action, resolved
parameters, start/end time, status, error if any) — separate from (but alongside) ordinary
human-readable console/log output. Plain `logging` module output is not a substitute: an audit
record needs to be machine-queryable (e.g. by an agent diagnosing why a run failed) as well as
human-readable, and the two have different jobs. Also the natural home for surfacing what
happened during a failed run, alongside the scratch copies from §6.3.1.

### 6.8 Error handling philosophy

Every error a user (or an agent) sees needs a plain-English message explaining what went wrong
and what to do about it, with the raw technical detail (exception type, message, traceback)
attached separately — never as the primary message. This applies to every error surface,
including runtime failures against real Excel/COM (an openpyxl exception, a COM error) — those
get translated the same way, with the technical original preserved as a secondary field in the
audit record (§6.7), not shown as the headline. See §9.1 for the bar validation errors need to
hit; runtime errors are held to the same standard.

### 6.9 Stopping a run early — DECIDED (2026-08-19)

A `stop` step halts the workflow before any later step runs — the answer to "if this lookup
comes back empty, don't bother running everything after it," without repeating the same `if:`
condition on every subsequent step. `stop` is a normal action driven by the existing `if:`
mechanism (§6.5) rather than a new per-step flag:

```yaml
- id: guard
  action: stop
  if: "{{ steps.find_it.status == 'error' }}"
```

When a `stop` step's `if:` is true (or absent, since `if:` is optional everywhere), the run
halts right there — every remaining step gets `StepResult(status="stopped")`, a status
deliberately distinct from `"skipped"` (§6.3's normal `if:`-false outcome), so the audit log can
tell "this step's own condition said don't run me" apart from "the run ended before we got
here." `stop` doesn't itself force the run to be treated as failed — whether the run commits
still depends only on whether any *earlier* step returned `status: "error"` (§6.8's
error-handling policy), unchanged. That means "not found → stop" naturally discards (the failed
lookup already set that), while a deliberate early exit on a success condition ("already
processed, nothing to do") naturally still saves whatever ran before it — no new commit logic
needed, `stop` only needs to end the loop early. See `docs/Specification.md` §6.1 for the
runner-side mechanics.

## 7. Action catalog

**Status: syntax is settled for most actions; a couple are explicitly flagged as still open.**

Three things that apply to every row below, not repeated per-row:
- **Backend (openpyxl vs. xlwings) is never a user choice** — see §6.1.
- **`range`/`column` fields accept both A1-style ranges (e.g. `"A1:C10"`) and Excel
  defined/named ranges (e.g. `"SalesData"`)**, resolved by checking the workbook's defined
  names first, falling back to A1 notation. Depth of named-range support in openpyxl/xlwings
  needs verifying in spec phase.
- **Quoting house style**: quote free-text string values — data the workflow author supplies,
  like sheet names, header patterns, cell values, file paths — because plain words like
  `yes`/`no`/`on`/`off`/`null` get silently parsed as booleans/null by some YAML loaders (YAML
  1.1 behavior, which PyYAML's default resolver follows). Don't quote fixed schema keywords
  (`action:` names, `mode`/`direction`/`target`/`operation` enum values) — those are validated
  against a known set regardless, so there's no ambiguity to guard against. Also use a YAML
  1.2-core-schema-compliant loader in spec phase, as a second line of defense.
- **Every action's `output` is always a keyed object, never a bare scalar** — see §10.4 for the
  exact shape per action.

| Action | Backend | Parameters | Notes |
|---|---|---|---|
| `open` | file/COM (whichever the workbook needs) | workbook, update_links (optional), mode: read_only\|read_write (optional override) | Optional — see §6.3 implicit open. Implicit path always uses the inferred mode; explicit `open` may override it. |
| `save` | file/COM | workbook | Optional — see §6.3 implicit save. |
| `close` | file/COM | workbook | Optional — see §6.3 implicit close. A forgotten close is still handled automatically. |
| `stop` | none — control flow only, no workbook | reason (optional string) | Halts the run before any later step runs — see §6.9. Not tied to any workbook; `workbook:` isn't a field on this action. |
| `copy` | file | source: {workbook, sheet, range}, target: {workbook, sheet, range} | `range` optional on `source` — omit for the whole sheet. |
| `read_range` | file | workbook, sheet (string, or a list for multi-sheet, or `all`), range, as: values\|formulas (optional) | Multi-sheet: an explicit list (`["North", "South"]`) is the real mechanism; `all` is authoring sugar that expands to the full sheet list before execution — one code path underneath either way. |
| `read_metadata` | file (properties/cells) or COM (textboxes) | workbook, target: properties\|textboxes\|cells, sheet (required if target=cells — clarification found during implementation, missing from the original catalog), cells (list, if target=cells) | Two distinct things live behind one action: (a) workbook/document properties (author, title, custom doc properties — file-backend), (b) the current value of an embedded ActiveX/form control (COM-only — openpyxl can't see live control state). Capability depends on `target:`, not fixed per-action (§6.1 exception). |
| `write_cell` | file | workbook, sheet, cell, value | Single-cell write. `value` can be a literal or a step-output reference (§10), resolved by the templating engine — no special encoding needed. |
| `write_range` | file | workbook, sheet, range, values (2D list) | Writes a computed block of values in one shot — the gap between single-cell `write_cell` and row/table-oriented `write_row`/`write_table`. |
| `write_row` | file | workbook, sheet, row, values: {column: value, ...} — or values_by_header + headers_from — or start_column + positional values | Three modes: explicit column-letter mapping (base form), by-header-name (using a prior `find_columns` step's output), and positional (an ordered list from a start column, no mapping at all). All three in v1. |
| `write_table` | file | workbook, sheet, source (step ref(s)), headers (list), filter: {column, equals} (optional) | Takes one or more prior steps' same-shaped tabular output, optionally filters rows, writes headers + surviving rows to a target sheet. |
| `insert_range` | file | workbook, sheet, at (e.g. `"C:C"` or `"C5:C10"`), direction: rows\|columns (see rule), header: {row, text} (optional) | **Direction rule**: for a whole column/row (`"C:C"`, `"5:5"`), `direction` is unambiguous from the syntax and optional. For a true partial range (`"C5:C10"`), `direction` is **required** — never inferred, since guessing risks silently doing the wrong thing. Whole-row/column is native to openpyxl and cheap; partial-range insert-with-shift needs hand-rolled cell-shifting logic. |
| `set_column_width` | file | workbook, sheet, columns, width: number\|"autofit" | |
| `find_headers_row` | file | workbook, sheet, search_range, patterns (list) | Returns the row number and which pattern matched which column — see §10.4. |
| `find_row` | file | workbook, sheet, column, search_value, header_row (optional) | |
| `find_column` | file | workbook, sheet, header_row, pattern | Single pattern → single column. |
| `find_columns` | file | workbook, sheet, header_row, patterns: {name: pattern, ...} | Multiple named patterns → a mapping of logical name to column letter, in one step. |
| `aggregate` — **flagged for discussion when we get to it, not resolved now** | file (pandas, in-memory) | source (step ref), group_by, value_column, operation: count\|sum\|avg | No workbook access — pure in-memory aggregation of a prior step's data. |
| `update_summary_table` — **kept as a dedicated action; exact parameters deliberately deferred, not a phase-1 design task** | file | *(not designed yet)* | Conceptually composes `find_headers_row` + `find_row` + `find_columns` + `aggregate` + `write_row` — see §11 item 19 for a worked composed-vs-wrapper comparison kept for reference. |
| `read_links` | file — **status downgraded, see below** | workbook | Reads current external-link target paths. |
| `write_links` | COM | workbook, links: {link_id or old_path: new_path} | Rewrites external-link target paths. COM-backed (`ChangeLink`/`LinkSources`) rather than file-backend — openpyxl's write support for external links isn't confirmed reliable enough to depend on. Enables the move-rewrite-refresh-restore scenario in §11 item 18. |
| `refresh_links` | COM | workbook | Forces Excel to actually re-pull values through the links — distinct from just changing where they point (`read_links`/`write_links` above). |
| `recalculate` | COM | workbook, mode: normal\|full\|full_rebuild (optional, default normal) | `mode: normal` uses xlwings' portable `app.calculate()` (works on Mac and Windows). `full`/`full_rebuild` require the raw COM object (`app.api.CalculateFull()`/`CalculateFullRebuild()`) — **Windows-only**; requesting either on Mac raises a clear unsupported-platform error, never a silent downgrade. |
| `run_macro` | COM | workbook, macro_name, args (optional list) | |
| `export_pdf` — **backlog, not scheduled for a phase yet** | COM | workbook, sheet or range, output_path | |

## 8. v1 vs. later phases

Updated against actual build progress (Specification.md §8/§4 track this in detail):

- **v1, built**: open, save, close, copy, read_range (single-range mode; multi-sheet mode still
  TBD, see §7), read_metadata (properties/cells sub-cases), write_cell, write_range, write_row
  (base column-mapping + positional modes), insert_range (whole-row/whole-column only —
  partial-range returns a structured error, not built), set_column_width, find_headers_row,
  find_row, find_column, find_columns.
- **v1, but COM-required rather than file-backend**: read_metadata (textbox-control sub-case
  only — openpyxl can't see live control state). Not built yet — COM phase.
- **Deferred — needs cross-step data access `runner.py` doesn't provide yet**: `write_table`
  (its `source: [step_ids]` param means "look up these steps' outputs", not a literal value),
  `write_row`'s by-header mode (`headers_from` is the same kind of step-id reference),
  `aggregate` (its `source` param is likewise a step reference — this compounds its existing
  "flagged for discussion" status from §7/§11.17). All three become buildable once `runner.py`
  threads accumulated step-output context through to action calls.
- **Deferred — empirically confirmed openpyxl limitation, see §7's `read_links` note**:
  `read_links`, alongside the already-deferred `write_links`.
- **Later phase (COM promotion)**: recalculate, run_macro, refresh_links, write_links.
- **Kept, syntax deferred (not phase-1 design work)**: update_summary_table.
- **Backlog**: export_pdf.

## 9. AI-agent authoring experience (goal, not deferred)

This is explicitly **as important as the engine working correctly** — not a nice-to-have
layered on top once the "real" work is done.

Target experience: a user describes what they want in plain language; an AI agent inspects the
actual workbook(s) involved (sheet names, headers, structure, sample data) and writes — or
expands a partial draft into — a complete, runnable `workflow.yaml` from that description,
without the user having to know the YAML syntax themselves.

Division of responsibility:
- **Inspection tooling is the agent's job, not excel_runner's, for v1.** An agent can inspect a
  workbook today with generic tools (read the file, run a snippet of code against it) without
  excel_runner providing anything special.
- **Once the core engine works, dedicated inspection actions get built as the next phase** —
  not indefinitely backlogged. Candidates: `list_sheets`, `describe_sheet` (headers, dimensions,
  a sample of rows), reusing `read_range` for ad-hoc inspection. Exact scope for spec phase.

Design consequence for everything in §6.5/§7: an action's ergonomics should be judged partly by
"would an agent reliably produce correct YAML for this from a short natural-language request,"
not only by "is this technically sufficient." Tier-1 schema validation needs to produce
specific, actionable error messages an agent can self-correct from — this is part of the
authoring experience, not just a safety net.

### 9.1 Validation error examples (making the promise concrete)

The claim above only means something if the errors are actually this specific. Examples of the
bar to hit:

```
Step "write_fail_summary" (action: write_table): field "headers" must be a list,
e.g. [Name, Status, Value] — got the string "Name, Status, Value". Wrap it in [ ].
```

```
Step "get_regional" (action: read_range): field "sheet" is "North, South" (a string) —
did you mean a list? [North, South]
```

```
Step "totals_by_region" (action: aggregate): field "source" references step id
"get_regional_data", which does not exist. Did you mean "get_regional" (defined at
step 3)?
```

```
Step "copy_named_range" (action: copy): source.range "ReservingTotal" is not a valid
A1 range and does not match any defined name in workbook "historical". Close matches:
"ReservingTotals".
```

**Corrected during implementation of tier-1 validation (§10.4's build): the fourth example
above is not actually catchable statically.** "Does not match any defined name in workbook
historical" requires opening that workbook to enumerate its defined names — which contradicts
this section's own claim, two lines up, that "no workbook needs to be opened to produce any of
them." The first three examples are genuinely workbook-access-free (list-shape checks, a
step-id reference check) and are what tier 1 actually implements. Checking a range against a
workbook's *real* defined names needs either a third validation tier that opens workbooks
read-only for checking purposes only (not designed), or gets demoted to a plain runtime error
when an action actually tries to use the range and openpyxl can't resolve it. Not decided —
carried to §12 as an open item, not silently dropped.

The first three examples are the actual bar for tier-1 validation: name the field, show the
expected shape, suggest the fix — not just "invalid config" or a raw pydantic traceback.

## 10. Syntax conventions

### 10.1 Templating expression syntax — DECIDED

**Decision: Jinja2 syntax (`{{ }}`), rendered by the actual `jinja2` library, resolved
per-field rather than as one whole-file text pass.** Picked on technical merit:

- **Reuses a mature, battle-tested library instead of hand-rolling a parser.** The syntax *is*
  Jinja2, so rendering and `if:` comparisons come from the `jinja2` package directly — one
  dependency, not a bespoke parser to maintain.
- **Corrected during implementation: not a whole-file render.** The original plan here was
  "render the whole YAML file as one Jinja2 pass, then parse the result as YAML" — building it
  surfaced that this can't actually work: at load time no step has run yet, so a `{{ steps.x }}`
  reference anywhere in the file has nothing to resolve against, and a real whole-file render
  would either error on every one of them (with the strict, fail-fast undefined-checking we
  want) or silently blank them out (which we don't want). Neither GitHub Actions nor Ansible
  actually do a single whole-document text substitution either — both parse the structure first,
  then evaluate each expression per field, in whatever context is available at that point. The
  built mechanism does the same: YAML is parsed directly (`{{ }}` is just string content to a
  YAML parser, no conflict), then `env:`/`workbooks:` fields are resolved once at load time
  (env-only context), and step `params`/`if:` are left raw and resolved per step during
  execution (env + accumulated step-output context) — see docs/Specification.md §2.2. The
  computed-dict-key case below still works under this model, more directly: resolving a step's
  `params` just before it runs already recurses through that dict's keys and values.
- **Native type preservation, Ansible-style.** A naive Jinja2 render always produces a string,
  which would make `row: {{ steps.find_target_row.output.row }}` render to the text `"5"`
  rather than the integer `5`. Ansible solves exactly this with "native Jinja2 types": when a
  field's value is *entirely* one `{{ }}` expression (nothing else in the string), substitute
  the actual Python object, not its string form; only stringify when the expression is embedded
  inside a larger string (e.g. a file path with one path segment substituted). Adopt the same
  rule — a proven solution to a problem we'd otherwise have to solve ourselves.

Reference examples:

```yaml
- id: recalc
  action: recalculate
  workbook: manip
  mode: full
  if: "{{ steps.copy_data.status == 'success' }}"
```

```yaml
- id: write_result
  action: write_row
  workbook: manip
  row: "{{ steps.find_target_row.output.row }}"
  values: { B: "{{ steps.totals.output.North }}" }
```

(`.output.row` and `.output.North` reference specific fields of a step's output — see §10.4 for
what shape every action's output actually takes.)

**Finding from writing the `update_summary_table` composition example (§11.19):** one version
of that example needs a *computed dict key* — a column letter found by an earlier step used as
the key of a `values:` mapping, not just a value:

```yaml
values: { "{{ steps.find_key_columns.output.total }}": "{{ steps.totals.output.North }}" }
```

Resolving `params` per-field (rather than any whole-document substitution) is what makes this
legal — the computed key is just another value `resolve_value` walks through on its way to
building the step's final, resolved parameter dict.

### 10.2 List/dict ergonomics for hand-authors — DECIDED

**Decision: real YAML structures everywhere — no parallel loose-string parsing mode.** §6.5
principle #3 stands as written. The safety net is validation, not a more forgiving syntax: both
the tier-1 static dry-run and, as a final backstop, a check immediately before a real run
touches any workbook, must catch a malformed field and explain exactly what's wrong and how to
fix it (§9.1's examples are the bar to hit). Given §9's AI-authoring goal, most hand-editing
will be tweaking an already-well-formed, agent-generated file rather than writing a workflow
from a blank page, which keeps the real-world error surface for this small.

### 10.3 Summary of §10 decisions

- **10.1 (templating syntax): decided** — Jinja2 `{{ }}`, resolved per-field (not a whole-file
  render — corrected during implementation, see §10.1), native type preservation.
- **10.2 (list/dict ergonomics): decided** — real YAML structures always; validation is the
  ergonomics fix.

### 10.4 Step output shapes reference

**Rule: every action's `output` is always a keyed object, never a bare scalar** — even when
there's only one meaningful value — so referencing it is always `steps.<id>.output.<field>`,
never sometimes-bare/sometimes-not. Keys are either fixed (known at design time) or dynamic
(data-dependent, e.g. group values), noted per action:

| Action | `output` shape |
|---|---|
| `read_range` | `{ values: <2D array, or scalar if a single cell> }` |
| `read_metadata` | `{ <requested property or textbox name>: <value>, ... }` (dynamic keys) |
| `find_headers_row` | `{ row: <int>, headers: { <pattern>: <column letter>, ... } }` |
| `find_row` | `{ row: <int> }` |
| `find_column` | `{ column: <letter> }` |
| `find_columns` | `{ <logical name from patterns:>: <column letter>, ... }` (dynamic keys, fixed by the step's own `patterns:` input) |
| `aggregate` | `{ <group value found in the data>: <aggregated value>, ... }` (fully dynamic keys — e.g. `.output.North` only makes sense if "North" is an actual value that appeared in the `group_by` column) |
| `read_links` | `{ <link identifier>: <current target path>, ... }` (dynamic keys) |
| `run_macro` | whatever the macro returns, if anything — shape TBD, COM detail for spec phase |
| `open`/`save`/`close`/`copy`/`write_cell`/`write_range`/`write_row`/`write_table`/`insert_range`/`set_column_width`/`refresh_links`/`recalculate` | no meaningful `.output` — use `.status`/`.error` in `if:` conditions instead |

## 11. Action examples — full capability reference

Registry header used throughout:

```yaml
env:
  input_folder: "./input"
  output_folder: "./output"

workbooks:
  historical:
    file: "{{ env.input_folder }}/2025 Historical analysis v1.xlsx"
  manip:
    file: "{{ env.output_folder }}/Backtesting Manip.xlsx"
    create_if_missing: true
  results:
    file: "{{ env.output_folder }}/Results.xlsx"
    create_if_missing: true
```

**1. `open`**
```yaml
- id: open_hist
  action: open
  workbook: historical
  mode: read_only   # optional override — omit to let the runner infer (§6.3)
```

**2. `save`**
```yaml
- id: save_manip
  action: save
  workbook: manip
```

**3. `close`**
```yaml
- id: close_hist
  action: close
  workbook: historical
```

**4. `copy`** — whole sheet, then a named range into a specific cell
```yaml
- id: copy_sheet
  action: copy
  source: { workbook: historical, sheet: "Reserving Data" }
  target: { workbook: manip, sheet: "Reserving Data" }

- id: copy_named_range
  action: copy
  source: { workbook: historical, sheet: "Reserving Data", range: "ReservingTotals" }
  target: { workbook: manip, sheet: "Summary", range: "B2" }
```

**5. `read_range`** — values, formulas, and multi-sheet mode
```yaml
- id: get_totals
  action: read_range
  workbook: manip
  sheet: "Outputs"
  range: "A1:D50"

- id: get_formulas
  action: read_range
  workbook: manip
  sheet: "Model"
  range: "C2:C10"
  as: formulas

# multi-sheet: explicit list is the real mechanism
- id: get_regional
  action: read_range
  workbook: manip
  sheet: ["North", "South", "East", "West"]
  range: "B2:D20"

# "all" is authoring sugar — expands to the full sheet list before execution, same code path
- id: get_all_regions
  action: read_range
  workbook: manip
  sheet: all
  range: "B2:D20"
```

**6. `read_metadata`**
```yaml
- id: get_doc_props
  action: read_metadata
  workbook: historical
  target: properties

- id: get_textbox_value
  action: read_metadata
  workbook: historical
  target: textboxes
```

**7. `write_cell`** — literal and formula
```yaml
- id: set_status
  action: write_cell
  workbook: manip
  sheet: "Summary"
  cell: "B2"
  value: "Complete"

- id: set_formula
  action: write_cell
  workbook: manip
  sheet: "Model"
  cell: "D10"
  value: "=SUM(D2:D9)"
```

**8. `write_range`**
```yaml
- id: write_block
  action: write_range
  workbook: manip
  sheet: "Summary"
  range: "B2:D4"
  values:
    - [10, 20, 30]
    - [40, 50, 60]
    - [70, 80, 90]
```

**9. `write_row`** — base form, plus the two convenience modes
```yaml
- id: write_summary_row
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  values: { B: "North", C: 1200, D: "PASS" }

# by header name, using a prior find_columns step's output
- id: write_summary_row_by_header
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  headers_from: find_key_columns
  values_by_header: { Region: "North", Total: 1200, Status: "PASS" }

# positional — just the values, in order, from a start column
- id: write_summary_row_positional
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  start_column: B
  values: ["North", 1200, "PASS"]
```

**10. Creating a new workbook** — no dedicated action. The `workbooks:` registry already
requires declaring a workbook upfront (see the header example above); `create_if_missing: true`
plus implicit lazy-open (§6.3) creates it the moment it's first referenced. Template selection
is a registry-entry field:
```yaml
workbooks:
  results:
    file: "{{ env.output_folder }}/Results.xlsx"
    create_if_missing: true
    template: historical   # optional
```

**11. `write_table`** — takes one or more prior steps' tabular output, optionally filters rows,
writes headers + the surviving rows to a target sheet. Both source steps below produce
same-shaped rows (same columns, different sheets) so they stack into one table:
```yaml
- id: get_north
  action: read_range
  workbook: manip
  sheet: "North"
  range: "A2:C50"

- id: get_south
  action: read_range
  workbook: manip
  sheet: "South"
  range: "A2:C50"

- id: write_fail_summary
  action: write_table
  workbook: results
  sheet: "Failures"
  source: [get_north, get_south]
  headers: ["Name", "Status", "Value"]
  filter: { column: Status, equals: FAIL }
```

**12. `insert_range`** — cheap whole-column case, then the costlier partial-range case.
**Direction rule**: when `at` is a whole column (`"C:C"`) or whole row (`"5:5"`), `direction` is
unambiguous from that syntax alone and can be inferred — no guessing involved, so `direction`
is optional there. For a true partial range (e.g. `"C5:C10"`), `direction` is **required**, not
inferred — silently guessing "tall means rows" for an ambiguous or square block risks doing the
wrong thing without telling you.
```yaml
- id: insert_flag_column
  action: insert_range
  workbook: manip
  sheet: "Summary"
  at: "C:C"
  header: { row: 1, text: "Flag" }   # direction omitted — unambiguous from "C:C"

- id: insert_partial_block
  action: insert_range
  workbook: manip
  sheet: "Summary"
  at: "C5:C10"
  direction: rows   # required here — "C5:C10" alone doesn't say which way to shift
```

**13. `set_column_width`**
```yaml
- id: widen_columns
  action: set_column_width
  workbook: manip
  sheet: "Summary"
  columns: "B:D"
  width: autofit
```

**14. `find_headers_row`.** String list items quoted per house style — plain words like
`yes`/`no`/`on`/`off` get silently parsed as booleans by some YAML loaders, and a real header
pattern could collide with one of these:
```yaml
- id: find_summary_headers
  action: find_headers_row
  workbook: manip
  sheet: "Summary"
  search_range: "A1:J10"
  patterns: ["Region", "Total", "Status"]
```

**15. `find_row`**
```yaml
- id: find_north_row
  action: find_row
  workbook: manip
  sheet: "Summary"
  column: B
  search_value: "North"
  header_row: 1
```

**16. `find_column` / `find_columns`**
```yaml
- id: find_status_col
  action: find_column
  workbook: manip
  sheet: "Summary"
  header_row: 1
  pattern: "Status"

- id: find_key_columns
  action: find_columns
  workbook: manip
  sheet: "Summary"
  header_row: 1
  patterns: { region: "Region.*", total: "Total.*", status: "Status" }
```

**17. `aggregate` — flagged for discussion when we get to it, not resolved now.**
```yaml
- id: totals_by_region
  action: aggregate
  source: get_regional
  group_by: "Region"
  value_column: "Value"
  operation: sum
```

**18. `read_links` / `write_links` / `refresh_links`** — the move-rewrite-refresh-restore
scenario. Restoring replays `read_links`' captured output directly, rather than hand-typing the
reverse mapping:
```yaml
- id: original_links
  action: read_links
  workbook: manip

- id: repoint_links
  action: write_links
  workbook: manip
  links:
    "\\\\server\\shared\\Source.xlsx": "./scratch/Source.xlsx"

- id: refresh
  action: refresh_links
  workbook: manip

- id: restore_links
  action: write_links
  workbook: manip
  links: "{{ steps.original_links.output }}"
```

**19. `update_summary_table` — kept as a dedicated action; exact syntax deferred, not a
phase-1 design task.** Composed-from-primitives vs. wrapper comparison kept below for
reference:

Composed (5 steps, for reference):
```yaml
- id: find_headers
  action: find_headers_row
  workbook: manip
  sheet: "Summary"
  search_range: "A1:J5"
  patterns: ["Region", "Total"]

- id: find_target_row
  action: find_row
  workbook: manip
  sheet: "Summary"
  column: B
  search_value: "North"

- id: find_key_columns
  action: find_columns
  workbook: manip
  sheet: "Summary"
  header_row: "{{ steps.find_headers.output.row }}"
  patterns: { total: "Total" }

- id: totals
  action: aggregate
  source: get_regional
  group_by: Region
  value_column: Value
  operation: sum

- id: write_result
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: "{{ steps.find_target_row.output.row }}"
  values: { "{{ steps.find_key_columns.output.total }}": "{{ steps.totals.output.North }}" }
```
(This is the example that surfaced the computed-dict-key question in §10.1.)

As a wrapper instead (1 step) — shorter, and sidesteps the computed-key question entirely
because it's resolved inside the action's own code, not in raw YAML:
```yaml
- id: update_north_total
  action: update_summary_table
  workbook: manip
  sheet: "Summary"
  search_range: "A1:J5"
  header_patterns: ["Region", "Total"]
  search_column: B
  search_value: "North"
  aggregate: { source: get_regional, group_by: Region, value_column: Value, operation: sum }
```

**20. `recalculate`**
```yaml
- id: recalc
  action: recalculate
  workbook: manip
  mode: full
```

**21. `run_macro`**
```yaml
- id: refresh_model
  action: run_macro
  workbook: manip
  macro_name: "RefreshModel"
  args: [2026, "Q1"]
```

**22. `export_pdf`** (backlog — shown for shape only)
```yaml
- id: export_summary
  action: export_pdf
  workbook: manip
  sheet: "Summary"
  output_path: "{{ env.output_folder }}/summary.pdf"
```

## 12. Open questions carried to spec phase

Still open:
- **openpyxl silently drops charts it can't fully parse, on any save (found 2026-08-19,
  evidence-backed, not folklore)** — `open`'s load, followed by any `save`, means any chart
  openpyxl's reader can't fully parse is gone from the file afterward, even if the workflow
  never referenced charts at all: openpyxl's chart *reader* catches a parse failure and warns
  rather than preserving the original part, so a re-save simply omits whatever it couldn't
  round-trip. Confirmed real via openpyxl's own documented reader-warning path (`"Unable to
  read chart {rel.id}..."`) and a local test: a chart openpyxl itself created survives our own
  `open`→`save` cycle twice over, but that's the friendly case — a chart authored by real Excel
  using a feature openpyxl's reader doesn't support is the actual risk, and confirming that
  needs a real Excel-generated fixture, the same limitation already blocking `read_links`/
  `write_links` (see §7's note, and `docs/Specification.md` §0 for the private-fixture-notes
  convention). Not a feature request — a risk to the core save path's basic promise, that a
  workflow which never touches charts could still silently strip them from the real file.
  Buildable mitigation, not designed yet: openpyxl's read failures are catchable warnings, so
  `open` could detect "N chart(s) I can't fully parse" and surface it as a structured warning
  before anything is saved, instead of losing them silently. Same question likely applies to
  pivot tables — not verified either way yet, and out of scope to chase down unless/until a
  workflow actually needs to touch one (see the pivot-table note below).
- **"Replay nice" mode (backlog idea, not committed)** — a desktop-comfort feature for users who
  see other Excel-MCP-style tools drive Excel live via COM and associate that visible activity
  with trustworthiness, even though (§6.1) that's not how this engine works and isn't going to
  be (raised 2026-08-19). Rejected approach: giving file-backend actions an alternate xlwings
  implementation — that inverts §6.1's "backend is never a user choice" decision and fights the
  scratch-copy safety model (§6.3.1), for a payoff that's purely psychological. Preferred
  approach instead: after a real run finishes safely via the file backend, open the
  already-committed result in visible Excel via xlwings and *replay* it — select/flash each
  step's affected range with a short pause, narrate using `ActionSpec.description` + the audit
  log's per-step output (§6.7), and visibly surface `error`/`skipped`/`stopped` steps rather than
  only the happy path (a `stop` step's reason, PRD §6.9, is a good moment to show the safety net
  actually working). This consumes an already-finished `RunResult` and only re-displays
  already-computed values — it never becomes a second execution path, needs no crash-safety
  design of its own, and doesn't touch `run_workflow`'s contract. Desktop/COM-only by nature,
  strictly opt-in, never part of an automated pipeline. A related but explicitly separate idea —
  pivot tables, charts, or other new *capabilities* some Excel-MCP tools offer — is a different
  axis (more actions, not better visibility into existing ones) and is explicitly **not** part
  of this idea; avoid something requested wanted before it's actually asked for.
- **A conversational, agent-driven spreadsheet-authoring product (adjacent idea, likely a
  different product, not this project's scope — raised 2026-08-19)** — distinct from §9's
  existing v1-adjacent goal (an agent inspects a workbook once and authors a complete, static
  `workflow.yaml` from a plain-language description, which then runs deterministically). What's
  being flagged here is more ambitious: a live, iterative "type an instruction, see it happen,
  ask for the next tweak" experience for building up a spreadsheet conversationally, closer to
  an Excel-copilot product than a workflow-automation engine. `excel_runner`'s action catalog and
  `list_actions()`'s schema-generation (§6.1/§6.3) would likely be reusable substrate for
  something like this later, but designing for it now would be building ahead of any real
  request — noted for the record, not scoped, not committed.
- **Grouped `if:` blocks (backlog idea, not committed)** — today, several steps that all depend
  on the same condition each repeat their own `if: "{{ ... }}"`. A block-level `if:` wrapping
  multiple steps would remove the repetition, but adds YAML nesting/structure for a problem
  `stop` (§6.9) may already cover well enough in practice. Raised alongside the `stop` design
  (2026-08-19); deliberately not designed further until real workflows show repeated `if:`
  conditions are an actual recurring pain, not just a theoretical one — possible over-engineering
  if `stop` turns out to be sufficient.
- **xlwings/Excel App instance ownership tagging mechanism (§6.2.1)** — exact way to record
  and recognize which running Excel process(es) belong to a given run, robust enough to
  survive a crash (so a *later* run or a cleanup utility can identify and safely reap an
  orphaned instance from a previous crashed run, without ever touching an instance it doesn't
  positively recognize as its own). PID tracking alone isn't enough across a crash-and-restart
  — needs a concrete design.
- **Scratch-directory collision-avoidance (§6.3.1)** — naming/locking scheme if two runs
  target the same workbook concurrently. Not addressed yet.
- **§9.1's fourth validation example (checking a range against a workbook's real defined
  names)** — not implementable in either validation tier as designed (both are explicitly
  workbook-access-free). Needs either a new tier that opens workbooks read-only for validation
  only, or demotion to a runtime error. Found during tier-1 validation's implementation
  (Specification.md §5.4); not decided.

Everything else originally tracked here is resolved. Kept for the record, not because
anything is still pending:

- ~~Environments/secrets handling~~ — **backlogged.** No env/secrets manager for now; for
  moving between environments, manual zip/move/unzip/clone of folders is acceptable for now.
- ~~Packaging/distribution~~ — not building a distribution mechanism yet; out of scope until
  the engine itself is proven.
- ~~COM backend testing strategy across platforms~~ — **resolved**: test what's testable on
  macOS now, finish COM-path integration testing on Windows later. Accepted as the plan (§4).
- ~~Whether the engine needs to be importable as a plain Python library~~ — **resolved: yes**,
  a v1 architectural requirement (§3, §9). Follow-up requirement this creates: a deliberately
  curated public API surface (e.g. an `api.py` with the stable importable symbols), separate
  from internals — so other projects, and a future MCP/CLI wrapper, depend on a stable surface
  rather than reaching into implementation details.
- ~~Full templating/expression grammar depth~~ — **resolved by §10.1**: real Jinja2 gives
  `if:` comparisons/booleans for free, no bespoke evaluator to design.
- **Whether openpyxl supports reading/writing external-link target paths — reopened, was
  wrongly marked resolved.** The earlier "reading looks solid" call was based on reading
  openpyxl's docs, not on testing it. Empirical test during implementation: writing a formula
  string containing an external reference (`=[Source.xlsx]Sheet1!A1`) via openpyxl and
  reopening the file leaves `workbook._external_links` empty — openpyxl doesn't create the
  underlying external-link relationship from a formula at all, it only reflects one that
  already exists in a file created by real Excel. Reading may still work *given* such a file,
  but there's no way to build or verify that without a real Excel-generated fixture (or manual
  zip/XML surgery to fabricate one) — a real task, not a quick check. `read_links` is
  downgraded from "resolved: file-backend" to **deferred, same bucket as `write_links`**, until
  that fixture work happens. Implementation notes in a private, gitignored file — see
  `docs/Specification.md` §0.
- ~~Exact xlwings API for recalculation depth~~ — **resolved**: `mode: normal` uses
  `app.calculate()` (cross-platform); `full`/`full_rebuild` require the raw COM object and are
  Windows-only (§7).
- ~~Multi-sheet `read_range` syntax~~ — **resolved**: explicit list is the mechanism, `all` is
  sugar over it (§7, §11 item 5).
- ~~`write_range`~~ — **resolved: keep**, gap confirmed real (§7, §8).
- ~~`write_row`'s header-based/positional convenience modes~~ — **resolved: keep both in v1**
  (§7).
- ~~`update_summary_table`~~ — **resolved: keep** as a dedicated action; exact syntax
  deliberately deferred, not a phase-1 design task (§7, §8, §11 item 19).
- ~~`read_metadata`~~ — **resolved: keep**, real usage confirmed (§7, §8).
