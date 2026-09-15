# Excel Runner

Excel Runner is declarative, YAML-driven Excel automation. A workflow describes workbook reads,
writes, lookups, copies, and control flow as data rather than as a bespoke Python script.
The runner selects its workbook backend automatically, stages work in a scratch directory,
validates workflows before changes begin, commits successful outputs atomically, records an
audit log, and cleans up the Excel instances it owns.

> **First-run safety:** never point a new or changed workflow at a shared folder. Copy the
> input workbooks locally, run `--dry-run` and `--check-existence`, validate the local output,
> then change to approved production paths.

## Install and run

### Script only deployment (non-package)

Deploy the Python files and `requirements.txt` in the chosen directory. The workflow YAML file
can have any name and be stored in any folder.

```text
cli.py
core.py
runner.py
engine.py
actions.py
backends.py
requirements.txt
workflow.yaml
```

Install the runtime dependencies, then run the workflow from that directory:

```powershell
python -m pip install -r requirements.txt
python cli.py workflow.yaml --dry-run
python cli.py workflow.yaml --check-existence
python cli.py workflow.yaml
```

`requirements.txt` is the runtime dependency for this source deployment. Use local
workbook copies for the first two commands and inspect the output before a production run.

### Development and contributing (package deployment)

For package development, create a virtual environment and install the project with its
development dependencies. If you are unfamiliar with Python virtual environments, refer to the
Unify Python best-practice document or an equivalent trusted guide.

```powershell
git clone <this repo>
cd excel-runner
python -m venv .venv
.venv\Scripts\python -m pip install -e ".[dev]"
```

Package interfaces remain available for development and compatibility:

```powershell
.venv\Scripts\python -m excel_runner workflow.yaml
```

### Command-line options

- `--env KEY=VALUE`: Override an `env:` value; repeat for multiple values.
- `--working-dir PATH`: Place `excel_runner_runs/<workflow-name>/` under `PATH`. This is where
  temporary files plus audit and run logs go.
- `--logging-level DEBUG|INFO|WARNING|ERROR`: Set console and log-file verbosity.
- `--no-logfile`: Do not create `run.log` beside the audit log.
- `--dry-run`: Validate without staging, changing, saving, or recalculating workbooks.
- `--check-existence`: Validate referenced workbook files, sheets, and defined names before a
  real run.

### Preflight and validation

Use the compiler-style read-only preflight checks before a run:

```powershell
# Validate only. No staging, Excel process, save, calculation, or action dispatch.
.venv\Scripts\python .\excel_runner\cli.py .\workflow.yaml --dry-run

# Run the same preflight, then execute only when it succeeds.
.venv\Scripts\python .\excel_runner\cli.py .\workflow.yaml --check-existence

# Suppress the default human-readable log file for either command.
.venv\Scripts\python .\excel_runner\cli.py .\workflow.yaml --no-logfile
```

A workflow is checked in up to three tiers before/while it runs:

1. **Structural** (always on) — every action exists, no unrecognized or missing-required
   params, param types match, step-id references resolve in order. No workbook access at all.
2. **Planning** (always on) — infers whether each workbook needs to be opened read-only or
   read-write, from which actions touch it. Still no workbook access.
3. **Read-only preflight** (`--dry-run`, or before execution with `--check-existence`) — opens
  referenced workbooks read-only and confirms literal sheet names, workbook-level defined
  names, and table boundaries/header/lookup references. It also validates literal text input
  files by parsing them and compiles literal regular expressions. Values derived from earlier
  step output are deferred to execution because their final values do not exist yet. A workbook
  that does not exist yet (`create_if_missing` with no template) is skipped.

The CLI writes all console log records to `excel_runner_runs/<workflow-name>/run.log` by default,
alongside the structured `audit.jsonl` created during execution. Pass `--no-logfile` to disable
the human-readable log. A default `--dry-run` creates only this `run.log`; it still does not
create a scratch copy, Excel process, workbook write, save, calculation, or action dispatch.
`--check-existence` performs the same checks, then executes normally. Validation errors stop the
real run before any action dispatch.


## Run a workflow from a script:

```python
from excel_runner import run_workflow

result = run_workflow("workflow.yaml")
print(result.status)          # "success" or "error"
for step in result.step_results:
    print(step.step_id, step.status)
```

Pass `env_overrides` to parameterize a run without editing the file:

```python
run_workflow("workflow.yaml", env_overrides={"output_folder": "/tmp/run-42"})
```

## Using it as a library

```python
from excel_runner import run_workflow, list_actions, RunResult, StepResult

result: RunResult = run_workflow("workflow.yaml")

for spec in list_actions():
    print(spec.name, "-", spec.description)
```

`list_actions()` returns every built action's name, description, capability, and parameter
schema — useful for building a tool wrapper (CLI, API and handoff, MCP server, agent framework) on top without
duplicating the action catalog.


## Quick start

A minimal workflow: read one cell, write it somewhere else.

```yaml
workbooks:
  my_book:
    file: "./my_book.xlsx"

steps:
  - id: get_total
    action: read_range
    workbook: my_book
    sheet: "Summary"
    range: "B2"

  - id: write_total
    action: write_cell
    workbook: my_book
    sheet: "Summary"
    cell: "D2"
    value: "{{ steps.get_total.output.values }}"
```

The workbook opens automatically and is saved automatically at the end if every step succeeds;
no explicit `save` step is needed (see [Workbook lifecycle](#workbook-lifecycle)).

## Workflow YAML explained

A workflow file has three top-level blocks:
* `env`: Variables set before the workflow starts, such as paths. The same values can be
  overridden from the command line or library code.
* `workbooks`: Logical workbook names used by steps. See the action reference for options such
  as templates.
* `steps`: Ordered actions. Each step needs a tracking `id` containing letters, numbers, and
  underscores, with no spaces and no leading number.

```yaml
env:                              # optional — plain values, referenced as {{ env.NAME }}
  input_folder: "./input"         # These can be overridden in the command line
  output_folder: "./output"

workbooks:                        # every workbook the workflow touches, by logical name
  historical:
    file: "{{ env.input_folder }}/historical.xlsx"
  results:
    file: "{{ env.output_folder }}/results.xlsx"
    template: historical          # begin every run from this workbook's content

steps:                            # run in order
  - id: some_step                 # unique, referenced by later steps
    action: read_range            # one of the actions below
    workbook: historical          # logical name from workbooks: above
    sheet: "Summary"
    range: "A1:D10"
```

Workflows can be authored manually or with GitHub Copilot using the
[`excel-runner-yaml` skill](.github/skills/excel-runner-yaml/SKILL.md). The skill is the
authoritative field-by-field YAML syntax reference; this README is an introduction and action
overview.

### Referencing another step's output

Values use the Jinja2 templating syntax. A `{{ ... }}` expression can refer to an `env` value or
the output of an earlier step.

Every step's result is available as `{{ steps.<id>.output }}` once that step has run:

```yaml
value: "{{ steps.get_total.output.values }}"
```

If a field's value is *entirely* one `{{ }}` expression, it resolves to the real Python value
(a number stays a number) rather than a string. Embedded in a longer string
(`"{{ env.output_folder }}/results.xlsx"`), it's substituted as text.

### Conditional steps

```yaml
- id: maybe_write
  action: write_cell
  workbook: results
  sheet: "Summary"
  cell: "E1"
  value: "found it"
  if: "{{ steps.find_it.status == 'success' }}"
```

A step whose `if:` evaluates false is skipped (not run, not an error). Every step's `status`
(`success`, `error`, `skipped`, or `stopped` — see [`stop`](#stop)) is available the same way as
its output.

### When an action doesn't find what it's looking for

Search actions (`find_row`, `find_column`, `find_headers_row`) return `status: "error"` when
nothing matches — that's a normal outcome, not a crash. The workflow keeps running; later steps
can check `if: "{{ steps.find_it.status == 'success' }}"` to react to it. A genuinely broken
step (bad parameters, an unsupported combination) raises instead and stops the run — nothing
is ever written back to the real files unless every step in the run succeeded.

### Workbook lifecycle

Excel Runner stages both read-only and read-write workflow sessions in its scratch directory.
Read-only copies are never committed; changed workbook copies are committed only after a
successful run. The read-only preflight is the exception: it opens declared workbook paths
directly to validate them before a run begins.

Workbooks open automatically on first reference — no `open` step required. On a fully
successful run, every workbook that was written to is saved automatically. Explicit `open`/
`save`/`close` steps exist for manual control (e.g. saving partway through) but are optional.

When a workbook declares `template: <logical name>`, its scratch copy always starts from that
template on every run, even when its own `file:` already exists. The successful result replaces
that file atomically. `create_if_missing: true` creates a blank workbook only when no template
is declared.

All work occurs in the workflow's scratch directory, not directly against declared workbook
paths. On success, changed scratch workbooks are closed, saved, and atomically committed to
their `file:` destinations; on an error, they remain only in the run artifacts. A template is
therefore suitable for repeatable output generation: an old output never becomes the next run's
hidden input.

### Save blockers

Excel Runner normally uses openpyxl, including for ordinary reads and writes. Before the first
write, it inspects the staged workbook for save blockers. The only empirically verified blocker
today is an outbound external-workbook link: openpyxl can damage that link metadata when it
saves, even when the workflow did not change the formula containing the link.

Read-only access remains on openpyxl. A blocker-bearing workbook that will be written is
automatically promoted to a locally spawned Excel/xlwings session before its first write and
stays there for the run. This requires a usable local Excel installation; the runner reports an
error rather than falling back to an unsafe save. Detected blockers and backend changes are
recorded in the audit log.

An external link to a workbook that is not modified by the run remains pointed at that workbook's
real path. It is not staged, rewired, or changed merely because another workbook links to it.
The linking workbook still uses Excel/xlwings for saving, because its outbound link is a save
blocker. Only an absolute link to another declared workbook that will also be modified is
temporarily rewired to that workbook's scratch copy and included in commit ordering.

Additional save blockers are added only when testing demonstrates that openpyxl cannot safely
round-trip that workbook feature.



## Changelog

- **2026-09-11**: Added `read_text_file` for read-only FAC/CSV/TSV/line-based input, preserving
  every field as text; regex replacement across sheets, ranges, and lookup-selected table cells;
  table discovery plus header-driven copy/update actions; and workbook-level defined-name targets
  for `write_cell` and `write_range`. `--check-existence` now validates defined names used by
  those write targets. These actions have unit and real-workbook workflow integration coverage.
- **2026-09-11**: Added shared read-only preflight validation. `--dry-run` validates and exits
  without staging or changing files; `--check-existence` runs the same preflight before a real
  workflow execution, so production runs fail before any workbook mutation.
- **2026-09-07**: Every `xlw_`/`com_` backend call is now wrapped in a structured error
  boundary (`backends._excel_operation`) — a live Excel/COM failure surfaces as
  `ActionExecutionError` with a readable message instead of a raw traceback; `ValueError`/
  `NotImplementedError` guards shared by both backend twins, plus `FileNotFoundError`/
  `TimeoutError`, still pass through unwrapped. `cli.main` also gained a catch-all so no
  unexpected exception can escape as a traceback. Fixed a real silent-corruption bug: opening
  an `.xlsm` for writing now sets `keep_vba` so saving no longer drops its VBA project;
  gated by file extension so a plain `.xlsx` doesn't get mislabelled macro-enabled. See
  [`demos/08_full_showcase/README.md`](demos/08_full_showcase/README.md) for the full-showcase
  demo's own docs (previously undocumented beyond a build-planning script).
- **2026-09-07**: Fixed a double-close bug in session teardown: an explicit `close` action
  step left `SessionManager` still tracking that session, so `close_all()`'s end-of-run
  cleanup tried to close it a second time — harmless against the file backend, but a fatal
  `pywintypes.com_error` against a live Excel (`xlw`) session, since the COM object was
  already disconnected. This surfaced as an `ExceptionGroup` and a non-zero exit code even
  though every real step had already succeeded. `SessionManager.forget_session()` is now
  called after a successful `close` action so `close_all()` never revisits it.

- **2026-09-03**: Added a third, opt-in validation tier — `--check-existence` (CLI) /
  `run_workflow(..., check_existence=True)` (library) opens every referenced workbook
  read-only via openpyxl, before any session/scratch machinery, and confirms every sheet and
  workbook-level defined name a step references by literal name actually exists (plain A1
  cell/range references like `"A1"` are deliberately never checked). Sheet existence is
  tracked step by step so a sheet added by an earlier `create_sheet` (or renamed/removed by
  `rename_sheet`/`delete_sheet`) counts correctly from that point on. Also added the `dump`
  control action (prints or writes any prior step's recorded output as JSON, for inspecting a
  workflow's internal per-step storage) and an always-on `working_dir/steps_dump.json` file,
  written at the end of every run, consolidating every step's output into one pretty-printed
  JSON object (a friendlier companion to the line-oriented `audit.jsonl`).
- **2026-08-26**: Added `recalculate` (live-Excel formula recalculation via xlwings/COM).
  `SessionManager` now supports bidirectional file<->xlw backend switching mid-run (closing
  and reopening a workbook's handle in place) instead of raising, and every xlw/com-capability
  session in a run shares one lazily-spawned Excel instance so cross-workbook links resolve
  correctly.
- **2026-08-21**: Crash/lock-safety hardening. `working_dir` (`excel_runner_runs/<yaml_stem>/`,
  under cwd by default or `--working-dir`) replaces the old random temp directory — a fixed,
  predictable location external tooling can construct itself from just the yaml's filename.
  Read-only sessions are now staged too (not opened directly against the real file), closing a
  hang-related file-lock exposure. Committing a workbook back to its real path is now
  rename-based with per-file rollback if a later workbook in the same run fails to commit.
  Nothing in `working_dir` is deleted automatically anymore. Added console logging (stdlib
  `logging`, `--logging-level` flag) alongside the existing structured audit log — the CLI no
  longer prints the `RunResult` as JSON to stdout either; since `working_dir` is a fixed path,
  an external caller can read `working_dir/audit.jsonl` and `working_dir/run.log` directly.
- **2026-08-20**: Added `create_sheet`/`rename_sheet`/`delete_sheet` actions (no prior way to
  add/rename/remove a worksheet). Also fixed a real design gap found while adding them: whether
  an action needs its workbook opened read-write was tracked in a hardcoded list disconnected
  from the action's own registration — now declared directly on the action via
  `@file_action(writes=True)`, so a new write action can't silently be left off it. Added a
  `demos/` folder: 7+ runnable YAML workflows exercising every action's syntax, for regression
  testing.
- **2026-08-20**: Fixed a real, intermittent (~50% of runs) hang in
  `OwnedInstanceRegistry.close_owned()` on Windows (`spawn()` now uses `add_book=True` so
  spawned instances register in `xw.apps`). Added CLI entrypoints (`cli.py`, `__main__.py`,
  `excel-runner` console script) to run a workflow YAML file and print its
  result as JSON — driven by the need to invoke a workflow from an external orchestration
  workflow as an external process. Full quality-gate suite (pytest, ruff, mypy --strict, pyright,
  radon) confirmed clean on Windows for the first time.

## Action reference

Every action needs `workbook: <logical name>` (from the `workbooks:` block), except `copy`,
which needs `source:`/`target:` instead. Fields are required unless marked optional.

### Basic

#### `open`

Confirms a workbook is open. Rarely needed explicitly — workbooks open automatically. No other
fields. This should almost never be needed. The `workbook` value names an entry in the
`workbooks:` section.

```yaml
- id: open_it
  action: open
  workbook: my_book
```

#### `save`

Saves the workbook now instead of waiting for the automatic end-of-run save. This can be useful
when testing a workflow stage, for example a `save` action followed by `stop` so you can review
the workbook at that point.

```yaml
- id: save_it
  action: save
  workbook: my_book
```

#### `close`

Closes the workbook, releasing its file handle.

```yaml
- id: close_it
  action: close
  workbook: my_book
```

#### `stop`

Halts the run right there — no workbook, no later step runs. Pairs with `if:` so you don't have
to repeat the same condition on every step downstream of a lookup that might fail:

```yaml
- id: guard
  action: stop
  reason: "region not found"    # optional — shows up in the audit log
  if: "{{ steps.find_it.status == 'error' }}"
```

Every step after a triggered `stop` gets `status: "stopped"` instead of running — distinct from
`skipped`, so you can tell "this step's own `if:` said don't run" apart from "the run ended
before we got here." Reaching `stop` isn't itself a failure: whether the run saves still depends
only on whether an *earlier* step returned `status: "error"` — "not found → stop" naturally
discards, but a deliberate early exit on success ("already done, nothing to do") still saves
whatever ran before it.

#### `dump`

Prints (or writes) the recorded output of prior steps as formatted JSON — for inspecting a
workflow's internal per-step storage while authoring/debugging a workflow. No `workbook:` field.

| Field | Required | Notes |
|---|---|---|
| `ids` | no | list of step ids to include; omit to dump every step recorded so far. An unknown/typo'd id is skipped with a logged warning, not an error |
| `to` | no | `"console"` (default, prints to stdout) or `"file"` |
| `path` | only with `to: file` | where to write the JSON; parent directories are created as needed |

```yaml
- id: show_progress
  action: dump
  ids: [get_total, find_it]
  to: console
```

Every run also always writes `working_dir/steps_dump.json` — every step's recorded output, in
one pretty-printed JSON object — regardless of whether a `dump` step is used.
The `dump` action's own step output is `{}`; the JSON it prints or writes is its side effect.

### Data
**Note:** Where an action explicitly supports a workbook-level defined name, it can be used in
place of an A1 cell or range reference. This is not a general rule for every range-like field;
`copy` still requires both source and target sheet names.

#### `copy`

Copies a range — or the source sheet's used range when `range` is omitted — from one workbook
into another. This uses Excel copy/paste through a shared live Excel session, so formulas and
formatting are preserved rather than copied as values only.

| Field | Required | Notes |
|---|---|---|
| `source.workbook`, `source.sheet` | yes | |
| `source.range` | no | omit to copy the whole sheet |
| `target.workbook`, `target.sheet`, `target.range` | yes | `target.range`'s top-left cell is where the copy starts |

```yaml
- id: copy_data
  action: copy
  source:
    workbook: historical
    sheet: "Reserving Data"
    range: "A1:AC50"
  target:
    workbook: my_book
    sheet: "Reserving Data"
    range: "A1"
```

#### `read_range`

Reads a cell or range, from one sheet or several. Output: `{{ steps.<id>.output.values }}` —
for a single sheet name, a single value for one cell or a 2D list of rows for a range (same
as before); for a list/`all`/`matching` sheet spec, a dict keyed by sheet name, one entry per
resolved sheet.

| Field | Required |
|---|---|
| `sheet`, `range` | yes |

`sheet` accepts four forms:

| Form | Meaning |
|---|---|
| `"North"` | A single sheet, by exact name. |
| `["North", "South"]` | An explicit list — multi-sheet capture. |
| `"all"` | Every sheet in the workbook. |
| `{ matching: "^A&H" }` | Every sheet whose name matches this regex (`re.search`, same convention as `find_row`/`find_headers_row`'s `patterns`). |

```yaml
- id: get_totals
  action: read_range
  workbook: my_book
  sheet: "Outputs"
  range: "A1:D50"
```

```yaml
- id: get_ah_status
  action: read_range
  workbook: my_book
  sheet: { matching: "^A&H" }
  range: "O6"
# .output.values is keyed by sheet name, e.g. {"A&H North": "Pass", "A&H South": "Fail"}
```

#### `read_metadata`

Reads workbook document properties, or a scattered list of specific cells.

| Field | Required | Notes |
|---|---|---|
| `target` | yes | `"properties"` or `"cells"` |
| `sheet`, `cells` | if `target: cells` | `cells` is a list of A1 references |

The action returns every non-empty standard document property available in the workbook. Common
properties include:

* `title`
* `subject`
* `creator`
* `keywords`
* `description`
* `lastModifiedBy`
* `created`
* `modified`
* `category`
* `contentStatus`
* `identifier`
* `language`
* `revision`
* `version`

```yaml
- id: get_props
  action: read_metadata
  workbook: my_book
  target: properties
```

```yaml
- id: get_specific_cells
  action: read_metadata
  workbook: my_book
  target: cells
  sheet: "Summary"
  cells: ["A1", "B3"]
```

Output for `properties`: property name → value (e.g. `.output.title`, `.output.creator`).
Output for `cells`: cell reference → value (e.g. `.output.A1`).

Set `formula: true` with `target: cells` to return formula text instead of a cached calculated
value. `target: textboxes` is not implemented and raises a clear runtime error.

#### `parse_date`

Parses a text value using Python `datetime.strptime` directives and returns a native Python
date. It has no `workbook:` field. Invalid input or calendar dates produce a structured action
error.

| Field | Required | Notes |
|---|---|---|
| `value`, `format` | yes | Example format: `"%Y-%m-%d"` |

```yaml
- id: parse_period_end
  action: parse_date
  value: "2026-06-30"
  format: "%Y-%m-%d"

- id: write_period_end
  action: write_cell
  workbook: my_book
  sheet: "Summary"
  cell: "B2"
  value: "{{ steps.parse_period_end.output.value }}"
```

#### `write_cell`

Writes one value to one cell. A value starting with `=` is stored as a formula. `cell` accepts
either A1 notation or a workbook-level defined name resolving to one cell. An A1 target
requires `sheet`; a defined name uses its own destination sheet and allows `sheet` to be
omitted or blank. A nonblank supplied sheet that differs from the name's destination is ignored
and logged as a warning.

| Field | Required |
|---|---|
| `cell`, `value` | yes |
| `sheet` | required for an A1 target; optional for a defined name |

```yaml
- id: set_status
  action: write_cell
  workbook: my_book
  sheet: "Summary"
  cell: "B2" # Or a one-cell workbook-level defined name.
  value: "Complete"

- id: set_formula
  action: write_cell
  workbook: my_book
  sheet: "Model"
  cell: "D10"
  value: "=SUM(D2:D9)"

- id: set_named_value
  action: write_cell
  workbook: my_book
  cell: "Inputs_Status"
  value: "Complete"
```

Note: openpyxl doesn't evaluate formulas — reading `D10` back gives `None`/stale data until
the workbook is recalculated (see `recalculate` below).

#### `write_range`

Writes a 2D block of values, anchored at the top-left cell of `range`. `range` accepts A1
notation or a one-area workbook-level defined name. An A1 target requires `sheet`; a defined
name uses its own destination sheet and allows `sheet` to be omitted or blank. A nonblank,
conflicting supplied sheet is ignored and logged as a warning.

| Field | Required |
|---|---|
| `range`, `values` | yes |
| `sheet` | required for an A1 target; optional for a defined name |

```yaml
- id: write_block
  action: write_range
  workbook: my_book
  sheet: "Summary"
  range: "B2"
  values:
    - [10, 20, 30]
    - [40, 50, 60]
```

#### `read_text_file`

Reads a text file as `{{ steps.<id>.output.values }}`, a 2D list of **strings** suitable for
`write_range`. It has no `workbook:` field and never changes the source file. `.csv` and `.fac`
default to comma-separated values; `.tsv` and `.txt` default to tab-separated values; other
extensions return one non-empty input line per one-cell row. Use `delimiter`, `quotechar`, or
`encoding` to override parsing. No type checking or conversion occurs, so values such as
`00123`, `202606`, and `1.50` remain text.

```yaml
- id: read_basis_fac
  action: read_text_file
  file: "./GLOBAL/BASIS_202606.fac"

- id: write_basis
  action: write_range
  workbook: my_book
  sheet: "BASIS"
  range: "B7"
  values: "{{ steps.read_basis_fac.output.values }}"
```

#### Table and replacement actions

`replace_text` replaces a regular expression in every populated cell of one sheet, a sheet list,
`"all"`, or `{ matching: "<regex>" }`. `replace_in_range` limits that operation to an A1 or
one-area defined-name range. Both return `{"replacements": <changed cell count>}`.

`read_table`, `copy_table_columns`, `update_table_cells`, and `replace_table_text` discover a
table from a literal top-left `header_cell`, such as `"B7"`: headers extend right to the first
blank, and data rows extend down to the first blank in the first table column. Header and lookup
matching in table-writing actions is case-insensitive; ambiguous names or lookup rows are errors.

```yaml
- id: copy_source_to_target
  action: copy_table_columns
  workbook: my_book
  sheet: "BASIS"
  header_cell: "B7"
  source_columns: ["SOURCE"]
  target_columns: ["TARGET"]

- id: set_table_value
  action: update_table_cells
  workbook: my_book
  sheet: "BASIS"
  header_cell: "B7"
  lookup_column: "BASIS_ITEM"
  lookup_rows: ["ACC_RATIO"]
  target_columns: ["TARGET"]
  value: "na"
```

#### `write_row`

Writes a row of values — either an explicit column mapping, or an ordered list starting at a
column.

```yaml
# explicit column mapping
- id: write_summary_row
  action: write_row
  workbook: my_book
  sheet: "Summary"
  row: 5
  values: { B: "North", C: 1200, D: "PASS" }

# positional — values in order, starting at start_column
- id: write_summary_row_positional
  action: write_row
  workbook: my_book
  sheet: "Summary"
  row: 5
  start_column: B
  values: ["North", 1200, "PASS"]
```

`start_column` is required when `values` is a list, ignored when it's a mapping.

### Structure

#### `insert_range`

Inserts a whole row or whole column, shifting existing content. Only whole-row (`"5:5"`) or
whole-column (`"C:C"`) references are supported — a partial range (`"C5:C10"`) returns a
structured error, not a crash.

| Field | Required | Notes |
|---|---|---|
| `at` | yes | e.g. `"C:C"` or `"5:5"` |
| `header` | no | `{row, text}` — only meaningful for a column insert |

```yaml
- id: insert_flag_column
  action: insert_range
  workbook: my_book
  sheet: "Summary"
  at: "C:C"
  header: { row: 1, text: "Flag" }
```

#### `set_column_width`

| Field | Required | Notes |
|---|---|---|
| `columns` | yes | e.g. `"B"` or `"A:C"` |
| `width` | yes | a number, or `"autofit"` |

```yaml
- id: widen_columns
  action: set_column_width
  workbook: my_book
  sheet: "Summary"
  columns: "A:C"
  width: autofit
```

#### `create_sheet`

Adds a new, empty worksheet.

| Field | Required | Notes |
|---|---|---|
| `name` | yes | name for the new sheet |
| `index` | no | 0-based position; appended at the end if omitted |

Returns a structured error (not a crash) if a sheet named `name` already exists.

```yaml
- id: add_data_sheet
  action: create_sheet
  workbook: my_book
  name: "Data"
```

#### `rename_sheet`

Renames an existing worksheet.

| Field | Required |
|---|---|
| `sheet`, `new_name` | yes |

```yaml
- id: rename_it
  action: rename_sheet
  workbook: my_book
  sheet: "Sheet"
  new_name: "Data"
```

#### `delete_sheet`

Removes a worksheet. Returns a structured error if `sheet` is the workbook's only remaining
sheet — a workbook can't have zero sheets.

| Field | Required |
|---|---|
| `sheet` | yes |

```yaml
- id: remove_scratch_sheet
  action: delete_sheet
  workbook: my_book
  sheet: "Scratch"
```

### Lookup

#### `find_headers_row`

Finds the row where every pattern (regex) matches some cell in that row.

```yaml
- id: find_headers
  action: find_headers_row
  workbook: my_book
  sheet: "Summary"
  search_range: "A1:J10"
  patterns: ["Region", "Total", "Status"]
```

Output: `.output.row` (the row number) and `.output.headers` (pattern → column letter).

#### `find_row`

Finds the row where a column equals a value.

```yaml
- id: find_north
  action: find_row
  workbook: my_book
  sheet: "Summary"
  column: "A"
  search_value: "North"
  header_row: 1        # optional — search starts after this row
```

Output: `.output.row`.

#### `find_column`

Finds one column by header pattern (regex).

```yaml
- id: find_status_col
  action: find_column
  workbook: my_book
  sheet: "Summary"
  header_row: 1
  pattern: "Status"
```

Output: `.output.column` (a letter).

#### `find_columns`

Finds several named columns in one call. A name whose pattern doesn't match anything is simply
absent from the output — not an error.

```yaml
- id: find_key_columns
  action: find_columns
  workbook: my_book
  sheet: "Summary"
  header_row: 1
  patterns: { region: "Region.*", total: "Total.*", status: "Status" }
```

Output: logical name → column letter (e.g. `.output.region`).

### Recalculation

#### `recalculate`

Forces Excel to recalculate formulas, then saves immediately. Needs a real, locally-spawned
Excel instance — the workbook's session switches to that backend automatically the first time
this (or another live-Excel) action needs it, closing/reopening its openpyxl handle in the
process. Every workbook a run opens this way shares one Excel instance, so cross-workbook
links resolve correctly.

| Field | Required | Notes |
|---|---|---|
| `scope` | no | `"sheet"`, `"workbook"` (default), or `"all"` (every workbook open in this run's shared Excel instance) |
| `mode` | no | `"normal"` (default), `"full"`, or `"full_rebuild"` — the latter two are always application-wide in Excel, so they require `scope: "all"` |
| `sheet` | no | only meaningful with `scope: "sheet"`; if omitted, the active sheet is used and `.output.warning` names which one |

```yaml
- id: recalc_my_book
  action: recalculate
  workbook: my_book
  scope: workbook
  mode: normal
```

Output: `.output.scope`, `.output.mode`, plus `.output.sheet`/`.output.warning` when `scope` is
`"sheet"`.


## Not yet available

Flagged clearly rather than silently missing:

- **Additional save blockers** — only outbound external-workbook links are currently verified.
  Other file types or workbook features need empirical safety testing before they are treated as
  blockers.
- **`copy` named-range sheet resolution** — `copy` requires source and target sheet names even
  when a named range is supplied. The action does not yet resolve the named range's sheet.
- **`read_metadata` cell properties** — formatting and layout details such as font, fill,
  number format, column width, and row height are not supported.
- **`write_table`, `aggregate`, `write_row's` dedicated by-header mode** — no action-specific
  contract has been designed. Whole-expression `{{ }}` templating can already compose prior
  step outputs into `write_range` or `write_row` values where suitable.
- **`read_links`, `write_links`** — reading/rewriting external workbook links. A real
  limitation in openpyxl (not just unbuilt), see `docs/PRD.md` §7.
- **`refresh_links`, `run_macro`, `export_pdf`** — all need a live Excel session
  (via xlwings), which these specific actions don't use yet.
- **`read_metadata` with `target: textboxes`** — same live-Excel limitation; raises a clear
  error if requested.
- **`update_summary_table`** — not designed yet.

## Running in Unify

**TBC:** the Unify deployment model, its command wrapper, working-directory ownership, and
operational credential handling have not yet been agreed. Until that decision is documented,
deploy the source bundle only in a locally controlled Windows environment with Excel installed;
keep input workbooks local for the first `--dry-run` and `--check-existence` validation.

## Development and contribution

When adding actions or modifying core code, read the specification and ensure GitHub Copilot is
also instructed to use it.

### GitHub Copilot and manual editing
Either in your prompt or instructions files ensure the following are done:
* Create a branch from the `develop` branch.
* Add unit tests for every part of the change.
* When using GitHub Copilot, instruct it to use test-driven development: write a failing test,
  implement the behavior, then rerun the test until it passes.
* Once unit tests pass, add sufficient integration tests for the new functionality or fix.
* Update all required documentation, including a README changelog entry.
* Run all required quality checks. Integration tests can take time because they operate through
  Excel.
* Commit with clear messages and submit a pull request explaining what changed, why, how, and
  the supporting evidence. Before a pull request, all tests must pass and every module must have
  at least 90% coverage.


```powershell
.venv\Scripts\python -m pip install -e ".[dev]"
.venv\Scripts\pytest tests\unit tests\integration
.venv\Scripts\ruff check .
.venv\Scripts\mypy --strict excel_runner tests vulture_whitelist.py
.venv\Scripts\radon cc --min C .
.venv\Scripts\vulture excel_runner vulture_whitelist.py --min-confidence 60
```

`docs/Progress_Tracker.md` tracks build status per component. `docs/Specification.md` §0
explains the sourcing policy for a prior, superseded tool this project doesn't reuse code or
structure from.

> **Maintenance requirement:** Every behavior change must update this `README.md`, the
> [`excel-runner-yaml` skill](.github/skills/excel-runner-yaml/SKILL.md), and
> [`docs/test_summary.md`](docs/test_summary.md), as well as the progress tracker and
> specification where they describe the changed behavior. New behavior is developed test-first
> with unit tests and real-workbook integration tests.
