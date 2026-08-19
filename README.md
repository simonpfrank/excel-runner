# excel_runner

Declarative, YAML-driven Excel automation. Describe a workflow — open workbooks, read and
write ranges, look things up, chain steps together — as data, not a script.

Design notes live in [`docs/PRD.md`](docs/PRD.md) and [`docs/Specification.md`](docs/Specification.md).

## Status

v1 file-backend engine is built and tested: loading, templating, validation, session/scratch
management, 16 actions, and the orchestration loop (`run_workflow`). There is **no CLI yet** —
this is a Python library, used from a script. Nothing here talks to a live Excel process (COM)
yet — everything runs against files directly via openpyxl, which is also why the action list
below skips macros, recalculation, and a few other Excel-specific things (see
[Not yet available](#not-yet-available)).

## Install and run

```bash
git clone <this repo>
cd excel-runner
uv venv .venv --python 3.12
source .venv/bin/activate
uv pip install -e .
```

Run a workflow from a script:

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

## Quick start

A minimal workflow: read one cell, write it somewhere else.

```yaml
workbooks:
  manip:
    file: "./manip.xlsx"

steps:
  - id: get_total
    action: read_range
    workbook: manip
    sheet: "Summary"
    range: "B2"

  - id: write_total
    action: write_cell
    workbook: manip
    sheet: "Summary"
    cell: "D2"
    value: "{{ steps.get_total.output.values }}"
```

The workbook is saved automatically at the end, if every step succeeded — no explicit `save`
step needed (see [Workbook lifecycle](#workbook-lifecycle)).

## Workflow YAML

A workflow file has three top-level blocks:

```yaml
env:                              # optional — plain values, referenced as {{ env.NAME }}
  input_folder: "./input"
  output_folder: "./output"

workbooks:                        # every workbook the workflow touches, by logical name
  historical:
    file: "{{ env.input_folder }}/historical.xlsx"
  results:
    file: "{{ env.output_folder }}/results.xlsx"
    create_if_missing: true       # create a blank workbook if the file doesn't exist
    template: historical          # optional — copy this workbook's content instead of blank

steps:                            # run in order
  - id: some_step                 # unique, referenced by later steps
    action: read_range            # one of the actions below
    workbook: historical          # logical name from workbooks: above
    sheet: "Summary"
    range: "A1:D10"
```

### Referencing another step's output

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

Workbooks open automatically on first reference — no `open` step required. On a fully
successful run, every workbook that was written to is saved automatically. Explicit `open`/
`save`/`close` steps exist for manual control (e.g. saving partway through) but are optional.

## Action reference

Every action needs `workbook: <logical name>` (from the `workbooks:` block), except `copy`,
which needs `source:`/`target:` instead. Fields are required unless marked optional.

### Basic

#### `open`

Confirms a workbook is open. Rarely needed explicitly — workbooks open automatically. No other
fields.

```yaml
- id: open_it
  action: open
  workbook: manip
```

#### `save`

Saves the workbook now, instead of waiting for the automatic end-of-run save.

```yaml
- id: save_it
  action: save
  workbook: manip
```

#### `close`

Closes the workbook, releasing its file handle.

```yaml
- id: close_it
  action: close
  workbook: manip
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

### Data

#### `copy`

Copies a range — or the whole sheet, if `range` is omitted — from one workbook into another.

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
    workbook: manip
    sheet: "Reserving Data"
    range: "A1"
```

#### `read_range`

Reads a cell or range. Output: `{{ steps.<id>.output.values }}` — a single value for one cell,
a 2D list of rows for a range.

| Field | Required |
|---|---|
| `sheet`, `range` | yes |

```yaml
- id: get_totals
  action: read_range
  workbook: manip
  sheet: "Outputs"
  range: "A1:D50"
```

#### `read_metadata`

Reads workbook document properties, or a scattered list of specific cells.

| Field | Required | Notes |
|---|---|---|
| `target` | yes | `"properties"` or `"cells"` |
| `sheet`, `cells` | if `target: cells` | `cells` is a list of A1 references |

```yaml
- id: get_props
  action: read_metadata
  workbook: manip
  target: properties
```

```yaml
- id: get_specific_cells
  action: read_metadata
  workbook: manip
  target: cells
  sheet: "Summary"
  cells: ["A1", "B3"]
```

Output for `properties`: property name → value (e.g. `.output.title`, `.output.creator`).
Output for `cells`: cell reference → value (e.g. `.output.A1`).

#### `write_cell`

Writes one value to one cell. A value starting with `=` is stored as a formula.

| Field | Required |
|---|---|
| `sheet`, `cell`, `value` | yes |

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

Note: openpyxl doesn't evaluate formulas — reading `D10` back gives `None`/stale data until
the file is recalculated in real Excel (recalculation isn't built yet, see below).

#### `write_range`

Writes a 2D block of values, anchored at the top-left cell of `range`.

| Field | Required |
|---|---|
| `sheet`, `range`, `values` | yes |

```yaml
- id: write_block
  action: write_range
  workbook: manip
  sheet: "Summary"
  range: "B2"
  values:
    - [10, 20, 30]
    - [40, 50, 60]
```

#### `write_row`

Writes a row of values — either an explicit column mapping, or an ordered list starting at a
column.

```yaml
# explicit column mapping
- id: write_summary_row
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  values: { B: "North", C: 1200, D: "PASS" }

# positional — values in order, starting at start_column
- id: write_summary_row_positional
  action: write_row
  workbook: manip
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
  workbook: manip
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
  workbook: manip
  sheet: "Summary"
  columns: "A:C"
  width: autofit
```

### Lookup

#### `find_headers_row`

Finds the row where every pattern (regex) matches some cell in that row.

```yaml
- id: find_headers
  action: find_headers_row
  workbook: manip
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
  workbook: manip
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
  workbook: manip
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
  workbook: manip
  sheet: "Summary"
  header_row: 1
  patterns: { region: "Region.*", total: "Total.*", status: "Status" }
```

Output: logical name → column letter (e.g. `.output.region`).

## Not yet available

Flagged clearly rather than silently missing:

- **`write_table`, `aggregate`, `write_row`'s by-header mode** — each needs to reference another
  step's output by its step id directly (not through `{{ }}` templating), which isn't wired up
  yet.
- **`read_links`, `write_links`** — reading/rewriting external workbook links. A real
  limitation in openpyxl (not just unbuilt), see `docs/PRD.md` §7.
- **`refresh_links`, `recalculate`, `run_macro`, `export_pdf`** — all need a live Excel (COM)
  session, which doesn't exist yet. Nothing in this engine spawns Excel.
- **`read_metadata` with `target: textboxes`** — same COM limitation; raises a clear error if
  requested.
- **`update_summary_table`** — not designed yet.
- A CLI. `run_workflow()` is the only entry point today.

## Using it as a library

```python
from excel_runner import run_workflow, list_actions, RunResult, StepResult

result: RunResult = run_workflow("workflow.yaml")

for spec in list_actions():
    print(spec.name, "-", spec.description)
```

`list_actions()` returns every built action's name, description, capability, and parameter
schema — useful for building a tool wrapper (CLI, MCP server, agent framework) on top without
duplicating the action catalog.

## Development

```bash
uv pip install -e ".[dev]"    # adds pytest, ruff, mypy, radon, vulture
source .venv/bin/activate
python -m pytest tests/unit/ tests/integration/    # 248 tests, 100% branch coverage
ruff check .
mypy excel_runner tests
radon cc excel_runner --min C
vulture excel_runner vulture_whitelist.py --min-confidence 60
```

`docs/Progress_Tracker.md` tracks build status per component. `docs/Specification.md` §0
explains the sourcing policy for a prior, superseded tool this project doesn't reuse code or
structure from.
