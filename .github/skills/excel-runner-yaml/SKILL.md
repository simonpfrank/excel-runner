---
name: excel-runner-yaml
description: 'Use when writing, editing, reviewing, or debugging excel_runner workflow YAML files (demos/*.yaml or any file with `workbooks:`/`steps:`/`action:` keys). Covers exact syntax for every action, required/optional fields, output shapes, `{{ }}` templating rules, `if:`/`stop` control flow, and common authoring mistakes (wrong param names, wrong output keys, YAML quoting pitfalls, copy''s nested shape). Use this instead of guessing from the README when constructing a new workflow file or fixing a validation error.'
---

# excel_runner Workflow YAML Syntax

This is the authoritative, exhaustive syntax reference for excel_runner workflow files. It
exists because the README is a narrative introduction, not a spec — agents writing YAML from
memory/pattern-matching have produced files that fail `validate_static` (Spec sec 5.4) or
silently do the wrong thing. When writing or editing a workflow YAML file, follow this
document field-by-field rather than inferring from partial examples.

The ground truth for every action's real signature is
[excel_runner/actions.py](../../../excel_runner/actions.py) (each function's docstring —
`Args:`/`Returns:`/`Raises:`) and the dataclasses in
[excel_runner/core.py](../../../excel_runner/core.py) (`WorkbookRef`, `Step`, `Workflow`). If
this document and the source ever disagree, re-read `actions.py` — it's authoritative.

## 1. Top-level file shape

Exactly three top-level keys. No others are recognized.

```yaml
env:            # optional dict — plain values only, no {{ }} referencing steps
  input_folder: "./input"

workbooks:      # required (at least one, unless the workflow only uses `stop`) — logical name -> ref
  manip:
    file: "./manip.xlsx"

steps:          # required — ordered list, executed top to bottom
  - id: get_total
    action: read_range
    workbook: manip
    sheet: "Summary"
    range: "B2"
```

There is no `name:`, `description:`, `version:`, or `on_error:` top-level key. Do not invent
one — unrecognized top-level keys are silently ignored by the loader (they are not validated),
which means a typo here fails silently instead of erroring.

## 2. `env:` block

- Plain dict, any keys, any JSON-scalar values (str/int/float/bool/null) or nested
  dicts/lists.
- Values may use `{{ env.OTHER_KEY }}` templating against **other `env:` keys already
  defined**, but **cannot** reference `steps.*` — step outputs don't exist yet when `env:` is
  resolved (it's resolved once, at load time, before any workbook opens).
- `env_overrides` passed to `run_workflow(path, env_overrides=...)` are merged over (take
  precedence over) the file's own `env:` block, same keys.

```yaml
env:
  output_folder: "./output"
  run_tag: "batch-{{ env.output_folder }}"   # OK — references another env key
```

## 3. `workbooks:` block

Dict of **logical name → ref**. The logical name is arbitrary (your choice) and is what steps
use in their `workbook:` field — it is NOT the sheet name and NOT the file path.

Each ref supports exactly these fields:

| Field | Required | Type | Notes |
|---|---|---|---|
| `file` | yes | str | Path to the `.xlsx` file. May use `{{ env.* }}` templating. |
| `create_if_missing` | no | bool | Default `false`. Creates a blank workbook at `file` on first reference if it doesn't exist. |
| `template` | no | str | Logical name of **another** workbook in this same `workbooks:` block, whose content is copied when creating (requires `create_if_missing: true`). |

```yaml
workbooks:
  historical:
    file: "{{ env.input_folder }}/historical.xlsx"
  results:
    file: "{{ env.output_folder }}/results.xlsx"
    create_if_missing: true
    template: historical
```

Common mistakes:
- Do **not** nest `sheet:` or `range:` under `workbooks:` — sheets/ranges are per-step fields,
  not per-workbook.
- `template:` must name another key in this same `workbooks:` dict, not a file path.
- A workbook referenced by a step but missing from `workbooks:` fails validation with
  `Workbook "X" is not declared in the workbooks: registry.` — every `workbook:` value used
  anywhere in `steps:` must have a matching entry here.

## 4. `steps:` block

A list (not a dict) of step objects, executed strictly top to bottom. Each step:

| Field | Required | Notes |
|---|---|---|
| `id` | yes | Unique across the whole file. Used for later `{{ steps.<id>.output }}` / `{{ steps.<id>.status }}` references. Must be a valid Jinja identifier: letters, digits, underscore, not starting with a digit. |
| `action` | yes | One of the action names in section 6 below. Unknown action names fail validation (with a "did you mean" suggestion if close). |
| `workbook` | yes, except `stop` and `copy` | The logical name from `workbooks:`. `copy` uses nested `source.workbook`/`target.workbook` instead (see 6.2). `stop` has no workbook at all. |
| `if` | no | A `{{ }}` (or bare) Jinja boolean expression. If false, the step is **skipped** — it does not run, and its `status` becomes `"skipped"`, not `"error"`. |
| ...action-specific params | varies | See section 6 — every action has its own exact field list. |

**Any field not recognized by the action's own signature is a validation error** —
`"<field>" is not a recognized parameter for this action.` Do not add extra fields
"for documentation" (e.g. a `description:` or `note:` field on a step) — they will be rejected.

**Every required field must be present** — `missing required field "<field>"`.

**Step ids can only reference earlier steps.** `{{ steps.later_id.output }}` where `later_id`
is defined at or after the current step is a validation error (forward/self reference). Order
steps so anything referenced already ran.

## 5. Templating rules (`{{ }}`)

- Available context: `env` (always) and `steps` (accumulated outputs of steps that already
  ran, keyed by step `id`). There is no `this`/`self`/implicit current-step context.
- Every step's outcome is available as `{{ steps.<id>.output }}` (a dict — see each action's
  output shape in section 6) and `{{ steps.<id>.status }}` (one of `"success"`, `"error"`,
  `"skipped"`, `"stopped"`).
- **Whole-expression native typing**: if a YAML value is *entirely* one `{{ ... }}` block with
  nothing else around it, it resolves to the real Python value (int/float/list/dict/bool),
  not a string:
  ```yaml
  value: "{{ steps.get_total.output.values }}"   # -> real number, not "123"
  ```
  If the expression is embedded inside a longer string, the whole thing renders as text:
  ```yaml
  file: "{{ env.output_folder }}/results.xlsx"    # -> a string, always
  ```
- Referencing something undefined (a typo'd step id, a key that doesn't exist in that step's
  output, an undefined `env.*` key) raises a validation/runtime error immediately — this
  engine uses Jinja2 `StrictUndefined`, there is no silent-empty-string fallback like default
  Jinja/Ansible. Double-check output key names against section 6 (e.g. `find_row`'s output key
  is `row`, not `value` or `result`).
- A dict key that is itself a real dict method name (`values`, `keys`, `items`) still works as
  a plain key lookup here (e.g. `steps.get_totals.output.values` reaches the output dict's
  `"values"` key, not Python's `dict.values` method) — this is a deliberate engine fix, not
  something to work around.
- A field whose value is *entirely* one `{{ ... }}` expression is exempt from static
  (pre-run) type checking — its real shape can't be known until execution, since it may
  depend on a previous step's output. This means a full Jinja expression is a legitimate way
  to build a value no dedicated action exists for yet, e.g. a headered table from a previous
  step's dict output, or a count of rows matching a value, both without any custom action:
  ```yaml
  # builds [["Sheet", "State"], ["North", "Pass"], ...] from a prior read_range step's
  # {"values": {"North": "Pass", ...}} output — no "build a table" action needed
  values: >-
    {{ [["Sheet", "State"]]
       + (steps.capture_status.output.values.items() | list | map('list') | list) }}

  # counts how many [sheet, state] rows have state == "Pass" — no "count"/"aggregate"
  # action needed; `last` is a builtin Jinja filter (last item of each row), not custom
  value: >-
    {{ steps.read_table.output.values | map('last') | select('equalto', 'Pass') | list | length }}
  ```
- `if:` accepts the same expression either with or without the `{{ }}` wrapper — both of these
  are valid:
  ```yaml
  if: "{{ steps.find_it.status == 'success' }}"
  if: "steps.find_it.status == 'success'"
  ```
  Prefer the `{{ }}` form for consistency with every other templated field.

## 6. YAML quoting pitfalls specific to this engine

- This project's YAML loader uses **YAML 1.2 boolean semantics**, not PyYAML's default 1.1:
  only `true`/`True`/`TRUE`/`false`/`False`/`FALSE` become booleans. **`yes`, `no`, `on`, `off`
  stay plain strings.** So `search_value: no` is the string `"no"`, not `False` — safe to use
  unquoted for lookups against literal "Yes"/"No" cell values, but always quote it anyway for
  clarity: `search_value: "No"`.
- A1-style ranges that start with a bare number or contain a colon **must be quoted** — YAML
  parses `5:5` or `A1:D10` as plain scalars fine in most cases, but any range/cell reference
  should still be quoted defensively (`"5:5"`, `"C:C"`, `"A1:D10"`) since an unquoted leading
  digit or trailing colon can be misparsed depending on context (e.g. inside a flow mapping).
- Formula values (starting with `=`) must be quoted — `value: "=SUM(D2:D9)"` — otherwise some
  YAML parsers/tools misinterpret a leading `=`.
- Always quote any string value containing `{{ }}` templating as a plain YAML string (as every
  example in this doc does) — unquoted `{{ }}` at the start of a YAML scalar is parsed as a
  flow mapping by plain YAML and will error.

## 7. Full action catalog

Every action below lists its exact required/optional fields and exact output keys. `workbook:`
is required on all of these except `stop` (no workbook at all) and `copy` (nested `source`/
`target`, see 7.2). Output is always a keyed dict (never a bare value) — reference sub-fields
as `{{ steps.<id>.output.<key> }}`.

### 7.1 Basic

**`open`** — confirms a workbook is open. Rarely needed (workbooks auto-open on first
reference). Fields: `workbook` only. Output: `{}`.

**`save`** — saves now instead of waiting for the automatic end-of-run save. Fields: `workbook`
only. Output: `{}`.

**`close`** — closes the workbook, releasing its handle. Fields: `workbook` only. Output: `{}`.

**`stop`** — halts the run; every later step gets `status: "stopped"`. **No `workbook:` field —
do not add one, it will be rejected.**
| Field | Required |
|---|---|
| `reason` | no (str, shown in audit log) |
```yaml
- id: guard
  action: stop
  reason: "region not found"
  if: "{{ steps.find_it.status == 'error' }}"
```
Output: `{"reason": ...}` if `reason` given, else `{}`.

**`stop` is not an abort/discard.** Reaching `stop` is not itself a failure — the overall run
`status` is still `"success"` (and every workbook written to so far still gets saved/committed
to its real path) as long as no *earlier* step returned `status: "error"`. Only an earlier
step's `status: "error"` prevents the commit. Every open workbook session is closed safely in
a `finally` block regardless of how the run ends (normal completion, `stop`, or an earlier
error) — closing is unconditional and independent of whether anything was saved.

### 7.2 `copy` — the one action with a different YAML shape

`copy` takes **no top-level `workbook:` field**. Instead it takes `source:` and `target:`
nested dicts. This is the single most common mistake — do not write `workbook:` on a `copy`
step.

| Field | Required | Notes |
|---|---|---|
| `source.workbook` | yes | logical name |
| `source.sheet` | yes | |
| `source.range` | no | omit to copy the entire sheet |
| `target.workbook` | yes | logical name |
| `target.sheet` | yes | |
| `target.range` | yes | only the **top-left cell** of this is used as the paste anchor |

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
Output: `{}`.

**Omitting `source.range` copies the whole sheet** — but specifically its **used range**
(openpyxl's `iter_rows()` over the sheet's tracked dimension), not literally every cell up to
the spreadsheet limit, and **values only**: no formatting, no merged cells, no column widths,
and formulas are copied as whatever `.value` holds (a formula string if the source was opened
without `data_only`), never recalculated.

### 7.3 Data

**`read_range`**
| Field | Required |
|---|---|
| `sheet`, `range` | yes |
`sheet` accepts a single name (`"North"`), an explicit list (`["North", "South"]`) for
multi-sheet capture, `"all"` (every sheet in the workbook), or `{ matching: "<regex>" }`
(every sheet whose name matches, `re.search`-style — same convention as
`find_row`/`find_headers_row`'s `patterns`).
Output: `{"values": ...}` — for a single sheet name, a single scalar for a single-cell `range`
or a 2D list of rows for a multi-cell range (unchanged); for a list/`all`/`matching` sheet
spec, a dict keyed by sheet name instead. Reference as `{{ steps.<id>.output.values }}`.

**`read_metadata`**
| Field | Required | Notes |
|---|---|---|
| `target` | yes | `"properties"` or `"cells"` (exactly these two strings — `"textboxes"` is not implemented and raises a clear error) |
| `sheet` | required if `target: cells` | omit if `target: properties` |
| `cells` | required if `target: cells` | list of A1 refs, e.g. `["A1", "B3"]` |
```yaml
- id: get_props
  action: read_metadata
  workbook: manip
  target: properties
```
Output for `target: properties`: property name → value directly at the top of `output`
(e.g. `.output.title`, `.output.creator`) — not nested under a `values` key.
Output for `target: cells`: cell reference → value directly at the top of `output`
(e.g. `.output.A1`).

**`write_cell`**
| Field | Required | Notes |
|---|---|---|
| `sheet`, `cell`, `value` | yes | a `value` string starting with `=` is stored as a formula, not evaluated (openpyxl doesn't recalculate) |
Output: `{}`.

**`write_range`**
| Field | Required | Notes |
|---|---|---|
| `sheet`, `range`, `values` | yes | `values` is always a 2D list (list of row-lists), even for one row: `[[1, 2, 3]]`. `range` only needs its top-left cell — the block is written starting there. |
Output: `{}`.

**`write_row`** — two mutually exclusive modes, both under the same `values` field:
| Field | Required | Notes |
|---|---|---|
| `sheet`, `row` | yes | |
| `values` | yes | **either** a dict `{column: value}` **or** a plain list of values written positionally |
| `start_column` | required only if `values` is a list | ignored/omitted if `values` is a dict |
```yaml
# dict mode — explicit column mapping, no start_column
- id: write_summary_row
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  values: { B: "North", C: 1200, D: "PASS" }

# list mode — positional, REQUIRES start_column
- id: write_summary_row_positional
  action: write_row
  workbook: manip
  sheet: "Summary"
  row: 5
  start_column: B
  values: ["North", 1200, "PASS"]
```
Output: `{}`. **Not implemented**: a "by header name" mode referencing another step's headers
directly — do not invent a `values_by_header:`/`headers_from:` field, it doesn't exist yet.

### 7.4 Structure

**`insert_range`** — whole row/column only. A partial range like `"C5:C10"` returns a
structured `status: "error"`, not a crash — do not rely on partial-range insert working.
| Field | Required | Notes |
|---|---|---|
| `sheet`, `at` | yes | `at` must be a whole-column ref (`"C:C"`) or whole-row ref (`"5:5"`) |
| `direction` | no | `"rows"` or `"columns"` — unused in practice since `at`'s shape is already unambiguous; safe to omit |
| `header` | no | `{row: int, text: str}` — only meaningful for a column insert |
Output: `{}`.

**`set_column_width`**
| Field | Required | Notes |
|---|---|---|
| `sheet`, `columns` | yes | `columns` is a letter (`"B"`) or range (`"A:C"`) |
| `width` | yes | a number, **or the literal string `"autofit"`** |
Output: `{}`.

**`create_sheet`**
| Field | Required | Notes |
|---|---|---|
| `name` | yes | new sheet name |
| `index` | no | 0-based insert position; appended at end if omitted |
Returns `status: "error"` (not a crash) if `name` already exists. Output: `{}`.

**`rename_sheet`**
| Field | Required |
|---|---|
| `sheet`, `new_name` | yes |
Output: `{}`.

**`delete_sheet`**
| Field | Required | Notes |
|---|---|---|
| `sheet` | yes | returns `status: "error"` if this is the workbook's only remaining sheet |
Output: `{}`.

### 7.5 Lookup — every one of these returns `status: "error"` (not a crash) when nothing is found

**`find_headers_row`**
| Field | Required |
|---|---|
| `sheet`, `search_range`, `patterns` | yes |
`patterns` is a **list** of regex strings — every one must match some cell in a row for that
row to count.
Output on success: `{"row": int, "headers": {pattern: column_letter}}`.

**`find_row`**
| Field | Required | Notes |
|---|---|---|
| `sheet`, `column`, `search_value` | yes | `column` is a single letter, e.g. `"A"` |
| `header_row` | no | if given, search starts on the row after it |
Output on success: `{"row": int}`. (Not `"value"`, not `"result"`.)

**`find_column`**
| Field | Required |
|---|---|
| `sheet`, `header_row`, `pattern` | yes |
Output on success: `{"column": str}` — a single letter.

**`find_columns`** — plural, multiple named lookups in one call.
| Field | Required |
|---|---|
| `sheet`, `header_row`, `patterns` | yes |
`patterns` here is a **dict** `{logical_name: regex}` (different shape from
`find_headers_row`'s list!). A name whose pattern doesn't match anything is simply **absent**
from the output — that's not an error at this level.
Output: `{logical_name: column_letter, ...}` for every pattern that matched — reference as
`{{ steps.<id>.output.<logical_name> }}`.

**`recalculate`** — needs a live Excel session (xlwings/COM), unlike every other action above.
The workbook's session switches to that backend automatically the first time it's needed
(closing/reopening the handle in place) — no separate `open`-on-a-different-backend step is
needed or possible.
| Field | Required | Notes |
|---|---|---|
| `scope` | no | `"sheet"`, `"workbook"` (default), or `"all"` (every workbook open in this run's shared Excel instance) |
| `mode` | no | `"normal"` (default), `"full"`, or `"full_rebuild"` |
| `sheet` | no | only meaningful with `scope: "sheet"`; if omitted, the active sheet is used and a warning naming it is added to `output.warning` |

`mode: "full"`/`"full_rebuild"` are always application-wide in Excel — there is no per-sheet or
per-workbook equivalent — so they **require** `scope: "all"`; combining either with
`scope: "sheet"`/`"workbook"` is a validation error. Giving `sheet` with any `scope` other
than `"sheet"` is also an error (ambiguous).
```yaml
- id: recalc_manip
  action: recalculate
  workbook: manip
  scope: workbook
  mode: normal
```
Output: `{"scope": ..., "mode": ..., "sheet": ... (only if scope: sheet), "warning": ... (only if sheet fell back to the active sheet)}`.

## 8. Actions that do NOT exist yet — do not write steps for these

Do not generate steps using these names; they will fail `unknown action` validation:
`write_table`, `aggregate`, `read_links`, `write_links`, `refresh_links`,
`run_macro`, `export_pdf`, `update_summary_table`. Also `read_metadata` with
`target: textboxes` is rejected at runtime (not a static validation error, but always fails).

## 9. Self-check before finalizing a workflow file

1. Every `workbook:` value used in `steps:` has a matching key in `workbooks:`.
2. Every step `id` is unique and valid as a Jinja identifier.
3. Every `{{ steps.X... }}` reference points to a step `id` defined **earlier** in the file.
4. `copy` steps use `source:`/`target:`, never a top-level `workbook:`.
5. `stop` steps have no `workbook:` field.
6. Every field name on a step matches the exact parameter names in section 7 above — no
   invented fields (`description`, `label`, `comment`, etc. on a step).
7. `find_*` action output keys match section 7.5 exactly (`row`, `column`, `headers`, or named
   keys for `find_columns`) — a wrong key name only surfaces as a runtime `StrictUndefined`
   error, not a static validation error.
8. `write_row`'s `values` shape (dict vs list) matches whether `start_column` is present.
9. Any A1 range, cell ref, or formula string is quoted.
10. `env:` values never reference `steps.*`.
