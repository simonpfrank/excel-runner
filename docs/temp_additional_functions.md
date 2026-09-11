# Additional Functions Specification

**Status:** complete. Implemented with unit and real-workbook integration coverage.

## 1. Purpose and scope

Add a dependency-free way to read `.fac`, CSV, TSV, and simple line-based text files into a
workflow, then write their values to Excel using the existing `write_range` action.

Also extend the existing `write_cell` and `write_range` actions so their reference fields
accept workbook-level defined names as well as literal A1 references. Do **not** add a separate
`write_named_range` action: it would duplicate the existing write actions.

This scope also adds broad sheet replacement, range-limited replacement, table discovery,
header-driven copying of data columns, lookup-driven table cell updates, and text replacement
within matched table cells. These actions match the existing table-creator script's behavior.
Dynamic source-table copying remains separate work.

## 2. `read_text_file`

### 2.1 Action contract

`read_text_file` is a read-only control action. It does not have a `workbook:` field and never
creates, moves, renames, or changes its source file.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `file` | yes | string | Path to the source text file. May use `{{ env.* }}` or earlier step output templating. |
| `delimiter` | no | string | Field delimiter. Overrides extension defaults. |
| `quotechar` | no | string | Single-character CSV quote character. Default: `"`. |
| `encoding` | no | string | Text encoding used to read the file. Default: `utf-8`. |

Output shape:

```yaml
output:
  values: [["Header A", "Header B"], ["one", "two"]]
```

Every returned field is a Python string. There is no type checking, inference, or conversion.
For example, `00123`, `202606`, and `1.50` all remain text.

### 2.2 Delimiter rules

When `delimiter` is omitted:

| File extension | Default behavior |
|---|---|
| `.csv` | comma-delimited |
| `.fac` | comma-delimited |
| `.tsv` | tab-delimited |
| `.txt` | tab-delimited |
| any other extension | one column per non-empty input line |

For a line-based file without a delimiter, this input:

```text
first line
second line
```

returns:

```yaml
values: [["first line"], ["second line"]]
```

An explicit `delimiter` always takes precedence over the extension. In YAML, use a double-quoted
tab escape for a tab delimiter: `delimiter: "\t"`.

### 2.3 Parsing and validation

- Use Python's standard-library `csv` module for delimited formats; do not add Pandas.
- Open delimited files with `newline=""` so `csv` handles `\n`, `\r\n`, and `\r` correctly.
- Use the supplied `quotechar`; reject a `quotechar` that is not exactly one character.
- Reject an explicitly supplied empty `delimiter` or delimiter longer than one character.
- Missing/unreadable files, invalid encoding names, decoding failures, malformed CSV quoting, and
  invalid options must raise a clear `ActionExecutionError` identifying the file and problem.
- Preserve blank fields within records. Ignore fully blank lines, matching the existing `.fac`
  reader's behavior.
- Do not use extension guessing beyond the defaults above, and do not use `csv.Sniffer`.

### 2.4 YAML example

```yaml
env:
  tables_path: "./input_tables"

workbooks:
  table_creator:
    file: "./TableCreator.xlsm"

steps:
  - id: read_basis_fac
    action: read_text_file
    file: "{{ env.tables_path }}/GLOBAL/BASIS_202606.fac"

  - id: write_basis
    action: write_range
    workbook: table_creator
    sheet: "BASIS"
    range: "B7"
    values: "{{ steps.read_basis_fac.output.values }}"
```

## 3. Defined-name support for existing writes

### 3.1 Contract

Extend these existing actions:

| Action | Existing reference field | New accepted reference forms |
|---|---|---|
| `write_cell` | `cell` | A1 cell or workbook-level defined name resolving to one contiguous area |
| `write_range` | `range` | A1 cell/range or workbook-level defined name resolving to one contiguous area |

The current `sheet` field remains required for schema consistency. When the reference is a
defined name, its destination sheet takes precedence and `sheet` is ignored, exactly as it is
for `read_range`.

Reuse the existing `resolve_range` and `xlw_resolve_range` logic. Do not create another defined
name resolver.

### 3.2 Semantics

- `write_cell`: A defined name must resolve to exactly one cell. Reject a multi-cell defined
  name with an actionable `ActionExecutionError`.
- `write_range`: A defined name may resolve to one contiguous cell/range. Use its top-left cell
  as the write anchor, matching the action's current A1-range behavior.
- Both backends must behave identically: openpyxl/file and xlwings/live Excel.
- Names resolving to multiple separate areas are rejected, matching read behavior.
- A literal invalid A1 reference and an unknown defined name must produce the same clear
  validation/execution error style already used by read actions.
- Extend `--check-existence` so `write_cell.cell` and `write_range.range` are validated as
  defined names whenever they are not literal A1 references.

### 3.3 YAML examples

```yaml
- id: set_closing_date
  action: write_cell
  workbook: table_creator
  sheet: "BASIS"
  cell: "Closing_Date"
  value: 202606

- id: import_basis
  action: write_range
  workbook: table_creator
  sheet: "BASIS"
  range: "Basis_Table_Anchor"
  values: "{{ steps.read_basis_fac.output.values }}"
```

## 4. `replace_text`

### 4.1 Action contract

`replace_text` is a writing file action that performs a blunt-force text replacement in every
cell of one or more selected sheets. `sheet` accepts the same forms as `read_range`: one exact
sheet name, an explicit list of names, `"all"`, or `{ matching: "<regex>" }`.

Reuse the existing `resolve_sheet_names` helper and its semantics; do not add a separate
multi-sheet resolution path for this action. Exact/list selections retain their supplied order,
while `"all"` and `{ matching: ... }` resolve in workbook sheet order.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string, list, or matching dict | Sheet or sheets in which to replace text. |
| `pattern` | yes | string | Regular expression to find in cell text. |
| `replacement` | yes | string | Text substituted for every regex match. |

The action visits every populated cell in each resolved sheet. It converts a selected cell value
to text only when it contains a match, then writes the replaced text back. Cells with no match
remain unchanged. Formulas, formatting, comments, validation, and column widths are not copied
or altered. Invalid regular expressions are actionable errors. Output is `{ "replacements": int
}` containing the total number of changed cells.

### 4.2 YAML example

```yaml
- id: replace_period_everywhere
  action: replace_text
  workbook: table_creator
  sheet: { matching: "^(BASIS|GLOBAL)$" }
  pattern: "\\d{4}(0[1-9]|1[0-2])"
  replacement: "{{ env.valn_date_yyyymm }}"
```

## 5. `replace_in_range`

### 5.1 Action contract

`replace_in_range` is a writing file action that performs the same replacement only within one
selected contiguous range. `range` accepts a literal A1 cell/range or a workbook-level defined
name resolving to one contiguous area. A single cell is valid. When `range` is a defined name,
its destination sheet takes precedence and `sheet` is ignored, matching `write_range`.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string | Worksheet containing a literal A1 range. |
| `range` | yes | string | A1 cell/range or workbook-level defined name. |
| `pattern` | yes | string | Regular expression to find in cell text. |
| `replacement` | yes | string | Text substituted for every regex match. |

Names resolving to multiple separate areas are rejected. Invalid range references and regular
expressions are actionable errors. Output is `{ "replacements": int }` containing the number of
changed cells.

### 5.2 YAML example

```yaml
- id: replace_period_in_basis
  action: replace_in_range
  workbook: table_creator
  sheet: "BASIS"
  range: "B8:AZ200"
  pattern: "\\d{4}(0[1-9]|1[0-2])"
  replacement: "{{ env.valn_date_yyyymm }}"
```

## 6. `read_table`

### 6.1 Action contract

`read_table` is a read-only file action which discovers a rectangular worksheet table from its
top-left header cell.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string | Worksheet containing the table. |
| `header_cell` | yes | string | Literal A1 reference to the table's top-left header cell, for example `"B7"`. |

Starting at `header_cell`, the action reads headers to the right until the first blank header
cell. It then reads data rows below until the first blank cell in the table's first column. The
header row is included in `values`.

Output shape:

```yaml
output:
  values: [["BASIS_ITEM", "I17_OPN"], ["ACC_RATIO", "na"]]
  headers: ["BASIS_ITEM", "I17_OPN"]
  range: "B7:C8"
```

The action must reject a blank `header_cell`, a duplicate or blank header within the discovered
header sequence, and a table with no data rows. It must not change the workbook.

### 6.2 YAML example

```yaml
- id: read_basis_table
  action: read_table
  workbook: table_creator
  sheet: "BASIS"
  header_cell: "B7"
```

## 7. `copy_table_columns`

### 7.1 Action contract

`copy_table_columns` is a writing file action. It discovers the table using `header_cell`, then
copies every data-row value from each named source column to its corresponding target column.
It leaves the header row unchanged.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string | Worksheet containing the table. |
| `header_cell` | yes | string | Literal A1 reference to the table's top-left header cell. |
| `source_columns` | yes | list of strings | Source header names. |
| `target_columns` | yes | list of strings | Corresponding target header names. |

`source_columns` and `target_columns` must be non-empty and have equal lengths. Every named
header must occur exactly once in the discovered table. The action copies values only; it does
not copy formulas, formatting, validation, comments, or column widths.

### 7.2 YAML example

```yaml
- id: copy_i17_columns
  action: copy_table_columns
  workbook: table_creator
  sheet: "BASIS"
  header_cell: "B7"
  source_columns: ["I17_CLS_DATA", "I17_CLS_DATA"]
  target_columns: ["I17_CLS_DATA_SAM_Infl", "I17_CLS_EXPINF"]
```

## 8. `update_table_cells`

### 8.1 Action contract

`update_table_cells` is a writing file action. It discovers the table using `header_cell`, finds
each requested row by its value in `lookup_column`, finds the corresponding target column by
header name, and writes `value` at each resulting intersection.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string | Worksheet containing the table. |
| `header_cell` | yes | string | Literal A1 reference to the table's top-left header cell. |
| `lookup_column` | yes | string | Header of the column used to locate data rows. |
| `lookup_rows` | yes | list of strings | Values to find in `lookup_column`. |
| `target_columns` | yes | list of strings | Corresponding target header names. |
| `value` | yes | any | Value written to every resolved row/column intersection. |

`lookup_rows` and `target_columns` must be non-empty and have equal lengths. Header names and
lookup values are matched case-insensitively, matching the old script. The lookup column and all
target columns must occur exactly once. Every requested lookup value must occur exactly once;
missing or duplicate matches are an actionable error. The action writes only the specified data
cells and leaves all headers and other cells unchanged.

### 8.2 YAML example

```yaml
- id: set_i17_ivp_values
  action: update_table_cells
  workbook: table_creator
  sheet: "BASIS"
  header_cell: "B7"
  lookup_column: "BASIS_ITEM"
  lookup_rows: ["ACC_RATIO", "ECO_TBL"]
  target_columns: ["I17_IVP", "I17_IVP"]
  value: "na"
```

## 9. `replace_table_text`

### 9.1 Action contract

`replace_table_text` is a writing file action. It discovers the table, resolves row/column
intersections using the same lookup fields as `update_table_cells`, and replaces matching text in
each selected existing cell. It is generic: the old `update_period` operation is one use case,
not a separate period-specific action.

| Field | Required | Type | Meaning |
|---|---|---|---|
| `workbook` | yes | string | Declared workbook name. |
| `sheet` | yes | string | Worksheet containing the table. |
| `header_cell` | yes | string | Literal A1 reference to the table's top-left header cell. |
| `lookup_column` | yes | string | Header of the column used to locate data rows. |
| `lookup_rows` | yes | list of strings | Values to find in `lookup_column`. |
| `target_columns` | yes | list of strings | Corresponding target header names. |
| `pattern` | yes | string | Regular expression to find in each selected cell's text. |
| `replacement` | yes | string | Text substituted for every regex match. |

`lookup_rows` and `target_columns` must be non-empty and have equal lengths. Header names and
lookup values are matched case-insensitively. The action converts the selected cell value to text
only for replacement; it writes the replaced text back. Values with no match remain unchanged.
Invalid regular expressions and missing/duplicate headers or lookup values are actionable errors.

### 9.2 YAML example

```yaml
- id: update_closing_period
  action: replace_table_text
  workbook: table_creator
  sheet: "BASIS"
  header_cell: "B7"
  lookup_column: "BASIS_ITEM"
  lookup_rows: ["I17_DISC_RATE_CRV_TBL", "I17_INFLATION_CRV_TBL"]
  target_columns: ["I17_CLS", "I17_CLS"]
  pattern: "\\d{4}(0[1-9]|1[0-2])"
  replacement: "{{ env.valn_date_yyyymm }}"
```

## 10. Out of scope and backlog

### Not part of this build

- Dynamic source-table copying from another workbook.

Known fixed A1 ranges can already be copied using the existing `copy` action. Tabular text-file
data can already be written using `write_range` once `read_text_file` exists.

### Backlog: type checking and conversion

The text reader returns strings only. A future requirement may add explicit type checking or
conversion for values intended to be Excel numbers, dates, or booleans. It must be deliberately
configured in YAML for the relevant action/columns; automatic conversion based on text content is
not allowed because it can corrupt identifiers such as `00123`.

## 11. Implementation plan and tests

Build test-first, using the project virtual-environment executables.

1. Add failing unit tests for `read_text_file`: `.fac` comma parsing, `.csv` quoting, `.tsv`
   defaults, line-based one-column files, delimiter override, preserved strings, missing files,
   invalid options, and malformed input.
2. Implement the action in `actions.py`, registered as a non-writing control action, using a
   small standard-library helper rather than Pandas.
3. Add failing unit tests for `replace_text` and `replace_in_range`: exact sheet, sheet list,
  `all`, matching-sheet selection, A1 cell/range, defined-name range, changed/nonmatching
  cells, and invalid regex/range/name handling.
4. Implement one shared replacement helper used by `replace_text`, `replace_in_range`, and
  `replace_table_text`; do not duplicate cell text replacement logic.
5. Add failing unit tests for `read_table`, `copy_table_columns`, `update_table_cells`, and
  `replace_table_text`: normal discovery, variable width/height, blank/duplicate/missing
  headers, no data rows, missing/duplicate lookup rows, paired-list length validation,
  copy/update behavior, regex replacement, unchanged nonmatches, and invalid regex handling.
6. Implement one shared table-boundary discovery helper used by all table actions; do not
  duplicate the header and first-column scanning logic.
7. Add failing file-backend and xlwings-backend tests for `write_cell` and `write_range` using
   single-cell, multi-cell, and multi-area defined names.
8. Extend both backend write primitives to resolve defined names before writing; then extend
   existence validation for the two write reference fields.
9. Add workflow-level integration coverage: read a `.fac` fixture, pass its `values` directly
   to `write_range`, discover a table, copy named table columns, update named row/column
  intersections, replace text in a sheet/range/table, and verify a write through a defined name.
10. Update the YAML skill and README action reference with the action syntax, text-only warning,
  extension defaults, defined-name write support, and all six table/replacement actions.

## 12. Acceptance criteria

- A `.fac` file is read without modification and its string rows can be passed directly to
  `write_range`.
- CSV quotes and embedded delimiters are parsed correctly by the standard library.
- A newline-only text file yields one string column.
- No value is inferred or converted to a number, date, or boolean.
- `write_cell` and `write_range` work with literal A1 references exactly as they do today.
- `write_cell` and `write_range` work with a valid single-area workbook-level defined name on
  both file and xlwings paths.
- Invalid, multi-area, and unsuitable named-range references fail before a workbook is committed.
- `replace_text` replaces matching text throughout exact, listed, all, and regex-selected sheets.
- `replace_in_range` replaces matching text only in its literal A1 or single-area defined-name
  target, leaving cells outside the target unchanged.
- `read_table` discovers a table from one `header_cell` reference and returns its values,
  headers, and exact A1 bounds without changing the workbook.
- `copy_table_columns` copies all data rows between columns addressed by header names, without
  changing headers or requiring hard-coded final row/column coordinates.
- `update_table_cells` resolves named row/column intersections and writes each requested value
  without changing headers or unrelated data cells.
- `replace_table_text` performs regex replacement only in the requested row/column
  intersections, including the old `YYYYMM` period-update use case.