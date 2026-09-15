# Preflight Table Scan Performance Investigation

**Status:** implemented and verified locally on 2026-09-14. This remains a temporary record of
the measured problem, design, and residual risks.

## Purpose

This note records why read-only workflow preflight was slow for the translated Table Creator
workflow, the smallest proposed change, and the correctness and safety risks that must be
addressed before implementing it.

The analysis concerns `--dry-run` and the preflight phase of `--check-existence`. It does not
concern workbook execution, saving, recalculation, or Excel/COM.

## Scenario And Files

The local test workflow is:

- `temp/table_creator_workflow.yaml`

It uses these copied local inputs:

- Source Table Creator workbook: `temp/table_creator_fixture/originals/table_creator.xlsm`
- Source Table Creator sheet: `BASIS`
- Repeated table anchor: `B7`
- Yield-curves workbook: `temp/table_creator_fixture/input/Economic Tables June 2026.xlsx`
- Yield-curves sheet: `IFRS17 Curves`
- Yield-curves table anchor: `B2`
- FAC inputs: `temp/table_creator_fixture/input/prior_tables/CentralCurves/` and
  `temp/table_creator_fixture/input/prior_tables/EXPENSES/`

The fixture copies are separate from the originally supplied files under `temp/`. Preflight opens
workbooks read-only and does not write to either location.

## The Problem

`excel_runner.engine.validate_existence()` performs literal table-reference validation for each
`read_table`, `copy_table_columns`, `update_table_cells`, and `replace_table_text` step.

For every such step, `_validate_table_references()` currently:

1. Locates the specified worksheet and header cell.
2. Reads individual cells across the header row until the first blank header.
3. Reads individual cells down the first table column until the first blank row.
4. Checks requested source/target/lookup columns against the discovered headers.
5. For each requested lookup value, loops the discovered table rows to confirm that it occurs
   exactly once.

This is correct in intent: it catches a missing sheet, bad header anchor, blank/duplicate header,
empty table, missing column, and missing/duplicate lookup value before execution begins.

The cost becomes high when many steps refer to the same table. The translated workflow contains
31 table-targeting steps:

| Workbook logical name | Sheet | Header cell | Step count |
|---|---|---:|---:|
| `table_creator` | `BASIS` | `B7` | 30 |
| `yield_curves` | `IFRS17 Curves` | `B2` | 1 |

Thus, the `BASIS!B7` structure and rows are rediscovered 30 times in the same read-only
preflight even though no action has run and the source workbook cannot have changed.

## Measured Evidence

All measurements below were made on 2026-09-14 using the project venv, `openpyxl`, and this
source workbook/sheet:

- Workbook: `temp/table_creator_fixture/originals/table_creator.xlsm`
- Sheet: `BASIS`
- Anchor: `B7`

No workbook was saved, no Excel process was launched, and no source file was modified.

| Measurement | Result |
|---|---:|
| Source `.xlsm` size | 6,539,165 bytes |
| Read-only workbook load (`read_only=True`, `keep_vba=True`) | 0.813249 s |
| `BASIS` reported dimensions | 1,195 rows x 124 columns |
| Existing-style random `worksheet.cell()` table scan | 2.502755 s |
| Sequential `worksheet.iter_rows(..., values_only=True)` table scan | 0.052802 s |
| Random/sequential relative cost | about 47.4x |

Both scan methods found the same table boundary:

- Header row: `B7:AJ7`
- First blank first-column cell: `B102`
- Data rows: `8:101`

The initial full `--dry-run` was manually stopped after more than 30 seconds without output. It
did not execute workflow actions, create a working workbook, save a workbook, recalculate, or
launch Excel.

Using the measured `BASIS` scan time as a simple estimate:

$$
30 \times 2.502755\text{s} = 75.082650\text{s}
$$

Adding the one other table step gives an estimated repeated-scan cost of approximately
$77.585405\text{s}$. This is an estimate, not an observed end-to-end preflight duration, because
it excludes YAML loading, workbook opening, named-range/sheet checks, FAC parsing, and the yield
curves table scan.

If each distinct literal table were scanned once with the sequential method, the corresponding
scan estimate is:

$$
2 \times 0.052802\text{s} = 0.105604\text{s}
$$

This estimate assumes the yield-curves table has a comparable scan cost. It is intended to show
the scale of the improvement, not promise a final total runtime.

## Implemented Solution

The runner now builds an in-memory **preflight table index** once for each unique literal table
reference.

The cache key should identify the workbook being inspected and table anchor:

```text
(inspected workbook path, sheet name, header cell)
```

The cached value should contain only information already required by current validation:

- Canonical/case-folded header names and their column offsets.
- The detected table boundary.
- Case-folded lookup-column values and the workbook row number(s) at which each appears.

The index should be built with a sequential `iter_rows()` pass rather than repeated
`worksheet.cell(row, column)` calls. A later step then validates its requested columns and lookup
rows from the index rather than rescanning the workbook.

Implemented flow:

1. `validate_existence()` opens each inspected workbook read-only once, as it does today.
2. For a literal table action, derive its cache key.
3. On the first occurrence of a key, scan the table sequentially and validate its structural
   rules.
4. Store the resulting index for this preflight invocation only.
5. On later occurrences of the same key, reuse the index to validate the step's requested
   columns and lookup values.
6. Close the workbooks and discard the cache in `finally`, as the current method already closes
   workbooks.

The cache must not persist to disk, cross workflow runs, or cross workbooks. It is only an
optimization of information already read during one preflight.

## Observed Result

After implementation, the following read-only command completed successfully:

```powershell
.venv\Scripts\python .\excel_runner\cli.py .\temp\table_creator_workflow.yaml --dry-run
```

Measured wall-clock time: `1.400 seconds`.

The command reported `Preflight complete: no workflow actions were executed.` It did not create
the working workbook, save a workbook, launch Excel, or recalculate.

The $1.400\text{s}$ measurement includes workflow loading, source workbook opening, table/sheet/
defined-name validation, and FAC parsing. It cannot be compared directly to a full old end-to-end
timing because the pre-optimization run was manually stopped after more than 30 seconds without
output. It does, however, confirm that the repeated table scan no longer blocks this workflow.

## Expected Benefits

- A preflight cost that scales with the number of distinct literal tables, rather than the number
  of table operations.
- For this workflow, one scan of `BASIS!B7` instead of 30 scans.
- No change to the YAML action vocabulary or user-authored workflow files.
- No workbook mutation, staging, Excel launch, save, commit, or recalculation.
- Faster feedback during authoring, especially for macro-enabled workbooks with many table
  operations.

## Risks And Possible Damage

### Validation Regression

A faulty index can weaken or change existing validation. It must retain the current behavior for:

- Blank, non-text, and duplicate headers.
- A blank table anchor.
- A table with no data rows.
- Missing source, target, or lookup columns.
- Lookup rows that occur zero times or more than once.
- Case-insensitive header and lookup matching.
- Template-derived `sheet`, `header_cell`, column names, or lookup values, which are intentionally
  deferred until runtime rather than guessed during preflight.

Mitigation: retain the existing error messages and add unit tests that compare cached and
uncached outcomes for successful tables and each failure mode.

### Wrong Cache Key

Caching only by `sheet` and `header_cell` could incorrectly combine two different workbooks that
happen to have the same sheet name and anchor.

Mitigation: include the resolved inspected workbook path in the key. The logical workbook name
alone is insufficient because different logical names can point at the same or different paths.

### Stale Data During Execution

The workflow can modify tables after preflight begins. Reusing a preflight index during execution
would be incorrect because the workbook state can legitimately change between actions.

Mitigation: use the index only inside `validate_existence()`. Discard it before normal execution
starts. Runtime table actions must continue to discover the live table from the active workbook.

### Existing Time-Of-Check/Time-Of-Use Gap

An external user/process can edit a source workbook after preflight completes but before the
execution phase opens it. This is already possible today and caching neither introduces nor fixes
it.

Mitigation: document that preflight validates the file state observed at preflight time. A future
stronger guarantee would require file locking or source fingerprints, which is outside this
performance change.

### Large Or Malformed Worksheets

A sequential scan still needs to find the first blank header and first blank first-column data
row. A malformed or intentionally hostile workbook could have a populated first column far down
the sheet, so any table-discovery approach can be expensive.

Mitigation: preserve the current table semantics initially, but consider an explicit maximum
scan limit and clear validation error in a separate, documented change. Do not silently truncate
a table during this optimization.

### Memory Use

Caching full table cell values could consume substantial memory for very large tables.

Mitigation: store only headers, boundary information, and lookup-column row locations needed for
validation. Do not cache every table cell.

### Macro/VBA Preservation

The source Table Creator workbook is `.xlsm`. Any accidental write/save through openpyxl could
risk VBA/package preservation.

Mitigation: preflight must keep using read-only workbook loading, never call save, and close the
handle after validation. The proposed cache is Python memory only and does not change the file.

### Audit And User Expectations

A faster preflight may reveal later missing FAC errors sooner, which is correct but can look like
a behavior change compared with a previously stalled run.

Mitigation: keep the existing ordered validation model and clear errors. Consider logging one
short debug message per newly indexed table, not per reused lookup, so normal CLI output remains
quiet.

## Verification Completed

1. A unit test with two table actions against the same workbook/sheet/anchor asserts the table
   inspector is invoked once.
2. The complete existence-validation module passed: `19 passed in 0.90s`.
3. The preflight integration tests passed: `5 passed in 0.93s`.
4. The translated workflow's `--dry-run` completed successfully in `1.400 seconds`.
5. The separate-cache and template-bypass cases remain desirable additional unit tests before
   considering the optimization fully hardened for arbitrary workflows.

## Decision Needed

The optimization is low-risk if it stays preflight-local, caches only derived validation metadata,
and preserves existing table-discovery semantics. It should not be implemented as a general
runtime table cache.

Before using `--check-existence` against the real fixture, run lint and the broader relevant test
suite, then review the translated workflow and its output-validation plan. The cache itself does
not make execution safe; it only makes read-only preflight efficient.
