# 08_full_showcase — planning script (v2, revised per feedback)

## Quick answers first

- **Can `id:` contain spaces?** No. Per the YAML skill doc (section 4): an `id` "must be a
  valid Jinja identifier: letters, digits, underscore, not starting with a digit." Spaces
  (and hyphens) are invalid — every `id` below is `snake_case`.
- **How will this be validated? Can we dogfood another workflow for it?** Yes — that's the
  plan. There are now **two** workflow files, not one:
  - **`08_full_showcase.yaml`** — the demo/build workflow. It builds `report.xlsx` from
    `catalog.xlsx` and writes fresh formulas, but contains **no assertions** — it only
    demonstrates actions.
  - **`08_full_showcase_validate.yaml`** — a second, separate workflow, run straight after the
    first, whose entire job is to open the two workbooks the first one produced/touched, read
    back specific cells, and use `stop` + `if:` guards to assert each one matches its
    independently-known expected value. This *is* excel_runner validating excel_runner's own
    output — no bespoke pytest/python assertion script needed for the demo itself (a real
    pytest integration test can still separately run both workflows in sequence and check the
    validate workflow's own `status` came back `"success"`, but the assertion logic itself
    lives in YAML).

## The story (why these steps happen in this order)

Both workbooks are generic and small:

- **`catalog`** (source, already exists) — "Products" sheet: header row + 5 product rows
  (Widget/Gadget/Gizmo/Sprocket/Cog), columns Product/Price/Quantity/Total (`Total` is a
  formula, `=Price*Quantity`). One defined name, `GrandTotalName`, over a single formula cell
  (`=SUM(Total column)`) elsewhere on the sheet — no separate defined name over the header
  row itself; a named range spanning only one header row is unnecessary ceremony when a plain
  by-sheet range does the same job. One disposable "Scratch" sheet. **No `open` step for
  `catalog` — its session opens implicitly**
  the moment the first real step below touches it (`find_products_headers`), which is
  precisely the point: most workbooks in this showcase never get an explicit `open` at all,
  because they don't need one.
- **`report`** (target, blank/`create_if_missing`) — built up entirely by the steps below.
  **`report` *does* get an explicit `open` step**, deliberately placed right before its first
  real use — so the file contrasts implicit open (`catalog`) against explicit open (`report`)
  side by side instead of using `open` as generic step 1 filler.

Every `find_*` step's output is consumed by a later step (no orphaned lookups), and the two
workbooks are genuinely **linked**: `report` gets a formula that references a `catalog` cell
by external reference (`'[catalog.xlsx]Products'!...`), so the final cross-workbook
`recalculate` step has a real reason to use `scope: all` — it's the only scope that recalculates
both workbooks' formulas together, including one that depends on the other.

### `08_full_showcase.yaml`

- **`find_products_headers`** — `find_headers_row` on `catalog`, `search_range: "A1:D3"`
  (a plain by-sheet range covering the top few rows, not a defined name — see note above),
  `patterns: ["Product", "Price", "Quantity", "Total"]`
  - Demonstrates header discovery by scanning a plain sheet range for the first row matching
    every pattern. Implicitly opens `catalog`. Its `row` output feeds every `header_row:`
    below. (Named-range resolution on `search_range` is still real/available on this action —
    it's just not exercised *here*, since a header row is too narrow a thing to name.)
- **`find_price_column`** — `find_column`, `header_row` from `find_products_headers.output.row`,
  `pattern: "^Price$"`
  - Its `column` output is used later to build a dynamic A1 reference for a per-product
    lookup.
- **`find_key_columns`** — `find_columns` (plural), same `header_row`, dict patterns for
  `product`, `total`
  - Its `product`/`total` outputs are used to build every A1 range referenced from here on
    (`copy`'s source range, the new formula's column, the linked-cell reference) — nothing
    downstream hardcodes a column letter once this step has run.
- **`find_widget_row`** — `find_row`, `column: "{{ steps.find_key_columns.output.product }}"`,
  `search_value: "Widget"`, `header_row` from `find_products_headers`
  - Its `row` output pins down exactly *which* product row the linked-cell demo (below) points
    at — "Widget"'s own `Total` cell.
- **`read_totals_as_values`** — `read_range`, range built from `find_key_columns.output.total`
  (e.g. `"{{ steps.find_key_columns.output.total }}2:{{ steps.find_key_columns.output.total }}6"`),
  `formula: false` (default)
  - Computed values for every product's `Total`.
- **`read_totals_as_formulas`** — `read_range`, same block, `formula: true`
  - Same cells, formula text instead — deliberately reused so the two outputs are directly
    comparable (`=Price*Quantity` vs. the number), not two unrelated reads.
- **`read_named_grand_total`** — `read_range`, `range: "GrandTotalName"` (the defined name,
  read directly — no `read_metadata` detour needed for a single-cell named range)
  - This becomes the "known-good expected total" the validate workflow checks the rebuilt
    total against later.
- **`read_products_table`** — `read_range`, `range: "A1:D6"` (the whole table, header row
  included — 5 rows x 4 columns)
  - A genuinely 2D capture: `output.values` in `steps_dump.json` is a nested list
    (`values[row][col]`), not a flat list — `read_totals_as_values` above only ever reads a
    single column, so it doesn't show this shape. `write_summary_block` (below) indexes into
    it directly (`values[find_widget_row.output.row - 1][0]`) to pull the Widget row's product
    name back out, demonstrating `{{ }}`'s Jinja2 list-indexing/arithmetic against a captured
    2D step output, not just its scalar/1D outputs.
- **`recalculate_catalog`** — `recalculate`, `workbook: catalog`, `scope: workbook`
  - Makes sure `catalog`'s formulas (including `GrandTotalName`) are fresh before anything
    reads from or copies it.
- **`open_report`** — `open`, `workbook: report`
  - The explicit-open contrast case described above — `report`'s session opens here on
    purpose, right before its first real use, instead of implicitly on the `copy` step below.
- **`read_catalog_properties`** — `read_metadata`, `target: properties`, `workbook: catalog`
  - No longer a dead end: its `title` output is written into `report` two steps later as a
    provenance label, so it's a real value flowing downstream, not an orphaned read.
- **`copy_products_to_report`** — `copy`, `source.workbook: catalog` / `source.sheet:
  "Products"` / `source.range` covering the whole table, `target.workbook: report` /
  `target.sheet: "Summary"` / `target.range: "A1"`
  - Proves formulas survive the COM copy (the `Total` column arrives as `=Price*Quantity`
    formulas, not frozen numbers) — the whole reason `copy` needed the COM rewrite.
- **`write_provenance_label`** — `write_cell` on `report`/"Summary", a cell below the copied
  block, `value: "Source: {{ steps.read_catalog_properties.output.title }}"`
  - Consumes `read_catalog_properties`'s output — see above.
- **`write_linked_total_cell`** — `write_cell` on `report`/"Summary", `cell` built from
  `find_key_columns`/`find_widget_row` outputs (one column over from the copied block, same
  row as "Widget"), `value: "='[catalog.xlsx]Products'!{{ steps.find_key_columns.output.total }}{{ steps.find_widget_row.output.row }}"`
  - The genuine **cross-workbook link**: `report` now has a live formula referencing `catalog`
    directly, not a copied value. This is what makes the final `recalculate` step's
    `scope: all` meaningful rather than cosmetic.
- **`write_new_grand_total_cell`** — `write_cell` on `report`/"Summary", one row below the
  copied block, `value` a `SUM(...)` formula over the *copied* `Total` column (still using
  `find_key_columns.output.total` for the column letter, since `copy` preserved the layout)
  - A brand-new formula written after the copy, proving the copied formulas and a fresh
    formula recalculate together correctly later.
- **`write_summary_block`** — `write_range` on `report`/"Summary", a small two-column legend
  (plain values, e.g. `[["Field", "Meaning"], ["Total", "Price × Quantity"]]`), plus a third
  row whose value is `{{ steps.read_products_table.output.values[steps.find_widget_row.output.row - 1][0] }}`
  - The plain multi-cell `write_range` example, extended to also show a 2D step output being
    indexed by `[row][col]` from inside a template expression — see `read_products_table`
    above.
- **`write_summary_row_by_name`** — `write_row`, dict mode, a "Generated" label row
  - Dict-mode `write_row` example.
- **`write_summary_row_positional`** — `write_row`, list mode with `start_column`
  - Positional-mode `write_row` example, directly contrasted with the previous step.
- **`insert_note_column`** — `insert_range`, whole-column insert (`at: "A:A"`) with a `header:`
  label on "Summary", run *after* the block above exists
  - Demonstrates a structural edit shifting already-written data, rather than inserting into
    an empty sheet (which would prove nothing).
- **`autofit_summary_columns`** — `set_column_width`, `columns: "A:F"`, `width: "autofit"`
  - Runs after the column insert, so there's a genuine reason to refit widths.
- **`add_archive_sheet`** → **`rename_archive_sheet`** → **`delete_scratch_sheet`**
  - `create_sheet` ("Archive" on `report`), `rename_sheet` ("Archive" → "Archived" on
    `report`), `delete_sheet` (removes `catalog`'s disposable "Scratch" sheet) — sheet
    lifecycle management, one step each.
- **`recalculate_everything`** — `recalculate`, `scope: all`, `mode: full_rebuild`
  - Recalculates `catalog` **and** `report` together in one pass — required because
    `write_linked_total_cell` only resolves correctly once both workbooks' formulas are
    current at the same time. `scope: workbook` here would leave the link stale.
- **`save_report`** — `save`, `workbook: report`
- **`close_catalog`** — `close`, `workbook: catalog` — no longer needed once recalculated and
  copied from.

### `08_full_showcase_validate.yaml` (separate file — the dogfooded validation)

Opens `report.xlsx` (the output of the file above) plus `catalog.xlsx` again (read-only, for
the independently-known expected total), and asserts purely via `read_range` + `stop`/`if:`:

- **`read_expected_total`** — `read_range` on `catalog`, `range: "GrandTotalName"`
  - The known-good expectation, recomputed fresh (not reused from the build run).
- **`read_actual_new_total`** — `read_range` on `report`/"Summary", the cell
  `write_new_grand_total_cell` wrote
  - The value to check.
- **`guard_total_mismatch`** — `stop`, `if: "{{ steps.read_actual_new_total.output.values !=
  steps.read_expected_total.output.values }}"`, `reason: "rebuilt grand total does not match
  catalog's own GrandTotalName"`
  - Never fires in the happy path; fires (and the run's own `status` reflects it) if the
    showcase's arithmetic chain broke anywhere.
- **`read_actual_linked_cell`** — `read_range` on `report`/"Summary", the linked-formula cell
  `write_linked_total_cell` wrote
- **`guard_linked_cell_mismatch`** — `stop`, `if:` comparing that value against
  `read_expected_total`'s (or the specific Widget-row total, whichever is the correct
  expectation) with its own descriptive `reason:`
  - Proves the cross-workbook link recalculated correctly, independent of the copy-based
    total above.
- A closing **`read_metadata`**, `target: cells`, `cells: ["GrandTotalName"]` on `catalog`
  - Placed here, at the very end, purely as a standalone illustrative example of
    `read_metadata`'s `cells` mode resolving a defined name — its output isn't wired into
    any guard (the same value is already proven correct above via `read_range`), it exists
    only to show this action's own named-range support once, on its own.

A real `tests/integration/` test can run both workflows back to back and assert the *validate*
workflow's own top-level `status == "success"` — but the interesting per-cell assertions are
expressed in YAML, by excel_runner, using its own actions.

## Coverage check (all 20 actions, ≥1 use each, every output consumed)

`open` (explicit, `report`), `save`, `close`, `stop` (×2, validate workflow), `copy`,
`read_range` (×5 across both files), `read_metadata` (×2: properties mid-story + a closing
cells example), `write_cell` (×3), `write_range`, `write_row` (×2 modes), `insert_range`,
`set_column_width`, `create_sheet`, `rename_sheet`, `delete_sheet`, `find_headers_row`,
`find_row`, `find_column`, `find_columns`, `recalculate` (×2 scopes, the second one genuinely
cross-workbook).

## Notes for the build phase (not yet built)

- A temporary generation script builds `catalog.xlsx` and a blank `report.xlsx`.
- Generated **originals** get committed to `demos/08_full_showcase/originals/` so both
  workflows can be re-run repeatedly from a clean slate without regenerating them.
