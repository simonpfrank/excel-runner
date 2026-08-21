# excel_runner — Test Summary

One-line-per-test overview, grouped by file/class, so it's easy to see what's covered without
reading every test body. See [Progress_Tracker.md](Progress_Tracker.md) for narrative history
and [Specification.md](Specification.md) for the design each area implements.

285 tests total: 15 integration, 270 unit.

## Integration tests (`tests/integration/test_run_workflow.py`)

Full-stack, zero-mock tests: real YAML files, real openpyxl workbooks, run through
`run_workflow()` end to end.

**TestHappyPath**
- Reads and writes succeed, and the workbook is saved automatically at the end — no explicit `save` step needed.
- An explicit `save` + `close` step still works the same way.

**TestIfConditions**
- A step whose `if:` evaluates false is skipped, not run.

**TestStepFailureDoesNotCrashTheRun**
- A normal search miss (e.g. `find_row` finding nothing) doesn't stop later steps from running.
- A run with any failed step never commits its changes to the real file.

**TestExceptionsPropagateAndStillCleanUp**
- A genuine authoring mistake (a raised exception, not a normal "didn't find it") propagates and leaves the real file untouched.

**TestValidationRunsBeforeAnyWorkbookIsTouched**
- An invalid workflow raises before any real file is opened or touched.

**TestCopyAcrossTwoWorkbooks**
- `copy` copies a range from one workbook into another correctly.

**TestStop**
- `stop` after a failed lookup marks every later step `"stopped"` and the run doesn't commit.
- `stop` on a deliberate early exit (no prior failure) still commits everything that ran before it.
- A `stop` step with a false `if:` is itself skipped, not triggered.
- Steps marked `"stopped"` still get an audit-log record.

**TestAuditLog**
- The audit log has exactly one record per step.

**TestCrashSafety**
- A crash mid-run leaves the real file untouched and the scratch copy in place for recovery.
- A later, valid run against the same workbook succeeds even after a previous crash.

## Unit tests — actions (`tests/unit/actions/*.py`)

Each action tested directly against a real openpyxl fixture workbook wrapped in a
`WorkbookSession` (`conftest.py`'s `file_session` fixture) — no mocks.

**test_close.py — `close`**
- Registers as a file action.
- Closes the session's handle and returns success.

**test_copy.py — `copy`**
- Registers as a file action.
- Copies an explicit range from source to target.
- Marks the target session dirty, not the source.

**test_create_sheet.py — `create_sheet`**
- Registers as a file action that writes.
- Adds a new sheet to the workbook.
- A duplicate sheet name returns a structured error, not a crash.

**test_delete_sheet.py — `delete_sheet`**
- Registers as a file action that writes.
- Removes a sheet from the workbook.
- Deleting the workbook's only remaining sheet returns a structured error, not a crash.

**test_find_column.py — `find_column`**
- Registers as a file action.
- Output shape matches PRD §10.4 (keyed object).
- Not found returns a structured error, not a crash.

**test_find_columns.py — `find_columns`**
- Registers as a file action.
- Output shape matches PRD §10.4.
- Names whose pattern didn't match anything are simply omitted, not an error.

**test_find_headers_row.py — `find_headers_row`**
- Registers as a file action.
- Output shape matches PRD §10.4.
- No matching row returns a structured error, not a crash.

**test_find_row.py — `find_row`**
- Registers as a file action.
- Output shape matches PRD §10.4.
- Not found returns a structured error, not a crash.

**test_insert_range.py — `insert_range`**
- Registers as a file action.
- Inserts a whole column with a header written into the new row.
- A partial range (not whole-row/whole-column) returns a structured error, not a raw exception.

**test_open.py — `open`**
- Registers as a file action.
- Returns success.
- Has no meaningful output.

**test_read_metadata.py — `read_metadata`**
- Registers as a file action.
- `target: "properties"` reads document properties.
- `target: "cells"` reads a scattered list of specific cells.
- `target: "cells"` without `sheet`/`cells` raises a clear error.
- An unsupported `target` value raises explicitly rather than silently acting as `"cells"`.

**test_read_range.py — `read_range`**
- Registers as a file action.
- Output is keyed under `values`, per PRD §10.4.
- Reads a multi-cell range.
- Returns success status.

**test_rename_sheet.py — `rename_sheet`**
- Registers as a file action that writes.
- Renames the sheet.

**test_save.py — `save`**
- Registers as a file action.
- Saves pending changes to the session's path.
- Saves to a different path than the original when the session path differs.

**test_set_column_width.py — `set_column_width`**
- Registers as a file action.
- Sets an explicit width.
- `"autofit"` sizes the column to its content.

**test_stop.py — `stop`**
- Registers as a control action (no backend, no `session`).
- Returns success with empty output by default.
- Returns the `reason` in output when one is given.

**test_write_cell.py — `write_cell`**
- Registers as a file action.
- Writes the value.
- A string starting with `"="` is written unchanged as a formula.
- Has no meaningful output.

**test_write_range.py — `write_range`**
- Registers as a file action.
- Writes a 2D block of values.
- Marks the session dirty.

**test_write_row.py — `write_row`**
- Base mode: registers as a file action; writes by explicit column mapping.
- Positional mode: writes values in order starting from a given column.
- Positional mode without `start_column` raises a clear error.

## Unit tests — action types, registry, public API

**test_action_types.py**
- `ActionResult` defaults (`error=None`) and is frozen.
- `WorkbookSession` constructs correctly and its `dirty` flag is mutable.
- `@file_action`/`@xlw_action`/`@com_action` each register the right capability.
- Every capability decorator returns the original function unchanged.

**test_registry.py — `discover_actions()`**
- Finds every capability-tagged function in `actions.py`.
- Each entry is an `ActionSpec`.
- Carries a description from the function's docstring (first line only, not the whole thing).
- Carries the real callable (`fn is actions.read_range`).
- Carries its capability (`"file"`, etc.).
- `stop` registers with capability `"none"` (no backend).
- Each `ActionSpec.name` matches its function's real name.
- Param schema excludes `session`, includes every other param, marks only no-default params required.
- An action with no extra params gets an empty schema.

**test_list_actions.py — `list_actions()`**
- Returns a tuple of `ActionSpec`s.
- Includes every one of the 19 built actions.
- Every entry has a description.
- Consistent with what `discover_actions()` itself returns.

**test_public_api.py — the `excel_runner` package's public surface**
- `run_workflow`/`list_actions` importable from the top-level package are the real functions.
- Result types (`RunResult`, `StepResult`, etc.) and workflow-construction types are the real classes.
- `ActionSpec` is the real class.
- `list_actions()` works when called through the public import path.

## Unit tests — backends (`tests/unit/test_backends*.py`)

Raw openpyxl-mechanics tests, no `ActionResult`/session wrapping.

**test_backends.py**
- `open_workbook`: opens read-write by default; opens read-only when asked.
- `read_range`: a single cell as a scalar; a multi-cell range as a 2D list; a single-row range as a 2D list.
- `write_cell`: writes a value; writes a formula string as-is.
- `save_workbook`: saves changes to the given path.
- `close_workbook`: doesn't raise; opening a missing file raises a clear `FileNotFoundError`.

**test_backends_batch2.py**
- `write_range`: writes a 2D block anchored at the top-left cell; a single-cell range still works.
- `set_column_width`: explicit width; a range of columns; `"autofit"` sizes by longest content.
- `insert_range`: inserts a whole column; with a header; inserts a whole row; a partial range raises a clear error.
- `copy_range`: copies an explicit range between workbooks; copies the whole sheet when `source_range` is omitted.

**test_backends_batch3.py**
- `find_headers_row`: finds the row matching every pattern; returns `None` when no row matches; a single-cell search range doesn't crash.
- `find_row`: finds a matching value; returns `None` when not found; works without a header row.
- `find_column`: finds a column by exact pattern; by regex pattern; returns `None` when not found.
- `find_columns`: finds multiple named columns; omits names that didn't match.
- `read_properties`/`read_cells`: reads document properties; reads specific scattered cells.

**test_backends_batch4.py**
- `create_workbook`: creates a blank workbook; creates one from a template.

**test_backends_sheets.py**
- `create_sheet`: adds a new empty sheet; inserts at a given index; a duplicate name raises `ValueError`.
- `rename_sheet`: renames an existing sheet.
- `delete_sheet`: removes a sheet; deleting the only sheet raises `ValueError`.

**test_backends_xlw.py** (xlwings-backed, live Excel)
- `xlw_open_workbook`: opens an existing workbook; a missing file raises `FileNotFoundError`.
- `xlw_close_workbook`: closes the book without quitting the whole app.
- `xlw_save_workbook`: saves in place and the change persists.

## Unit tests — validation (`tests/unit/test_validation.py`)

Both tiers: tier 1 (static schema, no workbook access) and tier 2 (dry-run step-graph planning).

**TestActionExists**
- A known action passes.
- An unknown action raises with a fuzzy-matched suggestion.
- An unknown action with no close match omits the suggestion.

**TestUnknownParams**
- An extra param not in the action's schema raises.
- `copy` is exempt from schema-shape checks (its `source:`/`target:` shape doesn't match its Python signature).

**TestRequiredParams**
- A missing required param raises.
- `stop` is exempt from the implicit `workbook` requirement.
- A missing implicit `workbook` field raises.

**TestParamTypes**
- A field expecting a list but given a string raises, with a "wrap it in `[ ]`" suggestion.
- A correctly-typed value passes.
- A `Literal` type rejects a value outside the allowed set.
- A union type rejects a value matching neither branch.
- Type-name fallback works for a plain type.

**TestStepReferences**
- A reference to a nonexistent step raises with a suggestion.
- A reference to a later step raises (forward references aren't allowed).
- A reference to an earlier step passes.
- Non-string param values aren't scanned for step references.
- A reference inside `if:` is checked too, not just step params.

**TestDryRunWorkbooksDeclared**
- A workbook name not in the `workbooks:` registry raises.
- Every workbook actually declared passes.

**TestExecutionPlanModeInference**
- A workbook never written to is planned `read_only`.
- A workbook written to by any step is planned `read_write`.
- A workbook referenced by both a read and a write action across steps is `read_write`.
- `copy` marks every workbook it touches as `read_write` (source and target can't be told apart statically).
- Workbook references nested inside a list (not just a flat dict) are still found.
- A workbook never referenced by any step defaults to `read_only`.

## Unit tests — loading, schema, templating

**test_loader.py**
- The `env:` block is captured.
- Workbook file paths are resolved against `env:` templating.
- Workbook registry fields (`create_if_missing`, `template`, etc.) are captured.
- Steps are captured in file order.
- Step params are left raw/unresolved at load time (resolved later, per step, during execution).
- `id`/`action`/`if` aren't duplicated into a step's `params`.
- `env_overrides` take precedence over the file's own `env:` block.
- `env_overrides` can add brand-new keys not in the file at all.
- Unquoted `yes`/`no`/`on`/`off` stay plain strings (YAML 1.2 behavior, not 1.1's boolean coercion).
- `true`/`false` (and case variants) still resolve as real booleans.

**test_schema.py**
- `WorkbookRef`/`Step`/`Workflow` dataclasses: defaults, explicit fields, frozen-ness, equality where relevant.
- `Step.params` stays a mutable dict for the caller even though `Step` itself is frozen.
- `Workflow.steps` is a tuple, not a list.

**test_templating.py**
- A plain string with no `{{ }}` is returned unchanged; non-string scalars pass through untouched.
- A dict key literally named `values`/`keys`/`items` resolves to the item, not the shadowing `dict` method.
- Real attribute access still works for non-dict objects.
- A value that's entirely one `{{ }}` expression resolves to its native Python type (int, dict, bool).
- Surrounding whitespace around a whole expression still counts as "whole".
- An expression embedded in a longer string is stringified, including when the embedded value is an int.
- Recursion resolves dict values, computed dict keys, list items, and nested structures.
- An undefined reference raises `ValidationError`, keeping the technical reason separate from the message.
- An undefined reference embedded in a larger string also raises.
- A syntactically invalid expression raises `ValidationError`.
- Two adjacent `{{ }}` blocks render as a concatenated string, not a single "whole expression".
- `evaluate_condition`: a wrapped expression evaluates true/false; a bare expression (no `{{ }}`) works too; non-boolean and falsy-zero results coerce correctly.

## Unit tests — errors (`tests/unit/test_errors.py`)

- `ErrorDetail`: defaults, explicit fields, is frozen.
- `ExcelRunnerError`: carries its `ErrorDetail`; `str(error)` is the plain-English message, not the technical reason; is a real `Exception`.
- `ValidationError`/`ActionExecutionError` are both `ExcelRunnerError`s, but distinct from each other.
- Catching the base `ExcelRunnerError` class catches both subclasses.

## Unit tests — audit logging (`tests/unit/test_audit.py`)

- Writes one JSON line per step.
- Multiple steps append rather than overwrite the log file.
- An `ErrorDetail` is serialized correctly into the log.
- A skipped step is recorded too, not omitted.
- Creates parent directories for the log path if they don't exist.

## Unit tests — scratch-copy execution model (`tests/unit/test_scratch.py`)

- `stage()`: copies the real file into the scratch dir; handles a real file that doesn't exist yet.
- `commit()`: moves scratch content back to the real path; `commit_all()` commits every staged workbook; creates the real file when it didn't exist before.
- `cleanup()`: wipes the scratch dir after a successful `commit_all()`; keeps the scratch dir by default if nothing was committed; doesn't raise on a scratch dir that was never created.

## Unit tests — session management (`tests/unit/test_session_manager.py`)

**TestGetOrOpen**
- Opens an existing workbook read-write by default.
- Read-write mode stages through scratch; read-only mode does not stage.
- A second call for the same name returns the cached session.
- An unknown workbook name raises a clear error.
- A missing file without `create_if_missing` raises a clear error.

**TestNeededBackend**
- `"file"` capability needs the file backend; `"xlw"` and `"com"` both need the xlw backend.
- `"depends_on_param"` and `"none"` capabilities aren't resolvable here (need a live action call).

**TestCapabilityBackendMismatch**
- A matching capability returns the session normally.
- The default capability (unset) is treated as `"file"`, unchanged from before this param existed.
- A mismatched capability raises clearly, whether the session is brand new or already open.

**TestCreateIfMissing**
- Read-only + `create_if_missing` creates the file at the real path.
- Read-write + `create_if_missing` creates a blank workbook at the scratch path.
- Creates from a template workbook when `template:` is given.

**TestCloseAll**
- Closes every opened session.
- Doesn't raise with no sessions opened.
- One failing close doesn't prevent others from closing.

**TestCheckpoint**
- Saves a dirty staged session to its scratch file.
- Never touches the real path.
- Clears the dirty flag after checkpointing.
- Skips a non-dirty session; skips read-only sessions entirely.

**TestCommitAll**
- Saves dirty staged sessions and commits them to the real path.
- Read-only sessions are never touched by commit.

## Unit tests — CLI (`tests/unit/test_cli.py`)

- A successful run returns exit code 0 and prints the `RunResult` as JSON.
- A run with a failed step returns exit code 1.
- `--env KEY=VALUE` (repeatable) is parsed into `env_overrides`.
- A `ValidationError` returns exit code 1 and prints a structured error, not a traceback.

## Unit tests — owned Excel instance registry (`tests/unit/test_owned_instance_registry.py`)

Live-Excel process lifecycle safety (never touching an Excel instance this process didn't
spawn itself).

**TestSpawn**
- `spawn()` returns a new app and tracks its PID.
- Never reuses an existing instance — always spawns fresh.

**TestCloseOwned**
- Quits every owned instance and clears tracking afterward.
- Calling with nothing spawned doesn't raise.
- One instance failing to close doesn't prevent others from closing.
