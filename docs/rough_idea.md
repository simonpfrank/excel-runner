# excel_runner — design notes (carried over from Project workspace conversation)
 
Context: these notes came from a discussion in the "Old Mutual / Project" workspace, but the
actual build should happen in a **new, separate repo** dedicated to this tool (working name
`excel_runner`). Use this file as the seed for a proper PRD/spec session in that new repo.
 
## Problem statement
 
Excel automation work (openpyxl/pandas for file access, win32com/xlwings for macros and live
recalculation) is currently built as one-off scripts per project. Nothing is abstracted or
reused — every project is "blackbox code." This is explicitly **not** an RPA tool (no
click-simulation/UI automation) — it's a declarative, code-based automation layer, closer in
spirit to Ansible/GitHub Actions/Kestra but for Excel workbook operations.
 
## Prior art reviewed
 
- **General workflow/orchestration engines** (Kestra, Prefect, Dagster, Airflow, Ansible, Argo
  Workflows) — all good at declarative step sequencing, retries, logging — but none has any
  Excel-specific vocabulary. You'd still hand-write a Python task that calls
  openpyxl/xlwings internally, so no real abstraction gain for this specific problem.
- **RPA tools** (UiPath, Power Automate Desktop) — explicitly out of scope, GUI/click-driven.
- **xlwings** has config/CLI scaffolding but nothing like declarative per-operation YAML.
- Conclusion: no existing "Ansible-for-Excel" tool exists. Genuine gap, worth building.
 
## Existing attempt: "Risk Demo" repo (lessons learned)
 
Location referenced in discussion: `C:\Users\simon2282\OneDrive - Willis Towers Watson\Documents\Dev\projects\Risk Demo`
 
What it already got right (keep these ideas):
- A declarative **step table** driving execution (Process name / Step reference / Step
  description / Source workbook / Source worksheet / Source range / Source property / Action /
  Target workbook / Target worksheet / Target range / Target property), authored as a sheet in
  an Excel "configuration workbook" rather than a script per project.
- Supporting concepts worth reusing:
  - `AuditManager` — structured run/audit log of what happened each step.
  - `EnhancedStorage` — lets later steps reference earlier step outputs via a
    `{step.result}`-style template string.
  - `ConfigurationManager` — environment-specific settings (DEV/UAT/PROD), input/output folders.
  - Class layering idea: app/session lifecycle vs. business actions vs. data processing
    (`ExcelCore` → `ExcelAdvanced` → `ProcessCoordinator`, plus `ExcelOperations`,
    `DataProcessor`).
- ~18 supported actions across categories: basic workbook ops (open/save/close), data movement
  (copy, get, output, update_row_cells), lookups (find_headers_row, find_row), advanced
  (aggregate_data, update_summary_table).
 
Where it went wrong (root causes of "clunkiness"):
1. **Everything routed through xlwings/COM**, even pure data moves (`copy`, `get`, `save`, even
   `open`). `ExcelCore.__init__` always spins up a live `xw.App`. This is slow, requires a real
   Excel process/license on the runtime machine, is fragile (dialogs/alerts, orphaned
   `EXCEL.EXE` processes — hence manual `psutil`/PID tracking), and can't run headless in CI.
   No use was made of openpyxl for the file-level operations that don't need a live Excel
   session.
2. **Config encoded as a bespoke string mini-language inside Excel cells**, e.g.
   `"metadata:A1,B2;search_value:Test123"` or `"O8=FAIL"`. This required hand-rolled regex
   parsers (`_is_contiguous_2d_range`, `_get_worksheet_column_letters`,
   `parse_property_parameters`) instead of structured, schema-validated fields. Much harder to
   validate up front, easy to typo, no editor tooling/autocomplete possible.
3. **Defensive "safe fallback" layers everywhere** (e.g. `_safe_header_detection` with three
   nested try/except recovery paths that silently invent fallback headers like `Col_1`,
   `Col_2`). Symptom of not trusting the string-DSL inputs — the system compensates for bad
   input at runtime instead of catching it before execution. This kind of defensive creep is
   exactly the complexity we want to avoid this time — fail fast with a clear validation error
   instead.
4. **No per-action engine selection** — every action pays the COM tax even when it's a pure file
   read/write.
 
## Design decisions for the new tool
 
### 1. Config format: YAML modeled on GitHub Actions conventions
 
Chosen over Ansible/Kestra-style specifically because GitHub Actions YAML is the most common
pattern any LLM (and most engineers) has seen — matters a lot given the intended AI-assisted
authoring workflow (fewer hallucinated/invalid fields, easier schema/autocomplete tooling).
 
Borrowed conventions:
- `steps:` list, each with an `id`.
- `${{ steps.<id>.output }}` style templating to reference earlier step results (replaces the
  Risk Demo `{step.result}` string templating with a familiar syntax).
- `if:` conditional step execution.
- Top-level `env:` block for global settings (replaces the separate "Settings" sheet in Risk
  Demo's config workbook).
- A `workbooks:` registry block at the top of the file, giving each workbook a short logical
  name referenced by steps (avoids repeating filenames/paths everywhere, reduces typos, natural
  home for any future per-workbook overrides).
 
Sketch discussed:
 
```yaml
name: Risk Backtesting Load
env:
  input_folder: "./input"
  output_folder: "./output"
 
workbooks:
  historical:
    file: "2025 Historical analysis v1.xlsx"
  manip:
    file: "Backtesting Manip.xlsx"
    create_if_missing: true
 
steps:
  - id: open_hist
    action: open
    workbook: historical
 
  - id: copy_data
    action: copy
    source: { workbook: historical, sheet: "Reserving Data", range: "A:AC" }
    target: { workbook: manip, sheet: "Reserving Data", range: "A1" }
 
  - id: refresh
    action: run_macro
    workbook: manip
    with: { name: "RefreshModel" }
 
  - id: get_results
    action: read_range
    workbook: manip
    source: { sheet: "Outputs", range: "A1:D50" }
 
  - id: save_manip
    action: save
    workbook: manip
    if: "${{ steps.refresh.status == 'success' }}"
```
 
Still to decide in the new repo's spec phase: full schema field-by-field, naming of every
action, how `workbooks:` interacts with multiple environments, error/status object shape
exposed to `${{ }}` expressions, whether expressions need a mini evaluator or can stay
restricted to simple dotted lookups.
 
### 2. Hiding the file-access vs. xlwings split from the user
 
Users should never have to write an `engine:` field. Each **action** declares its own required
capability, and the runner decides the backend transparently per workbook:
 
```python
ACTIONS = {
    "open":        Capability.FILE,       # openpyxl / just track path
    "copy":        Capability.FILE,       # openpyxl read+write
    "read_range":  Capability.FILE,
    "write_cells": Capability.FILE,
    "save":        Capability.FILE,       # unless workbook has a pending COM session
    "run_macro":   Capability.COM,        # must be xlwings
    "recalculate": Capability.COM,
    "export_pdf":  Capability.COM,
}
```
 
Runtime rule:
- Every workbook defaults to a **file-backed session** (openpyxl, kept open in memory as the
  runner progresses through steps).
- The moment a step needs `Capability.COM` for a given workbook, the runner **promotes** that
  workbook to a live xlwings session transparently: flush openpyxl-held state to disk, open via
  xlwings, run the COM step(s).
- Subsequent `Capability.FILE` steps against that same workbook re-read via openpyxl from the
  xlwings-saved file (or a simple lookahead keeps it on xlwings if more COM steps are queued
  next, to avoid save/reopen thrashing).
- At `save`/`close`, whichever backend currently holds the workbook does the final write.
 
So `copy` is always file access under the hood; only actions that inherently need Excel's
calc engine or VBA (`run_macro`, `recalculate`, pivot/chart refresh, `export_pdf`) trigger COM —
and that's driven by the action's declared capability, never a user choice.
 
### 3. Proposed module architecture
 
```
excel_runner/
  schema.py          # pydantic models for the YAML (Workflow, Step, WorkbookRef, ...)
  loader.py           # parse + validate YAML -> Workflow object, resolve ${{ }} refs
  registry.py         # ACTIONS dict: name -> (capability, handler fn, param model)
  session.py          # WorkbookSession: wraps openpyxl or xlwings handle, handles
                       #   promotion/demotion between backends
  backends/
    file_backend.py    # openpyxl implementations of copy/read/write/open/save
    com_backend.py      # xlwings implementations of run_macro/recalculate/export
  actions/
    basic.py            # open/save/close (dispatch to session)
    data.py              # copy/read_range/write_cells/aggregate
    macro.py             # run_macro/recalculate
  runner.py           # Workflow -> executes steps in order, handles if/needs, retries
  audit.py            # from Risk Demo's AuditManager idea — structured run log (json/csv)
  storage.py          # from Risk Demo's EnhancedStorage idea — step-output store +
                       #   ${{ }} resolution
  validator.py         # static validation + dry-run (see below)
  cli.py              # `excel-runner run workflow.yaml`, `--dry-run`, `--validate`
```
 
Reuses the genuinely good ideas from Risk Demo (audit log, step-output storage/templating) but
separates three concerns that were tangled together in `ExcelCore`/`ExcelOperations`/
`ProcessCoordinator`: **backend** (how — openpyxl vs xlwings), **action** (what — copy, save,
run_macro...), **session** (which backend is currently active per workbook, and how/when it
gets promoted). Also removes the string-DSL parsing entirely in favour of structured,
pydantic-validated fields.
 
### 4. Authoring workflow: AI-assisted YAML + validation + dry-run
 
Three tiers of feedback, increasing cost/confidence, to catch AI-authoring mistakes before a
real run:
 
1. **Schema validation (instant, free)** — pydantic models double as JSON Schema for
   editor/AI-tool autocomplete and inline errors. Catches unknown actions, missing required
   fields, malformed ranges, `${{ }}` references to step ids that don't exist, invalid
   capability combinations.
2. **Static/logical dry-run (no Excel, seconds)** — walk the step graph without touching real
   files: confirm referenced workbooks are declared, sheet/range syntax is valid, `${{ }}`
   references only point at steps that run earlier, no circular `if`/dependency chains. Produces
   a plain-English execution plan (e.g. "step 3 will copy Reserving Data!A:AC from historical →
   manip") for the user to sanity-check.
3. **Live dry-run (real files, safe mode)** — actually opens real workbooks (openpyxl for file
   steps; xlwings only if COM steps are reached) but redirects all writes/saves to scratch
   copies, and runs macros against copies rather than originals. Surfaces real Excel errors (bad
   macro name, broken range, etc.) without risking real data. Slower — run on demand, not on
   every edit.
 
Intended user loop: write/ask-AI for YAML → live schema validation while typing →
`excel-runner validate file.yaml` (tiers 1+2, instant, tight AI-iteration loop) →
`excel-runner run file.yaml --dry-run` (tier 3, safe but realistic) →
`excel-runner run file.yaml` for real.
 
## Open questions / follow-ups for the PRD & spec session in the new repo
 
- Full action catalog and parameter schema per action (port + rationalize Risk Demo's ~18
  actions rather than reinvent from scratch).
- Exact `${{ }}` expression grammar and how much logic it should support (simple dotted lookups
  vs. a small expression evaluator).
- How environments (DEV/UAT/PROD) and secrets/paths are handled — replacement for Risk Demo's
  `ConfigurationManager`.
- Concurrency/parallel workbook handling — does the runner ever need more than one Excel COM
  instance at a time?
- Packaging/distribution — internal package vs. installable CLI tool, how other projects adopt
  it.
- Testing strategy for the COM backend (hard to unit test without real Excel — consider a
  fixture/sample workbook approach, mirroring Risk Demo's `tests/` folder).
- Whether the "Excel config sheet as YAML source" idea from Risk Demo is worth keeping as an
  optional authoring aid (compile Excel sheet -> YAML) alongside direct YAML editing.
 
 
 