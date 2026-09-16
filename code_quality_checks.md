# Code Quality Checks

Date: 2026-09-16

This document records the non-pytest quality checks for the canonical `excel_runner/` package,
the intentionally limited tool scope, and the remaining findings accepted for this release.
It is an evidence record, not a claim that static analysis replaces execution testing. The final
quality gate remains `pytest --cov --cov-branch` in a supported Windows-and-Excel environment.

## Scope and configured exclusions

The commands intentionally inspect `excel_runner/` rather than the repository recursively. This
excludes virtual environments, generated run artifacts, demos, temporary files, and test code so
the results describe the deployable product code. The compatibility launcher is outside that
directory; it has focused deployment tests and was checked separately with Ruff.

The following exclusions are explicit and version-controlled:

| Tool | Configuration | Reason |
| --- | --- | --- |
| Pyright | `pyrightconfig.json` selects `.venv` | Ensures analysis uses the project dependencies rather than global Python. |
| Ruff | `vulture_whitelist.py`: `B018`, `I001` | A Vulture whitelist deliberately consists of otherwise-unused attribute references. |
| Ruff | `run_excel_runner.py`: `E402`, `I001` | The launcher must add its validated deployment directory to `sys.path` before importing the package. |
| Pylint | Ignore `.venv/` and `excel_runner_runs/` | They are third-party/generated content, not product source. |
| Pylint | Allow `open` and `range` as redefined built-ins | These are established, readable public YAML names. `parse_date` was changed to the clearer `date_format` field instead of retaining `format`. |
| Pylint | Existing disabled style rules | The project documents its public API and complex behavior in module/class documentation; disabling generic docstring, duplicate-code, protected-access, small-class, and module-length metrics avoids noise without suppressing correctness checks. |
| Vulture | `vulture_whitelist.py` | Records known static-analysis false positives for dataclass fields, protocol methods, dynamic action discovery, inspected wrapper metadata, and third-party properties. |

The Vulture whitelist is deliberately tracked in Git. It is supplied as an input to Vulture rather
than hiding candidates globally, so every accepted false positive is visible for review.

## Commands and results

Run these commands from the repository root with the project virtual environment:

```powershell
.venv\Scripts\ruff check excel_runner
.venv\Scripts\pylint excel_runner --output-format=text
.venv\Scripts\vulture excel_runner vulture_whitelist.py --min-confidence 60
.venv\Scripts\pyright excel_runner
.venv\Scripts\mypy --strict excel_runner
.venv\Scripts\radon cc --min C excel_runner
```

| Tool | Result | Release assessment |
| --- | --- | --- |
| Ruff | Pass: 0 findings | Import, syntax, and selected correctness rules are clean. |
| Pyright | Pass: 0 errors, 0 warnings | Strict package import and nullable-value analysis is clean. |
| Radon | Pass: no complexity above C | The most complex routines are measurable and remain within the configured threshold. |
| Pylint | 23 findings; score 9.88/10 | Remaining findings are structural metrics, not local correctness warnings. See below. |
| MyPy | 2 errors | Both are local lambda-inference limits in one action. See below. |
| Vulture | 14 candidates | Ten are explainable dynamic/COM cases; four require future usage/removal review. See below. |

## Clean checks

### Ruff

Ruff passes with no findings. It covers selected syntax, import, upgrade, bugbear, and formatting
rules. Its two scoped exclusions are listed above and do not apply to the package itself.

### Pyright

Pyright passes with no errors or warnings. The project explicitly selects `.venv`, avoiding the
previous false diagnostics caused by a global Python interpreter. The table-reference narrowing
uses `TypeGuard[str]` while preserving valid omitted and templated YAML fields.

### Radon

Radon reports ten grade-C routines and nothing worse. The grade-C routines are action parsing,
table validation, link planning/commit logic, and the main workflow orchestration path. These
routines are intentionally branch-rich because they validate external workflow input or coordinate
workbook state. Splitting them solely to lower a metric would disperse related safeguards and make
the execution path harder to review. Grade C is the project's accepted maximum.

## Staged pytest execution

Pytest and coverage are intentionally separate from the static checks above. The suite is run in
increasing environment cost: non-Excel unit tests, file-backend integration tests, live-Excel unit
tests, live-Excel integration tests, then the full coverage command.

### 1. Non-Excel unit tests

```powershell
.venv\Scripts\pytest tests\unit -m "not skipif" -q
```

Result on 2026-09-16: pass, `487 passed`, `0 failed`, `0 skipped`, and `91 deselected` in
`11.52s`. The deselected tests use the shared live-Excel `skipif` decorators and are deferred to
the live-Excel stages below.

### 2. File-backend integration tests

```powershell
.venv\Scripts\pytest tests\integration -q
```

Result on 2026-09-16: pass, `26 passed`, `0 failed`, `0 skipped`, and `0 deselected` in `7.98s`.
This integration suite executes real YAML workflows against real openpyxl workbooks with no mocks;
it contains no live-Excel/xlwings or COM tests.

### 3. Live-Excel unit tests

```powershell
.venv\Scripts\pytest tests\unit -m skipif -q
```

Result on 2026-09-16: pass, `91 passed`, `0 failed`, `0 skipped`, and `487 deselected` in
`203.67s` (`0:03:23`). The deselected tests are the non-Excel unit stage already reported above.

### 4. Live-Excel integration tests

No separate live-Excel integration tests currently exist. The integration suite above is limited to
the file backend/openpyxl; live-Excel coverage is presently exercised by the marked unit tests.

### 5. Full branch coverage

```powershell
.venv\Scripts\pytest --cov --cov-branch
```

Result on 2026-09-16: pass, `604 passed` in `202.63s` (`0:03:22`), with `94%` total branch
coverage. The per-module coverage is `100%` for `__init__.py`, `0%` for the unexercised
`__main__.py` module entry point, `91%` for `actions.py`, `97%` for `backends.py`, `95%` for
`cli.py`, `99%` for `core.py`, `92%` for `engine.py`, and `96%` for `runner.py`.

## Accepted findings

### Pylint structural metrics

Pylint reports 23 findings, all in its refactor category:

| Finding family | Count | Why it is not changed in this release | Risk and follow-up |
| --- | --- | --- | --- |
| Too many arguments / positional arguments | 12 | The affected action and backend functions mirror explicit YAML fields such as source/target workbook references and table columns. Bundling them only to satisfy Pylint would either degrade the generated YAML schema or add a translation layer. | Low runtime risk; review only when a real action-contract redesign is needed. |
| Too many locals / statements / returns | 7 | These are orchestration, validation, and table-inspection functions. Their branches implement visible validation and transaction behavior. | Medium maintenance risk; refactor only with focused regression coverage and a behavior-preserving design. |
| Too many instance attributes | 3 | The affected classes model live workbook sessions, backend primitives, and execution state. Their fields represent distinct state rather than an unstructured dictionary. | Low runtime risk; extract a cohesive value object only when it removes real complexity. |
| Too many ancestors | 1 | A structured error type intentionally inherits the behavior needed to group and surface related validation/execution errors. | Low runtime risk; changing exception inheritance can alter callers' catch behavior, so defer to a dedicated compatibility review. |

The previously reported low-risk Pylint findings were addressed: text I/O now specifies UTF-8,
private constants use compliant names, `NoneType` checks are idiomatic, and necessary dispatch or
cleanup interface exceptions are scoped and documented. Broad cleanup catches aggregate and
re-raise failures; they do not swallow errors.

### MyPy lambda inference

MyPy reports two `misc` errors in `actions.replace_in_range` (lines 799 and 802): it cannot infer
the types of backend-specific lambdas. The runtime branches are already covered by focused action
tests, and Pyright accepts the same code with no diagnostics.

The release does not add broad `Any` casts or `type: ignore` comments merely to make MyPy green.
That would weaken checking at the action boundary. A future focused change can replace the lambdas
with explicitly typed helpers or protocols. This is low runtime risk and medium static-maintenance
risk.

### Vulture candidates

Vulture is conservative and reports symbols without direct static references. After applying the
tracked whitelist, 14 candidates remain:

| Candidates | Why they remain | Release decision |
| --- | --- | --- |
| `parse_date`, `dump`, `replace_text`, `replace_in_range`, `read_table`, `copy_table_columns`, `update_table_cells`, `replace_table_text` | Actions are registered by `engine.discover_actions()` through introspection, which Vulture cannot trace. | Likely false positives. Do not expand the whitelist without a short source review confirming each registration remains active. |
| `display_alerts`, `AskToUpdateLinks` | Dynamic Excel/COM properties assigned through xlwings. | Likely false positives. Retain the explicit code and review before whitelisting. |
| `copy_range`, `com_update_link`, `checkpoint`, `_check_supported_link_layout` | Vulture found no direct static caller. They may be deferred infrastructure or obsolete code. | Do not whitelist in this release. Confirm active use, add focused coverage, or remove each item in a separate change. |

Vulture candidates are not proof of a runtime defect. Keeping the final four visible is safer than
silencing potentially dead code without an ownership decision.

## Public YAML naming decision

The YAML API intentionally retains `range:` and `action: open` because they are natural Excel
terms. They shadow Python built-ins only within the corresponding implementation function, not
globally. The known `range(...)` collision was repaired with an explicit built-in alias and a
focused regression test. `format:` was renamed to `date_format:` because that improved clarity
without sacrificing the YAML vocabulary.

This is a deliberate usability trade-off: public workflows remain readable and stable, while
Pylint exemptions are narrow and documented. New actions should avoid built-in names whenever an
equally clear domain term exists.

## Release conclusion

The package passes Ruff, Pyright, and the configured Radon threshold. The remaining Pylint, MyPy,
and Vulture results are isolated, understood, and either represent deliberate public/API design
choices or require a future design decision rather than a riskier metric-driven rewrite. They do
not justify changing workflow behavior in this release.

The final full suite passes with `604` tests and `94%` branch coverage. Together with the staged
pytest results above, this completes the automated quality-gate evidence for the supported
Windows-and-Excel environment.