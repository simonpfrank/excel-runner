# Quality Gate Report

Date: 2026-09-15

## Commands Run

```powershell
& .venv\Scripts\ruff check .
& .venv\Scripts\pylint excel_runner tests vulture_whitelist.py
& .venv\Scripts\vulture . --min-confidence 60
& .venv\Scripts\pyright .
& .venv\Scripts\mypy --strict .
& .venv\Scripts\radon cc --min C .
& .venv\Scripts\pytest --cov --cov-branch
```

## Results

| Check | Result | Details |
| --- | --- | --- |
| `ruff check .` | Pass | No findings. |
| `pylint excel_runner tests vulture_whitelist.py` | Not run | Pylint was not installed in the virtual environment when this report was created. |
| `vulture . --min-confidence 60` | Fail | Exits `3` because it scans `.venv` and reports extensive third-party unused symbols. It also reports five repository/test findings. |
| `pyright .` | Fail | `76` errors and `41` warnings. The environment cannot resolve installed `pytest`, `xlwings`, and `openpyxl` sources; several concrete test typing issues are also reported. |
| `mypy --strict .` | Incomplete | The literal command recursively scans `.venv` and stalls. The equivalent repository-scoped check, `mypy --strict excel_runner tests vulture_whitelist.py`, completes with `74` errors in `15` files. |
| `radon cc --min C .` | Pass | The repository-scoped check reports only grade `C` functions and methods, satisfying the target of no worse than `C`. |
| `pytest --cov --cov-branch` | Incomplete / fail | Collected `601` tests and confirmed three failures, then stalled in `tests/unit/test_backends.py` before totals and coverage were emitted. |

## Confirmed Test Failures

### `tests/integration/test_run_workflow.py::TestStop::test_stopped_steps_still_get_an_audit_record`

The test fails at `tests/integration/test_run_workflow.py:713`:

```text
KeyError: 'step_id'
```

The test assumes every JSON line in the captured audit output includes a `step_id`. At least one
record does not, which means the audit output contains a non-step record or mixed record format
that the test does not filter.

### `tests/unit/actions/test_additional_functions.py::test_write_actions_accept_a_defined_name`

The test fails at `tests/unit/actions/test_additional_functions.py:82`:

```text
KeyError: 'Worksheet hello does not exist.'
```

The call `write_cell(session, "BASIS", "TargetCell", "hello")` reaches the backend with
`hello` interpreted as the worksheet name. The test's `write_cell` call contract does not match
the function or backend parameter order.

### `tests/unit/actions/test_additional_functions.py::test_write_cell_rejects_a_multi_cell_defined_name`

The test fails at `tests/unit/actions/test_additional_functions.py:93` with the same error:

```text
KeyError: 'Worksheet hello does not exist.'
```

The test is intended to validate rejection of a multi-cell defined name, but it does not reach
that validation because the argument-order mismatch makes the backend look for a worksheet named
`hello`.

## Test-Run Blocker

The full suite progressed through the openpyxl-only tests in `tests/unit/test_backends.py`, then
stalled at the first live-Excel test:

```text
TestOpenWorkbookForFormulaRead::
test_reads_the_formula_text_even_when_a_real_cached_value_exists
```

The `requires_excel` marker checks only that the platform is Windows. It does not verify that
Excel can launch and respond through xlwings/COM. The suite remained blocked for more than two
minutes and was terminated. A rerun excluding `requires_excel` tests also stalled near the same
point. Final pass/fail totals, skipped tests, and the coverage percentage are therefore not
available.

## Vulture Findings Outside `.venv`

- `excel_runner/backends.py`: unused `AskToUpdateLinks` attribute.
- `tests/unit/test_action_types.py`: unused `_example_file_action`, `_example_xlw_action`, and
  `_example_com_action` helpers.
- `tests/unit/test_errors.py`: unreachable `else` branch.
- `tests/unit/test_registry.py`: unused `as_` variable.

## Static Type-Check Summary

The `mypy` errors have these main causes:

- Bare intra-package imports such as `import core` and `import backends` cannot be resolved when
  type checking the package.
- The action registration decorator is untyped, which makes most action functions untyped under
  `--strict`.
- `excel_runner/backends.py` has two return annotations that promise `tuple[str, str]` but may
  return `None` as the first item.
- Several test mocks and patches are typed as `object`, causing invalid `copy2` calls and optional
  worksheet indexing.
- `vulture_whitelist.py` accesses instance-only dataclass fields from class objects, which strict
  mypy rejects.

No source files were modified while these checks ran.