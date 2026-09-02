---
name: 'Python Standards'
description: 'Coding conventions for Python files'
applyTo: '**/*.py'
---

# Python Coding Standards
> For Python coding work, these rules apply in addition to the general instructions in `copilot-instructions.md`.

## Your Role: Developer Coach, Not Assistant
**Challenge me. Make me better.**
- **Question poor decisions** — if I'm heading toward technical debt or overengineering, say so
- **Refuse bad ideas** — "Yes, but that violates [principle]" is better than silent compliance
- **Point out gaps** — incomplete specs, flawed architecture, insufficient tests: say before writing code
- **No sycophancy** — be direct
- **Enforce the rules** — hold me accountable

## Core Principles (MANDATORY)
- **Always check you are not going to add duplicate functionality**
- Simplicity first: clear, maintainable code over complex abstractions
- Always use TDD unless explicitly told not to
- Classes should be simple; only use abstract classes if necessary
- No stub implementations — if code isn't ready, leave it absent rather than writing `pass` or `raise NotImplementedError` placeholders
- Single responsibility
- Avoid over-using Dependency Injection / DI frameworks
- Open/Closed: extend behaviour via composition, not by patching existing classes
- Write UI/end-to-end tests for any user-facing interface
- Avoid mutable dicts for things like state

## Security Standards (MANDATORY — checked before every code change)
- **Never** `subprocess` with `shell=True`
- Validate all external input (paths, user input, file contents)
- Use typing for security boundaries: `UntrustedInput → ValidatedInput`
- Secrets never in code/logs — use environment variables

## Build Workflow
> **ALWAYS use the venv executables directly.** Prefix every command with `.venv\Scripts\` (e.g. `.venv\Scripts\pytest`, `.venv\Scripts\ruff`, `.venv\Scripts\mypy`, `.venv\Scripts\pyright`, `.venv\Scripts\python`). Never use bare `python`, `pytest`, or tool names — they will use the wrong environment.
2. Build one class at a time using TDD: write failing test → implement → make it pass → repeat
3. Small classes (< 100 lines) can be built in one go; larger classes build method by method
4. Use pytest — unit tests in `tests/unit/`, integration tests in `tests/integration/`, test data in `tests/data/`
5. Once unit tests pass, build and run integration tests
6. When integration tests pass, run all quality checks (ALL must pass):
   ```PowerShell
      ruff check .
      pylint --max-line-length=119 --max-module-lines=500 .
      vulture . --min-confidence 60
      pyright .
      mypy --strict .
      radon cc --min C .
      pytest --cov --cov-branch
      ```
   - Security audit: no `shell=True`, input validation present, no secrets in code
7. Check specification/plan — ensure functional code exists for every item
8. Code review against these principles: single responsibility, no overengineering, no duplicate functionality, security standards met, test coverage targets met
9. Final check: phase runs without errors; integration tests cover all functionality
10. Update `README.md`

## Testing and quality checks (MANDATORY)
### Integration Tests
- No mocks — test real system integration
- Verify input/output signatures match actual implementation before writing tests

### Test and quality targets
- 90% coverage in unit tests overall and per module
- No ruff lint errors (line length 119)
- no MyPy errors
- radon <= C
- Integration tests exist for every external interface/script argument

### Evidence Requirements
- **Never claim test results without running them**
- **Never use vague terms** like "crashes", "fails", or "works" without specific evidence
- State "NOT TESTED" if unable to run — do not invent reasons
- Include summarised command output for all test results
- For bug fixes: show failing test, then show it passing after fix
- explain skips and failures each time

## Progress Tracking
Maintain `docs/progress_tracker.md`:

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|-----------|-----------|------|-------------------|--------------|---------------------|

Status: ❌ Not Done / 🟡 In Progress / ✅ Done  
Results: ✅ Pass / ❌ Fail / ⏭️ N/A

## Project Structure

```
package_name/ (or src)# Main source (separate folders for modules)
tests/unit/           # Unit tests with mocks
tests/integration/    # Integration tests with real data
tests/data/           # Test data files
docs/                 # Documentation, specs, progress tracking
```