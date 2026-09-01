---
description: Python development guidelines
globs:
alwaysApply: true
---
# Coding Context

This is a coding environment. The personal assistant context from the parent Documents folder does not apply here.
# Development Guidelines

## Core Principles
- **always check you are not going to add duplicate functionality**
- **Simplicity first**: Clear, maintainable code over complex abstractions
- Do not overengineer
- No stubs
- Always use Test Driven Development unless specifically told not to
- Build incrementally with ~60 lines per iteration
- Classes should be simple
- Only use abstract classes if totally necessary
- only use pydantic if absolutely necessary
- The human reader of the code should be able to easily understand the class usage and hierarchy
- Always say what you will be doing next and maintain the active task list in `docs/Progress_Tracker.md`

## Progress Tracking
- **MUST maintain Progress_Tracker.md** in docs/ folder tracking development status
- Use `Progress_Tracker.md` as the persistent task list for current work. Update in-progress items and next actions there instead of relying on any external todo tool.
- Update progress tracker after completing each class/method
- Columns: Component, Unit Tests, Code, Integration Tests, Unit Results, Integration Results
- Status values: ❌ Not Done, 🟡 In Progress, ✅ Done
- Results: ✅ Pass, ❌ Fail, ⏭️ N/A

### Session Handoff
At the END of each work session, update the TOP of Progress_Tracker.md with:

```markdown
## Last Session (YYYY-MM-DD)
**Status:** [In Progress / Blocked / Ready for Next Phase]
**Working on:** [Brief description of current task]
**Next step:** [What to do when resuming]
**Notes:** [Any blockers, decisions needed, or context]
```

This allows any new session to read Progress_Tracker and immediately continue.

## Overall Development Workflow
Unless told otherwise you should ALWAYS follow the following workflow. If the user starts breaking the flow, challenge them.
1. Create PRD with user (in docs folder)
2. Create specification and plan with user (in docs folder) do not work off your own plans, use the content of your own plans in the docs folder for the plan document in use. Include architecture diagram showing file/class breakdown respecting max 500 lines per file constraint.
3. Build class by class, small classes (e.g. < 100 lines, can be built in one go)
4. Larger classes should be built incrementally following TDD for each iteration or chunk
5. Use Test Driven Development to build
6. Use pytest for tests with unit tests in tests/unit/ and integration tests in tests/integration/. Test data for integration tests should be in tests/data/
7. Save a test_summary.md in ./docs with a section for unit and integration tests. document each test with single line bullet point to describe the test so it is easy for the reader to see what tests there are. You can segment by class or functional area with sub headings
8. Once a feature / phase's unit tests pass, build and run integration tests
9. When integration tests pass, run automated quality checks (ALL must pass):
   - pylint/ruff (code quality - no violations)
   - mypy --strict (type safety - 100% coverage)
   - pyright (additional type-checking pass)
   - vulture --min-confidence 60 (dead code detection)
   - radon cc --min C (complexity - no functions rated D or worse)
   - pytest --cov (90%+ branch coverage required)
   - Security audit (no shell=True, input validation present, no secrets in code)
   - Performance benchmarks (startup < 500ms, memory < 50MB baseline)
10. Then you MUST go through the specification/plan for the current phase and double check there is functional code for each item in the phase. If not repeat 5,6,7 until complete unless not possible in which case document in Project_Tracker.md in ./docs
11. Perform a harsh code review on the new code and save in docs
12. Final check that phase can be run without errors, ensure integration tests test all functionality
14. Update README.md
15. Every 3 features/phases perform overall harsh code review to ensure that the new code does not need refactoring

## Post-Refactor/Scope-Change Checklist

After ANY scope change, model deletion, or major refactor, run this before claiming complete:

1. **Search for ALL references to deleted/changed items**
   ```bash
   grep -r "deleted_model_name" . --include="*.py" --include="*.html"
   grep -r "deleted_url_name" . --include="*.py" --include="*.html"
   ```

2. **Fix ALL linter errors - zero tolerance**
   ```bash
   ruff check . --fix
   # If errors remain, FIX THEM. Never skip "pre-existing" errors.
   ```

3. **Manual smoke test - actually use the app**
   - Start the server and click through: login → main flow → key features
   - If forms need data to function, create seed data first

4. **Verify the app runs without errors**
   ```bash
   python manage.py check  # Django
   # or equivalent for other frameworks
   ```

### TDD Methodology
TDD must be followed for new functionality, changes and bug fixing.
1. **Write failing unit tests first** to define expected behavior
2. **Implement code** until unit tests pass
3. **Write integration tests** only after methods exist and unit tests pass
4. **Verify all claims** with actual test output before reporting completion
5. Use mocks for dependencies in unit tests

### Integration Tests (Post-Implementation)
- Write ONLY after methods exist and unit tests pass
- **ZERO mocks in integration tests** — no MagicMock, no @patch, no unittest.mock. If it's in tests/integration/, it runs for real.
- If a test needs an API key, use `@pytest.mark.skipif` — it skips cleanly, it doesn't mock
- Every method call must correspond to existing, working code
- **GATE CHECK: Run `grep -r "MagicMock\|@patch" tests/integration/` after writing — if it returns anything, delete the mock and write a real test**
- Integration tests that pass with mocks but would fail with real calls are worse than no tests — they create false confidence
- Use `dir(class_instance)` to verify methods exist before testing
- Verify input and output signatures match the actual functionality by reading the function implementation

## Evidence Requirements
**ALL test claims and bug fixes must be backed by evidence:**
- **NEVER claim test results without running them** - Always execute tests and show output
- **NEVER use vague terms** like "crashes", "fails", or "works" without specific evidence
- **State "NOT TESTED"** if unable to run tests - don't invent reasons or assume outcomes
- Include summarized command output for all test results in your responses
- Reference specific error messages, stack traces, or log output for failures
- For bug fixes: Show the failing test, then show it passing after the fix
- For new features: Show the test passing with actual output, not hypothetical
- When claiming performance improvements: Include benchmark numbers before/after
- Document any assumptions or limitations discovered during testing


# Python Settings

## Testing & Quality
- Always use pytest for tests
- BDD-style integration tests for all user workflows
- Performance tests: max startup time, max memory usage, max command count
- Cross-platform tests: Linux, macOS, Windows (where applicable)
- Minimum 90% branch coverage (not just line coverage)
- Every bug fix MUST include regression test
- Use pathlib not os.path for cross-platform compatibility
- Avoid hardcoded commands (e.g., ls) - use platform-specific alternatives

### Quality Checks
- use ruff as much as possible
- `pylint --max-line-length=119 --max-module-lines=500` or ruff equivalent
- `radon cc --min C` (complexity checker)
- `mypy --strict` (type checking)
- `pyright` (additional type-checking pass)
- `vulture --min-confidence 60` (dead code detection)

## Important Principles
Try to stick to these without adding complexity. Code should always be easy to read
- Single responsibility
- Avoid over-using Dependency Injection / DI frameworks
- Open/Closed: extend behaviour via composition, not by patching existing classes
- BDD testing for user interaction
- Avoid mutable dicts for things like state

## Code Quality
- Type checking: `mypy` (if configured)
- type hints required
- Line length: 119 characters
- Import order: stdlib → third-party → local (blank line separated)
- Naming: `PascalCase` classes, `snake_case` functions/variables, `UPPER_SNAKE_CASE` constants
- Google-style docstrings with Args/Returns/Raises
- Python logging format: Date, Time, Level, Module, Function, Line, Message — concretely,
  `"%(asctime)s %(levelname)-8s %(module)s %(funcName)s:%(lineno)d %(message)s"` with
  `datefmt="%Y-%m-%d %H:%M:%S"`. WARNING and above go to stderr; DEBUG/INFO go to stdout.
  INFO should narrate what's actually happening in production (each meaningful step: staging/
  committing a workbook, switching backend, opening/saving/recalculating via Excel COM,
  spawning/closing an Excel instance) — enough that someone watching a live run understands
  progress without reading code. DEBUG adds the detail needed to debug a specific failure
  (resolved params, exact paths, link names). Library modules never attach handlers
  themselves (`logging.getLogger(__name__)` only) — only the CLI entrypoint configures
  handlers/formatting.
- Max 500 lines per file
- ~50 lines per function (guideline, not hard limit — 60 is fine if the function is a clear linear sequence; don't split just to hit a number)
- Max 2 levels of nesting
- Max 5 parameters per function (composition roots may need more)
- try to avoid nested closures
- Delete commented code immediately - trust git history
- Never commit placeholder/unimplemented functions
- Only extract helpers when the extraction has a clear name and genuinely improves readability — not to satisfy a line count
- **If a function is getting long, ask whether it's doing too many things. If yes, refactor. If it's just a long linear sequence (e.g. wiring up a composition root), leave it.**

## Documentation Standards (MANDATORY)
- Every public class/function: Google-style docstring with Args/Returns/Raises/Example
- Every module: Module docstring explaining purpose and usage
- Progressive examples required: 01_basic.py through 05_advanced.py
- Architecture diagram in docs/ before coding (Step 2 of workflow)
- README must include "How It Works" section explaining internal architecture

## Type Safety (MANDATORY)
- 100% type hint coverage - mypy --strict must pass
- Use NewType for domain types: UserID = NewType('UserID', str)
- Generic types required: List[str] not list, Dict[str, Any] not dict
- All function signatures fully typed including return types
- No Any types except where genuinely necessary (document why)

## Security Standards (MANDATORY)
- Never subprocess with shell=True
- Validate all external input (paths, user input, file contents)
- Use typing for security boundaries: UntrustedInput → ValidatedInput
- Document security assumptions in docstrings
- Secrets never in code/logs - use environment variables

## Performance Requirements
- Startup < 500ms - measure with time.perf_counter() in tests/
- Memory < 50MB baseline - track with tracemalloc
- Command execution overhead < 100ms (excluding command logic)
- Document performance characteristics in README
- Performance regression tests required for critical paths

## Project Structure
- `package_name/` - Main source code (separate folders for modules)
- `tests/unit/` - Unit tests with mocks
- `tests/integration/` - Integration tests with real data
- `tests/data/` - Test data files for integration tests
- `docs/` - Documentation, specifications, and progress tracking
