# excel_runner — Progress Tracker

## Last Session (2026-08-20d, Windows)
**Status:** Ready for Next Phase
**Working on:** Added a `demos/` folder — 7 runnable YAML workflows exercising every existing
action's syntax, for regression testing, plus a `demos/generate_workbooks.py` script that
builds the sample fixture workbooks they read from. Ran all 7 by hand; found and fixed a real
bug (`set_column_width`'s `width` field is typed `float`, but plain YAML ints like `width: 20`
were being rejected — `_matches_type()` in `engine.py` now accepts `int` anywhere `float` is
expected, per PEP 484's numeric tower; `bool` explicitly excluded since it's technically an
`int` subclass but not a sane width). Then, prompted by the demos needing sheet
creation/rename to build believable fixtures, found and fixed a real architectural
inconsistency: whether an action needs its workbook opened read-write was tracked in a
hardcoded `_WRITE_ACTIONS` set in `engine.py`, completely disconnected from the action's own
`@file_action`/`@xlw_action`/`@com_action` registration decorator — a second place to remember
to update by hand whenever a new write action was added. Fixed: the same decorator now takes
`writes=True` (e.g. `@file_action(writes=True)`), stored on `ActionSpec.writes`; `plan()` now
takes the registry and reads `registry[step.action].writes` instead of the hardcoded set.
Added the previously-missing `create_sheet`/`rename_sheet`/`delete_sheet` actions (TDD,
`backends.py` + `actions.py` + unit tests), the actual motivating gap: no action could
add/rename/remove a worksheet at all. Docs (PRD §7/§8, Specification §5.4, README) updated to
match. All 285 tests pass (was 271).
**Next step:** Build a `00_generate_demo_workbooks.yaml` using the new sheet actions (the
originally-requested "yaml-only" fixture generator — was blocked until `create_sheet` existed).
Then resume `_switch_backend` (PRD §6.2.2, Spec §5.2).

**Backlog (not started, explicitly deferred, discussed this session):**
- **Actions-folder-per-file structure**: `discover_actions()` currently scans one module
  (`actions.py`, 365 lines pre-this-session's additions, still under the project's 500-line
  cap). User wants one-file-per-action (`actions/create_sheet.py`, etc.) supported *before the
  repo is "finished"* — `discover_actions()` would need to scan a package's submodules, not
  just one module. Deliberately not done opportunistically this session; do it as its own
  scoped change.
- **Revisit the `__init__.py` approach** — flagged alongside the above, not detailed yet;
  needs its own discussion before changing anything.

## Last Session (2026-08-20c, Windows)
**Status:** Ready for Next Phase
**Working on:** Fixed two CLI usability issues (`cli.py` had no direct-run guard, and `--env`'s
help text didn't explain it's repeatable, not comma/semicolon-separated). Then ran the full
quality-gate suite (pytest, ruff, mypy --strict, pyright, vulture, radon) on Windows for the
first time — all clean, after installing `types-PyYAML`, `types-openpyxl`, and `pyright` as
dev dependencies (none were previously in `pyproject.toml`'s `dev` extra; `mypy --strict` failed
on missing stubs until installed, `pyright` wasn't installed at all).
**Next step:** Commit this session's work (Windows reliability fix, CLI entrypoint, quality-gate
tooling). Then resume `_switch_backend` (PRD §6.2.2, Spec §5.2).

**CLI fixes**: `cli.py` had no `if __name__ == "__main__":` guard of its own (only
`__main__.py` did), so `python excel_runner\cli.py --help` silently did nothing — it imported
the module and defined `main()` without ever calling it. Added the guard directly to `cli.py`
too, since a UiPath xaml step is likely to invoke a script path directly rather than
`python -m excel_runner`. Also clarified the `--env` help text: repeat the flag once per
KEY=VALUE pair (`--env a=1 --env b=2`), not comma/semicolon-separated — not obvious from the
original one-line help text.

**Quality gates, first Windows run**: `mypy --strict excel_runner` failed with
`import-untyped` errors for `yaml` and `openpyxl` (no stubs installed) plus two
`no-any-return` errors in `find_row`/`find_column` that turned out to be a consequence of the
same missing stubs (openpyxl's `Any` types propagating) — installing `types-PyYAML` and
`types-openpyxl` resolved all of it, no source changes needed. `pyright` wasn't installed at
all yet; installed it, 0 errors. `vulture --min-confidence 60 excel_runner
vulture_whitelist.py` and `radon cc --min C excel_runner` were already clean (the whitelist
already covered `OwnedInstanceRegistry` and the `xlw_*` functions from the previous session).
Added `types-PyYAML`, `types-openpyxl`, and `pyright` to `pyproject.toml`'s `dev` extra so
they're installed automatically via `pip install -e ".[dev]"` from now on.

## Last Session (2026-08-20b, Windows)
**Status:** Ready for Next Phase
**Working on:** First Windows session (previously Mac-only). Set up the venv with pip (not
uv) per user's request. Root-caused and fixed a real, reproducible (~50% of runs) intermittent
hang in `OwnedInstanceRegistry.close_owned()` on Windows, added a CLI entrypoint, fixed two
now-stale tests, and fixed a syntax error I introduced mid-session in a previous incomplete edit.
**Next step:** Run the full quality-gate suite (ruff, mypy --strict, pyright, vulture, radon)
on Windows for the first time and confirm they're clean here too — hasn't been checked on this
platform yet, only pytest. Then resume `_switch_backend` (PRD §6.2.2, Spec §5.2), now unblocked.

**Reliability bug: `OwnedInstanceRegistry.close_owned()` intermittently hanging on Windows.**
`app.quit()` on the `xw.App` object returned by `spawn()` would hang indefinitely roughly half
the time, even though the underlying Excel process was confirmed (via `tasklist` and direct
visual observation) to still be a real, quickly-terminating process. Root cause, found by
comparing against an older, previously-reliable reference implementation
(`Risk Demo/src/excel_core.py`, a different project) and testing hypotheses one at a time
against real spawned Excel instances: `spawn()` was using `add_book=False` (a bookless
instance), and a bookless `xw.App` **never registers in `xw.apps` at all** — confirmed directly
by spawning two bookless instances and finding `xw.apps` empty immediately after. Quitting an
instance xlwings itself doesn't know about is what caused the hang. Fix: `spawn()` now uses
`add_book=True` (an initial "Book1" is created and the instance registers immediately), and
`close_owned()` quits via a fresh `xw.apps[pid]` lookup (falling back to the stored object's own
`.quit()` for anything not in `xw.apps`, e.g. the fake in
`test_one_failing_close_does_not_prevent_others_from_closing`). Verified reliable across
repeated spawn/quit cycles (multiple runs, zero hangs) before and after applying to
`backends.py`. Not expected to affect macOS — `add_book`/`xw.apps[pid]` are both part of
xlwings' portable API, not the Windows-only COM layer, and the old reference code used
`add_book=True` as its own cross-platform default.

**Consequence of the fix, caught by the full suite**: `test_backends_xlw.py`'s
`test_closes_the_book_without_quitting_the_app` assumed a fresh spawn has zero books — no longer
true with `add_book=True`. Fixed the test's assumption (checks the *opened* book is gone, not
that the book list is empty), not the new behavior.

**Also fixed**: `test_owned_instance_registry.py`'s `_process_alive` helper used
`os.kill(pid, 0)`, a POSIX idiom that raises `OSError: [WinError 87]` on Windows instead of a
clean liveness signal — first time this suite has run on Windows. Fixed with `ctypes`/
`OpenProcess` on `os.name == "nt"`, stdlib only, no new dependency.

**Self-inflicted bug, also fixed**: an earlier edit this session (adding then abandoning a
timeout+force-kill approach to `close_owned()`) had accidentally deleted the
`class OwnedInstanceRegistry:` line and its docstring's opening — a real `SyntaxError` on
import, caught by re-running the full suite, not by static checks (the broken state hadn't been
imported since the edit). Restored. Also removed the now-unused `subprocess`/`sys`/`threading`
imports left over from the abandoned approach.

**CLI entrypoint added** (`excel_runner/cli.py`, `excel_runner/__main__.py`, `excel-runner`
console script in `pyproject.toml`) — the actual driver for this: the user needs to invoke a
workflow from a UiPath xaml workflow as an external process. `main(argv)` parses a workflow path
and repeatable `--env KEY=VALUE` overrides, calls `run_workflow()`, prints the `RunResult` as
JSON to stdout, and returns 0 (success) or 1 (any step errored, or the workflow itself failed to
load/validate/execute — caught as `ExcelRunnerError` and printed as structured JSON too). 4 new
unit tests (mocking `run_workflow` itself — the zero-mock convention is for the actions/backends
layer, not for testing the CLI's own argument-parsing/formatting responsibility, which is what
these test). A separate `visible` config surface (YAML/CLI flag to show the Excel window) was
discussed but deliberately not built yet — `spawn(visible=...)` already accepts it, but nothing
upstream (`SessionManager`) calls `spawn()` yet, so there's no live call site to wire it to
without adding an unused stub.

All 271 tests pass (was 258 pre-session; +4 CLI, +9 elsewhere from this session's changes),
verified on Windows for the first time. Quality gates beyond pytest not yet re-run on this
platform (see Next step above).


**Status:** Ready for Next Phase
**Working on:** Build order items 1–8 (Spec §8) — the entire v1 file-backend core engine — are
now done, plus the crash-safety integration test. Git repo initialized, nine commits on `main`.
Only item 9 (COM phase, Windows-dependent, later) and item 10 (deferred/flagged items) remain
from the original build order.

**Items 1–6, summarized** (full detail in git log / Spec §2–§5 correction notes): data model +
errors (1); loading/templating, found the whole-file "render then parse" plan couldn't work for
`{{ steps.* }}` (2); action registry + first 5 actions, found `ActionResult`/`WorkbookSession`
needed to live in `core.py` to avoid a circular import (3); 9 more actions taking the total to
15 (miscounted as 14 at the time — corrected in item 8's notes below), established the
`ActionResult(status="error")` vs. raised-exception policy, found
`read_links` empirically broken (4); `SessionManager`/`ScratchManager`, mode caller-specified
pending tier 2, coverage caught a missing-`mkdir` bug (5); both validation tiers, found PRD
§9.1's fourth example message isn't implementable without workbook access — corrected, not
faked (6).

**Item 7 (`runner.py` — orchestration + audit logging, Spec §6.1/§6.2)** — built. Resolved
`copy`'s two-session dispatch and the `workbook`-field-stripping translation deferred since
items 3–4, plus a genuine open design question: an action's `ActionResult(status="error")`
does *not* stop the loop — resolved by what `if:` is actually for (PRD's own
`if: steps.x.status == 'success'` example only makes sense if a failed step doesn't abort the
run before later steps can check it). `RunResult.status` is `"error"` if any step failed, and
that (not loop completion) gates whether anything commits — only a *raised* exception aborts
the loop outright. Surfaced three real bugs, all fixed with regression tests: the audit log was
being written inside the directory `ScratchManager.cleanup()` deletes, so a successful run
deleted its own audit trail; a dict output key literally named `"values"` (`read_range`'s own
PRD §10.4 output shape) was shadowed by Python's real `dict.values()` method under Jinja2's
default attribute resolution, fixed generically (a custom Jinja `Environment` trying item
access first) rather than by renaming around the one collision; and `@file_action`/
`@com_action` were erasing every action's parameter types (`Callable[..., ActionResult]`),
silently defeating mypy at every call site — switched to `ParamSpec`, which then let us catch
and fix `read_metadata` quietly mishandling an unsupported `target` instead of rejecting it.
First genuine end-to-end suite, `tests/integration/test_run_workflow.py`; fixture workbooks
generated in code, not committed as static files (correction to the original `tests/data/`
plan, Spec §7).

**Crash-safety follow-up (user-requested before moving to item 8)** — found a fourth real bug,
the most significant one this session: **the scratch copy left behind after a crash didn't
actually contain any in-progress work.** openpyxl writes stay in memory until an explicit save,
and nothing was flushing them to the scratch file mid-run — so the "recovery artifact" PRD
§6.3.1 promises was, in practice, just an unchanged copy of the original. Fixed with
`SessionManager.checkpoint()`, called after every step in the loop, saving each dirty staged
session to its scratch file as it goes — found by writing the crash test properly (asserting
the scratch file's *content*, not just its existence) and watching it fail on the first
attempt. `TestCrashSafety` now covers: real file untouched, scratch copy genuinely has the
prior step's work, audit log survives, and — the strongest cross-platform proof sessions were
actually closed — a later valid run against the same file just works. "No orphaned Excel
process" stays untested until a COM backend exists (item 9) to spawn one.

**Item 8 (public API surface, Spec §6.3)** — built: `excel_runner/__init__.py` re-exports
(`run_workflow`, `RunResult`, `StepResult`, `Workflow`, `Step`, `WorkbookRef`, `list_actions`,
`ActionSpec`), and `list_actions()` itself (`discover_actions` wired to the real module — not a
second source of truth). Found a real gap: `ActionSpec` never actually had the `description`
field the original design's own words assumed ("name/docstring/param_schema") — without it,
`list_actions()` couldn't fulfill its one stated purpose (a future agent-tool wrapper needs a
description, not just a name and parameter shape). Fixed: `description` is now populated from
each action's docstring, first line only. Also caught, via a test asserting the real action
count: the "14 actions" repeated across items 4–7's notes was an off-by-one — it's 15. Corrected
in Spec §5.1/§6.3 and above; not rewriting git history for it.

All quality gates pass: 239/239 tests, ruff clean, mypy --strict clean (`excel_runner` +
`tests`), radon cc clean, vulture clean, **100% branch coverage** across all 6 modules.
`README.md` written and committed (`4eb5a37`) — every documented example verified end-to-end
against real workbooks, not just asserted.

**Design decision, 2026-08-19**: user flagged that without a way to halt a run early, avoiding a
failed lookup's downstream steps means repeating the same `if:` on every one of them. Agreed
design: a `stop` control-flow action, driven by the existing `if:` mechanism (no new per-step
field), with a distinct `StepResult(status="stopped")` for every step after it — kept separate
from `"skipped"` so the audit log can tell "this step's own condition skipped it" from "the run
ended before we got here." Written up in PRD §6.9 and Specification.md §4/§6.1/§8 (now build
order item 9, renumbering COM to 10 and deferred items to 11). A related but separate idea —
grouped `if:` blocks spanning multiple steps, to avoid repeating the same condition — was
deliberately **backlogged, not designed**: possible over-engineering, parked in PRD §12 until
real workflows show it's an actual recurring need.

**Item 9 (`stop` control-flow action, PRD §6.9/Spec §4/§6.1/§8)** — built via TDD. New
`@control_action` decorator (capability `"none"`) since `stop` takes no `session` and has no
`workbook:` field at all; joined `copy` in `_SCHEMA_EXEMPT_ACTIONS` for the same reason. New
`StepResult(status="stopped")`, distinct from `"skipped"`, for every step after a triggered
`stop` — kept `RunResult.step_results` at one entry per workflow step, and audit-logs the
stopped steps too. `stop` doesn't set `any_failed` on its own, so a deliberate early exit still
commits prior work while "not found → stop" still discards (the failed lookup already set
that). No new commit logic needed — only the loop's early-exit needed writing. Action count is
now **16** (was 15) — `list_actions()`'s count assertion caught every place that needed updating.
All quality gates pass: 248/248 tests, 100% branch coverage, ruff/mypy --strict/radon/vulture
clean.

**Item 10 (xlwings / live-Excel phase) started — `OwnedInstanceRegistry` (backends.py §3.1)
built and tested for real** against a locally-spawned Excel instance (Excel is installed on this
dev Mac, so real testing is possible here for what actually works). Along the way: fixed an unrelated
but serious environment bug — this project's venv lived under `~/Documents` (iCloud-synced),
which was silently breaking `.pth`-file processing (editable installs, `appscript`) via macOS's
hidden-file flag being repeatedly reapplied by the iCloud daemon; fixed by relocating the venv
to `~/.venvs/excel-runner` (symlinked back as `.venv`) and documented in the dev-tree-wide
CLAUDE.md so every future project in this tree gets it right from the start. Confirmed
empirically: open/read/write cell values work reliably via
xlwings on macOS; `save()` does not (`Parameter error -50`, reproduced consistently, matches
known xlwings GitHub issues) — real write-path testing needs the Windows environment the
user will provide. Two Mac-specific `quit()` behaviors found and documented in Spec §3.1 (async
termination; inconsistent error-vs-no-op on a redundant quit) — neither is a bug in the class.

**Naming correction, same session**: the original two-tier "COM" naming (file-backend
unprefixed, `com_`-prefixed for everything live-Excel) was inaccurate — there is no COM on
macOS at all, only Apple Events; xlwings abstracts the difference (PRD §4), which is exactly why
it was chosen over raw `win32com`. Corrected to three tiers throughout code and docs: `file`
(openpyxl), `xlw` (xlwings' portable API — the normal case), `com` (the genuine exception: raw,
Windows-only COM object via xlwings' `.api`, e.g. `recalculate`'s full/full_rebuild modes).
`ACTION_CAPABILITIES`/`ActionSpec.capability` widened accordingly; `com_action` now means the
raw-COM case specifically, with a new `xlw_action` covering what it used to mean;
`WorkbookSession.backend` is `Literal["file", "xlw"]`. `promote_to_xlw` (was `promote_to_com`)
renamed to match. All gates re-verified green after the rename (254/254 tests).

**`xlw_open_workbook`/`xlw_close_workbook`/`xlw_save_workbook` built** (backends.py), mirroring
the file-backend's original open/save/close first slice. Tested for real: open (existing file +
`FileNotFoundError` on a missing one) and close both verified working reliably on macOS;
save gated behind a new `requires_working_xlwings_save` marker (`tests/unit/conftest.py`,
Windows-only for now) since it's confirmed broken here — ruled out the "in-place vs. save-as"
distinction as the cause along the way (tested both explicitly; both fail identically). 257
tests passing, 1 skipped (the gated save test), `backends.py` at 99% branch coverage (the one
gap is exactly that skipped line, expected to close on Windows). All other gates clean.
**User asked to pause spawning real Excel processes for now** (permission-dialog fatigue) — no
more live-Excel test runs until told otherwise; anything needing real Excel verification should
wait.

**Design correction, 2026-08-20 — bidirectional backend switching (PRD §6.2.2).** User flagged a
real gap: a workflow can legitimately need to alternate one workbook between file-backend and
live-Excel operations more than once in a run (write → recalculate, which needs live Excel since
openpyxl can't evaluate formulas → write more, reading the values recalculation just produced).
The originally-sketched `promote_to_xlw` only anticipated a one-way upgrade (file → xlw, stay
there) — insufficient here. Corrected design: `SessionManager.get_or_open` gains a `capability`
param; whenever a step's capability doesn't match its workbook session's current backend, a new
`_switch_backend` does save-if-dirty → close → reopen on the other backend, at the same scratch
path, entirely automatic and invisible to the workflow author (matches PRD §6.1's "backend
selection is never exposed to the author," extended to switching too). Reuses
`OwnedInstanceRegistry` as-is (one shared App per run) and fits inside the existing
scratch-copy/checkpoint model with no new crash-safety mechanism needed. Written up in PRD
§6.2.2 and Specification.md §5.2 — design only, not built yet.

**Capability-param threading built** (the first of the two "no Excel yet" pieces above).
`SessionManager.get_or_open` gained a `capability` param, threaded through from `runner.py`'s
`_dispatch` (which already looked up the registry entry to find `fn`, so passing `.capability`
too was the only change needed there). New `_needed_backend(capability)` maps capability to the
backend it needs (`file`→`file`; `xlw`/`com`→`xlw` — `com` reaches deeper via xlwings' `.api` on
an xlw-backed session, no distinct backend state needed). `get_or_open` now raises a clear
`ActionExecutionError` — not a stub, not a silent wrong-backend return — whenever a session's
current backend doesn't match what the capability needs, since the actual switch
(`_switch_backend`: save-if-dirty → close → reopen) isn't built yet. Every action built so far
is `file`-capability, so this boundary is never hit in real usage today — the full existing
suite (258 tests, non-Excel portion) passes unchanged. 8 new tests, all real (no mocks),
covering `_needed_backend`'s full mapping and `get_or_open`'s match/mismatch behavior on both a
brand-new and an already-open session. Also fixed a stale docstring found along the way:
`SessionManager`'s class docstring still said `promote_to_com`/"COM phase (PRD sec 8, Spec sec 8
item 9)" — a leftover from before both the com/xlw rename and the item-9→10 renumbering, missed
by the earlier sweeps since it's source-code, not a docs file.

**Next step:** Build `_switch_backend` itself (PRD §6.2.2, Spec §5.2) — the actual
save-if-dirty → close → reopen mechanism, plus `SessionManager` holding an `OwnedInstanceRegistry`
per run. This is the second "no Excel yet" piece, and it's the point where that stops being
possible — verifying a real switch needs a live Excel instance. Hold here until the user lifts
the Excel-spawning pause.
**Notes:** `aggregate` and `update_summary_table`'s exact parameters are still explicitly
flagged as open in the PRD — don't block on them. Five things parked in PRD §12, none designed
or scheduled: grouped `if:` blocks; a "replay nice" desktop-comfort mode (visible Excel replay
of an already-finished real run, via xlwings — deliberately not a second execution backend);
openpyxl silently dropping charts it can't parse on any save (a real, evidence-backed risk to
the core save path, not just a feature idea — confirmed via openpyxl's own reader-warning
mechanism and a local round-trip test; verifying the real-Excel-chart case still needs a fixture
we don't have, same blocker as `read_links`); a conversational agent-driven
spreadsheet-authoring product, explicitly flagged as likely a different product from this one;
and a defensive caution to watch for xlwings object references (`Sheet`/`Range`/`Book`) going
stale if held across calls once the `xlw_`/`com_`-tier actions are actually built and tested —
not a confirmed issue here, current design already re-resolves per call rather than caching.
Tracker below stays function/class-granular even though source files are consolidated — see
Spec §7.

## Status legend
❌ Not Done · 🟡 In Progress · ✅ Done — Results: ✅ Pass · ❌ Fail · ⏭️ N/A

## Phase 1 — `core.py`: data model, errors, templating (Spec §2)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Data model dataclasses (`WorkbookRef`, `Step`, `Workflow`) — §2.1 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Error types (`ErrorDetail`, exception classes) — §2.3 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Loading pipeline (`load`) — §2.2 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Templating (`resolve_value`, `evaluate_condition`) — §2.2 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |

## Phase 2 — Registry + first action slice (Spec §5.1, §4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Action registry (`ActionSpec`, `discover_actions`) — `engine.py` §5.1 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `open` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `save` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `close` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `read_range` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_cell` action — `actions.py` | ✅ | ✅ | ❌ | ✅ | ❌ |
| `WorkbookSession`/`ActionResult` (moved to `core.py`) | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| File-backend primitives for the 5 actions above — `backends.py` §3 | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 3 — Remaining v1 file-backend actions (Spec §3, §4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| File-backend primitives for this batch — `backends.py` §3 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `copy` action (two-session signature — see notes above) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_range` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_row` action (base + positional modes) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `write_row` by-header mode | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_table` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `insert_range` action (whole-row/column only) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `insert_range` partial-range support | ❌ | ❌ | ❌ | ❌ | ❌ |
| `set_column_width` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_headers_row` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_row` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_column` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `find_columns` action | ✅ | ✅ | ❌ | ✅ | ❌ |
| `aggregate` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_metadata` action (properties/cells) | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 4 — Execution model (Spec §5.2, §5.3)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| `SessionManager` — `engine.py` §5.2 (bidirectional backend switching excluded — needs item 10, PRD §6.2.2) | ✅ | ✅ | ❌ | ✅ | ❌ |
| `ScratchManager` — `engine.py` §5.3 | ✅ | ✅ | ❌ | ✅ | ❌ |
| `backends.create_workbook` (supports `create_if_missing`) | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 5 — Validation (Spec §5.4)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Tier 1: static schema validation — `engine.py` §5.4 (real-defined-names check excluded, PRD §12) | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| Tier 2: dry-run / step-graph validation — `engine.py` §5.4 | ✅ | ✅ | ❌ | ✅ | ❌ |

## Phase 6 — Runner, audit, public API (Spec §6)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| `AuditLogger` — `runner.py` §6.2 | ✅ | ✅ | ✅ | ✅ | ✅ |
| `run_workflow` orchestration — `runner.py` §6.1 | ⏭️ | ✅ | ✅ | ⏭️ | ✅ |
| `SessionManager.checkpoint()` — `engine.py` §5.2 | ✅ | ✅ | ✅ | ✅ | ✅ |
| Public API surface (`__init__.py` re-exports, `list_actions()`, `ActionSpec.description`) — `runner.py` §6.3 | ✅ | ✅ | ✅ | ✅ | ✅ |
| Crash-safety (mid-run interruption) integration test (Spec §7) | ⏭️ | ⏭️ | ✅ | ⏭️ | ✅ |
| `stop` control-flow action (`actions.py`) + `"stopped"` status + loop early-exit — `runner.py` §6.1, PRD §6.9 | ✅ | ✅ | ✅ | ✅ | ✅ |

## Phase 7 — xlwings / live-Excel (Windows-dependent, later phase per PRD §8)

| Component | Unit Tests | Code | Integration Tests | Unit Results | Integration Results |
|---|---|---|---|---|---|
| Remaining `xlw_`/`com_`-tier backend primitives — `backends.py` §3 | ❌ | ❌ | ❌ | ❌ | ❌ |
| `OwnedInstanceRegistry` — `backends.py` §3.1 | ✅ | ✅ | ⏭️ | ✅ | ⏭️ |
| `recalculate` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `run_macro` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `refresh_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `write_links` action | ❌ | ❌ | ❌ | ❌ | ❌ |
| `read_metadata` action (textbox sub-case) | ❌ | ❌ | ❌ | ❌ | ❌ |

## Backlog / deferred (per PRD §8 and §12)

| Item | Status |
|---|---|
| `update_summary_table` — exact parameters | Not designed yet, deliberately deferred |
| `aggregate`, `write_table`, `write_row` by-header mode | Need step-output context — blocked on `runner.py` (build order item 7), not just deferred by choice |
| `read_links` (+ already-deferred `write_links`) | Blocked on a real Excel-generated fixture (or manual XML/zip surgery) — empirically confirmed openpyxl can't create the relationship itself |
| `export_pdf` | Backlog |
| AI-authoring inspection actions (`list_sheets`, `describe_sheet`) | Planned, next phase after core engine works (PRD §9) |
| Instance-ownership tagging mechanism across crashes (PRD §6.2.1/§12) | Open question |
| Scratch-directory collision avoidance for concurrent runs (PRD §6.3.1/§12) | Open question |
| CLI / MCP wrapper | Deferred (PRD §3/§5) |
