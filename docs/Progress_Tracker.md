# excel_runner — Progress Tracker

## Last Session (2026-08-28b, Windows)
**Status:** In Progress — gap-2 (link-aware repointing/recalc, `docs/recalc_and_link_refresh_plan.md`)
**Working on:**
- `discover_write_intent_link_graph(workbook_paths, write_intent)` added to `engine.py`: scans
  every write-intent workbook's real file (via `scan_external_link_targets`), keeps only
  `"absolute"`-classified links (R3/R4), resolves each and matches it against the other
  declared workbooks' real paths — an edge is added only when the resolved target is itself a
  write-intent workbook (R4); a match against a non-write-intent declared workbook (R3) or no
  match at all is ignored. Produces exactly the `{name: set(targets)}` shape
  `compute_link_commit_order()` consumes.
- 6 new tests in `tests/unit/test_link_commit_graph.py`. `scan_external_link_targets` is
  monkeypatched in these tests (deliberate, documented choice) — it's already covered by real-
  Excel tests in `test_link_discovery.py`, and a genuinely absolute (R3/R4) link is hard to
  reproduce portably in a fixture (Excel collapses same-drive links back to relative form, per
  last session's probe finding).
- Full unit suite: 329 passed (up from 323), no regressions. ruff + mypy --strict clean.
- **Important scope finding, raised with and acknowledged by the user**: the actual production
  blocker in `input/workflow.yaml` (config step 9, repointing `Premium Risk Back-testing.xlsx`
  at `Backtesting Manip.xlsx`) is almost certainly R1 (same-folder — both workbooks live under
  the same `input_folder`), not R4. R1 needs no link repointing at all per the plan doc, just
  the `scratch/working/` same-folder co-location mechanism (sec 1.2) — a much smaller change
  than the full R4 dependency-graph/commit-order machinery just built. User explicitly chose to
  finish the R4 graph-wiring work first (this and the prior two sessions' work) before
  switching to the R1/`ScratchManager` rework — this was a deliberate priority decision, not an
  oversight, and should not be second-guessed without asking again.

**Next step:** The R4 side (`compute_link_commit_order`, `discover_write_intent_link_graph`,
`scan_external_link_targets`/`classify_link_target`/`resolve_link_target`) is now feature-
complete as pure/composable functions, but still not wired into any real run — no caller
invokes `discover_write_intent_link_graph` yet. Two remaining threads, in the order the user
chose:
1. Finish R4 end-to-end: wire staging-time `ChangeLink` (real path -> target's scratch path,
   once per R4 link) and commit-time revert-and-save (switch to `xlw` backend if needed,
   `ChangeLink` back to real path, save — no extra recalc call needed per probe10) using
   `compute_link_commit_order()`'s order. Needs a real invocation point — likely wherever
   `ScratchManager`/`SessionManager.plan()` orchestration happens in the runner, since `plan()`
   itself is explicitly no-workbook-access (tier-2, static-only).
2. Then the R1/`ScratchManager` rework (`scratch/working/` + `scratch/originals/` subfolders
   per plan doc sec 1, replacing the current flat `scratch/` dir) — this is what will actually
   unblock `input/workflow.yaml`'s config step 9, once R4 is done.
**Notes:** `unify_storage/` remains an untracked directory in git status — not yet discussed
with the user, not currently blocking anything.

## Last Session (2026-08-28, Windows)
**Status:** In Progress — gap-2 (link-aware repointing/recalc, `docs/recalc_and_link_refresh_plan.md`)
**Working on:**
- Static, pure-Python external-link discovery pass added to `engine.py` — no Excel/COM
  involved, safe to call during planning:
  - `scan_external_link_targets(path)`: reads an xlsx's `xl/externalLinks/_rels/*.rels` via
    plain `zipfile`/`xml.etree`, returns the raw `Target` strings (filtered to
    `TargetMode="External"`). Confirmed empirically (throwaway probe, deleted after use) that
    Excel always serializes a link's Target relative to the workbook's own folder when
    possible (even one typed as an absolute path collapses back to a bare filename or a
    forward-slash subpath if the target is on the same drive) — genuinely absolute/UNC/`file://`
    Targets only occur when a relative form isn't expressible at all.
  - `classify_link_target(target)`: pure string classification into `"same_folder"` (R1),
    `"relative_subpath"` (R2, backlog/unsupported), or `"absolute"` (R3/R4) — bare filename vs.
    contains a separator vs. drive-letter/UNC/`file://`.
  - `resolve_link_target(target, linking_workbook_path)`: resolves any of the three forms to a
    real absolute `Path`, relative ones resolved against the linking workbook's own folder.
  - 13 new tests in `tests/unit/test_link_discovery.py` (10 pure, 3 `@requires_excel` using a
    real Excel-saved fixture only to prove the scan reads real Target strings correctly).
- Full unit suite: 323 passed (up from 310). One unrelated flaky failure seen in
  `test_owned_instance_registry.py` (Excel process teardown timing) — passes in isolation, not
  a regression, pre-existing flakiness.
- ruff + mypy --strict clean on all changed files.

**Next step:** These three functions are still standalone — not yet wired into anything that
builds the `link_targets` graph `compute_link_commit_order()` consumes. Still needed, in order:
1. A function that, given the workflow's real per-workbook paths and `plan()`'s read/write
   modes, scans every declared workbook that exists on disk, resolves each `"absolute"`-
   classified link target against the other declared workbooks' real paths, and — for targets
   that resolve to another *write-intent* declared workbook — builds the `link_targets` graph
   for `compute_link_commit_order()`. Note: `plan()` itself explicitly promises no workbook
   access (tier-2, static-only) — this scanning function is a separate pass, not an addition to
   `plan()`; still deciding where it's actually invoked from (likely at scratch-staging time,
   alongside `ScratchManager`, not tier-2 validation).
2. `ScratchManager` restructuring per plan doc sec 1: `scratch/working/` (original basenames,
   for R1 resolution) + `scratch/originals/` (write-intent pre-edit backups), replacing the
   current single flat `scratch/` dir. Bigger rewrite of `test_scratch.py` too.
3. Staging-time `ChangeLink` (real path -> target's scratch path) wiring, once per run per
   R4 link.
4. Commit-time revert-and-save wiring (per plan doc sec 3): switch to `xlw` backend if needed,
   `ChangeLink` back to the real path, save — no extra recalc call needed (probe10). Must run
   in `compute_link_commit_order()`'s order.
5. Rewrite `ScratchManager.commit()`/`commit_all()`/`_rollback()` to the copy-based (not
   rename-based) semantics from plan doc sec 3, since commit is no longer "just" a rename —
   R4 workbooks need the link-revert-and-save step first.
**Notes:** `unify_storage/` remains an untracked directory in git status — not yet discussed
with the user, not currently blocking anything.

## Last Session (2026-08-27, Windows)
**Status:** In Progress — gap-2 (link-aware repointing/recalc, `docs/recalc_and_link_refresh_plan.md`)
**Working on:**
- Built and committed the raw COM link primitives in `backends.py`: `com_link_sources`,
  `com_change_link`, `com_update_link` (all thin `.api` wrappers, same convention as the
  existing `com_calculate_*` primitives). `_XL_LINK_TYPE_EXCEL_LINKS = 1` module constant.
- Fixed a real hang bug found while testing against live Excel: `xlw_open_workbook` now always
  passes `update_links=False` to `app.books.open()` — without it, opening a workbook with an
  existing external link in a headless spawned Excel instance could trigger an invisible-but-
  modal dialog and hang forever.
- Centralized alert suppression: `OwnedInstanceRegistry.spawn()` now unconditionally sets
  `app.display_alerts = False` and `app.api.AskToUpdateLinks = False` — project-wide decision,
  confirmed with the user, that this tool never wants interactive Excel prompts.
- `TestLinkPrimitives` (4 real-Excel tests) added to `tests/unit/test_backends_xlw.py`, all
  passing. Full unit suite: 302 passed at that point (up from ~298).
- Revised `docs/recalc_and_link_refresh_plan.md`'s R4 to the final three-phase stage/commit-
  once design, added R5 (commit order is a dependency graph), R6 (reject cycles), R7 (reject
  link chains beyond one hop, backlog). Added `_link_probe9.py`/`_link_probe10.py` as committed
  evidence scripts proving the new design (including under manual calculation mode).
- Implemented `compute_link_commit_order()` in `engine.py`: pure topological-sort function
  over a `{workbook_name: set_of_r4_link_targets}` graph, implementing R5 (dependency order),
  R6 (raises `ValidationError` on a 2-node cycle), and R7 (raises `ValidationError` on a chain
  — a workbook that's both an R4 target and has its own outbound link to a *different*
  workbook than what links to it; a clean mutual pair is left to the cycle check instead of
  being misreported as a chain). 8 new tests in `tests/unit/test_link_commit_order.py`, all
  passing. Full unit suite: 310 passed (up from 302), no regressions.
- Confirmed relevance to the real production `input/workflow.yaml`: its "STOPPED HERE" blocker
  (config step 9, repointing `Premium Risk Back-testing.xlsx`'s links at `Backtesting
  Manip.xlsx`) is exactly gap-2. Per the finalized R4 design, this will NOT become an explicit
  new YAML action — repointing happens automatically at staging/commit time based on statically
  detected links, not something a workflow author writes as a step.

**Next step:** `compute_link_commit_order()` is pure and not yet wired to anything — still
needed, in order:
1. A static link-discovery pass (likely a fast zipfile/XML scan of real `.xlsx` files, no
   Excel/COM needed) that finds R1 same-folder siblings and R4 absolute/UNC link targets,
   wired into `plan()`.
2. `ScratchManager` restructuring per plan doc sec 1: `scratch/working/` (original basenames,
   for R1 resolution) + `scratch/originals/` (write-intent pre-edit backups), replacing the
   current single flat `scratch/` dir. This is a bigger rewrite of `test_scratch.py` too.
3. Staging-time `ChangeLink` (real path -> target's scratch path) wiring, once per run per
   R4 link.
4. Commit-time revert-and-save wiring (per plan doc sec 3): switch to `xlw` backend if needed,
   `ChangeLink` back to the real path, save — no extra recalc call needed (probe10). Must run
   in `compute_link_commit_order()`'s order.
5. Rewrite `ScratchManager.commit()`/`commit_all()`/`_rollback()` to the copy-based (not
   rename-based) semantics from plan doc sec 3, since commit is no longer "just" a rename —
   R4 workbooks need the link-revert-and-save step first.
**Notes:** `unify_storage/` remains an untracked directory in git status — not yet discussed
with the user, not currently blocking anything.

## Last Session (2026-08-26, Windows)
**Status:** `recalculate` action built and fully tested — first of the "known gaps" from the
`input/` migration project (see `input/Instructions.md`).
**Working on:** Built via TDD, real Excel throughout (no mocks, project convention):
- **`SessionManager` bidirectional backend switching**: `get_or_open` now switches an
  already-open session's backend in place (file<->xlw, both directions) instead of raising —
  save (if dirty) -> close -> reopen, strictly in that order, to avoid a Windows file-lock
  race. Verified empirically that opening two workbooks via `app.books.open()` on the same
  `xw.App` puts both in the same Excel process, so `SessionManager` now spawns and reuses one
  shared `OwnedInstanceRegistry` App instance for the whole run (not one per workbook) — needed
  for cross-workbook links to resolve/recalculate correctly later. `close_all()` now dispatches
  save/close by backend and also quits the shared owned instance.
- **`backends.py` recalculation primitives**: `xlw_calculate_all` (portable, scope=all/normal),
  `com_calculate_workbook` (loops `Worksheet.Calculate()` over every sheet — `Workbook.Calculate`
  does not exist in the COM object model at all, found by a failing test), `com_calculate_sheet`,
  `com_calculate_full`/`com_calculate_full_rebuild` (always application-wide — no per-workbook
  equivalent), `com_wait_until_calculation_done` (polls `CalculationState`, bounded timeout).
- **`recalculate` action** (`@com_action(writes=True)`): `scope` (`sheet`/`workbook`/`all`,
  default `workbook`), `mode` (`normal`/`full`/`full_rebuild`, default `normal`), `sheet`
  (optional, only for `scope: sheet`, falls back to the active sheet + `output.warning` if
  omitted). Validates `full`/`full_rebuild` require `scope: all`, and `sheet` requires
  `scope: sheet` — both raise `ActionExecutionError` otherwise. Always saves before returning.
- Updated `test_registry.py`/`test_list_actions.py` allowlists, `vulture_whitelist.py`, README
  (status, changelog, action reference, not-yet-available list), and the `excel-runner-yaml`
  SKILL.md action catalog.

298 unit tests pass (up from 293), full quality gate clean (ruff, mypy --strict, pyright on
changed files, vulture, radon — no D+ complexity introduced).

**Next step:** Extend `input/workflow.yaml` with a `recalculate` step for `backtesting_manip`
(config step 6), re-validate (dry-run), re-back-up `input/_backup/` fresh, run for real, verify
output. Then continue to the next gap in `input/Instructions.md`'s "Known gaps" list
(external-link repointing, config step 9) — same discuss -> approve -> TDD -> extend -> test ->
docs -> repeat cycle.

## Last Session (2026-08-21i, Windows)
**Status:** Ready for item 13/14 or item 10's remaining live-Excel actions
**Working on:** User manually verified item 12's console logging by attaching a handler
(`logging.basicConfig()`) — confirmed `INFO`/`DEBUG` records are genuinely emitted, just
invisible by default since neither the library nor the CLI attaches a handler (by design).
Then asked to remove the CLI's `RunResult`-as-JSON stdout print entirely (it was redundant/
noisy now that console logging exists). Flagged the tension first — that JSON print was the
CLI's originally-documented automation contract — but proceeded once the user confirmed: since
`working_dir` is now a fixed, predictable path (item 12), an external caller can read
`working_dir/audit.jsonl` directly instead of parsing stdout, so this isn't actually a
regression for that use case. `cli.py main()` now only returns an exit code; a caught
`ExcelRunnerError` is logged at `ERROR` instead of printed as JSON. Updated `test_cli.py`
(dropped `capsys`-based JSON assertions, added a `caplog`-based one for the error-logging path),
Specification.md §1/§6.4, and the README changelog to match. Also clarified for the user (no
code change) that `.bak` files live next to the *real* workbook, not inside
`working_dir/scratch/`, and are deleted immediately on a fully successful commit — not seeing
one after a normal run is expected, not a bug.

293 tests pass, full quality gate clean.

**Next step:** Item 13 (live-Excel hang safety + configurable timeouts) and item 14
(linked-workbook refresh) both need a live Windows Excel instance to build/verify for real —
now available. Alternatively, resume item 10's remaining live-Excel actions (`recalculate`,
`run_macro`, `refresh_links`, `write_links`, `read_metadata`'s textbox sub-case) —
`read_links`/`write_links` specifically are unblocked now that a real Excel-generated fixture
can be produced to test against.

## Last Session (2026-08-21h, Windows)
**Status:** Build order item 12 (crash/lock-safety hardening) — **done**. Ready for item 13/14
(needs a live Windows Excel instance) or item 10's remaining live-Excel actions.
**Working on:** Built everything decided in the multi-session design thread, TDD throughout:
- **`working_dir` relocation**: `ScratchManager(working_dir)` replaces `tempfile.mkdtemp()`;
  fixed path `<base>/excel_runner_runs/<yaml_stem>/`, `<base>` = cwd or `--working-dir`.
  `run_workflow()` gained a `working_dir` param. The `working_dir:` YAML field originally
  sketched in the PRD wasn't built — no clear use case beyond the CLI flag ever showed up.
- **Read-only staging**: `SessionManager._open_read_only` now calls `scratch.stage(...,
  writes=False)`, same as read-write (`writes=True` default) — `ScratchManager.commit_all()`
  skips any workbook staged with `writes=False`, since nothing about it ever changes.
- **Rename-based commit + rollback**: `ScratchManager.commit()` prepares new content at a
  `.tmp` sibling before touching `real_path` at all, renames the original aside to `.bak`
  (zero-copy), then renames `.tmp` into place. `commit_all()` rolls back every
  already-committed workbook (reverse order) on a later failure, recording per-file whether
  rollback itself succeeded — a file whose rollback also fails is named explicitly in the
  raised error as needing a human, with its `.bak` deliberately kept. The originally-sketched
  `CommitFailure` dataclass wasn't built — `ErrorDetail`'s existing fields already carry
  everything needed as clear text, no new public type needed.
- **`cleanup()` removed entirely** — nothing in `working_dir` is ever auto-deleted now.
  `AuditLogger` truncates `audit.jsonl` on open instead of appending, so a re-run against the
  same fixed `working_dir` never mixes with a previous run's records.
- **Console logging**: `runner.py` gained a module logger; `INFO` on every step's
  start/skip/completion, `DEBUG` on resolved params (logged once in `_dispatch`, not
  duplicated in the main loop), `ERROR` on a failed step. CLI gained `--working-dir` and
  `--logging-level` flags — the latter only sets `logging.getLogger("excel_runner").setLevel(...)`,
  no handler/formatter configuration at all (that's explicitly someone else's job, per the
  user's own correction earlier this thread).
- **Real bug found and fixed mid-build**: every integration test calling `run_workflow()`
  without an explicit `working_dir=` was defaulting to cwd — which during a pytest run *is* the
  actual repo directory, so running the test suite was littering the real project with
  `excel_runner_runs/` test artifacts. Fixed by passing `working_dir=str(tmp_path)` explicitly
  in every integration test; `.gitignore` also covers `/excel_runner_runs/` as a safety net for
  real manual CLI usage inside the repo.
- Confirmed (via a repo-wide check) no leftover internal-tool-name references reappeared in
  Specification.md's new content this session — the earlier redaction pass already covered it.

293 tests pass (was 285 at the start of this build). Full quality gate (pytest --cov,
ruff, mypy --strict, pyright, vulture, radon) clean — 99% branch coverage, all new code fully
covered. Specification.md build order item 12 marked done; §5.3/§6.1/§6.2.1/§6.4 status headers
updated from "not yet built" to "built".

**Next step:** Item 13 (live-Excel hang safety + configurable timeouts) and item 14
(linked-workbook refresh) both need a live Windows Excel instance to build/verify for real —
now available and confirmed. Alternatively, resume item 10's remaining live-Excel actions
(`recalculate`, `run_macro`, `refresh_links`, `write_links`, `read_metadata`'s textbox
sub-case) — `read_links`/`write_links` specifically are now unblocked (see the 2026-08-21g
session note below) since a real Excel-generated fixture can finally be produced to test
against.

## Last Session (2026-08-21g, Windows)
**Status:** Ready to build — item 12 (crash/lock-safety hardening) unblocked; items 10/13/14
now also unblocked (Windows + real Excel confirmed available)
**Working on:** Two things. First, redacted every remaining reference to the user's internal
tool name and automation platform (a few the user's own edits had already caught, several more
found in `docs/Progress_Tracker.md`, `README.md`, and `excel_runner/cli.py`'s own module
docstring) — replaced with generic language ("external orchestration/automation", "a 3rd party
workflow tool"). Repo-wide `git grep` sweep confirms none remain anywhere.

Second: confirmed we're on Windows with a real Excel install available now — this directly
unblocks `read_links`/`write_links`, previously deferred specifically because verifying them
needed a real, Excel-generated fixture workbook with a genuine external link (openpyxl can only
*reflect* an external-link relationship real Excel already created, not create one from
scratch) and no way to generate/verify that without Windows + Excel. PRD updated (§7's catalog
rows, §12's reopened item) to record this as unblocked rather than carried forward.

Also confirmed the "persist `read_links`' results in memory so `write_links` can use them as a
source" requirement is **already fully designed for** — PRD §11 item 18 already shows
`write_links.links: "{{ steps.original_links.output }}"`, replaying a `read_links` step's whole
output directly via the existing step-output/templating mechanism (already in-memory for the
run's duration). No new mechanism needed; `read_links`' output shape and `write_links`' `links`
input shape just need to stay symmetric, which the catalog already specifies. Made this
explicit in the PRD rather than leaving it implicit.

**Next step:** Begin building. Specification.md build order item 12 (working_dir relocation,
read-only staging, rename-based commit/rollback, console logging, CLI flags) has no external
dependency and is fully TDD-buildable now. Items 10/13/14 (live-Excel actions incl.
`read_links`/`write_links`, hang safety, timeouts, linked-workbook refresh) are also unblocked
now that Windows + Excel are confirmed, but need a live Excel instance to build/verify against
directly, rather than reasoning about it in the abstract.

## Last Session (2026-08-21f, Windows)
**Status:** Design complete, spec written — ready to build once reviewed
**Working on:** Wrote Specification.md sections for the entire crash/lock-safety design thread
from this multi-session discussion (PRD §6.2.3/6.2.4/6.3.2/6.3.3/6.3.4/6.7.1, all decided).
Concrete additions:
- **§1**: fixed a stale claim ("no cli.py in v1") — `cli.py`/`__main__.py` added to the layout,
  noted as built for a real reason (an external orchestration/automation driving need), not the
  originally-deferred scope.
- **§3.2/§3.3** (new): `run_with_timeout`'s process-isolation mechanism for live-Excel hang
  safety; `recalculate`/`run_macro`'s `timeout` param + `CalculationWaitSummary` audit
  summarization (`state_counts`/`last_state`/`poll_count`/`elapsed_seconds`/`outcome`).
- **§3.4** (new): `redirect_external_links`/`restore_external_links` for the linked-consumer-
  workbook scenario (Option 2, decided) — depends on `write_links` existing first.
- **§5.3** (rewritten): `working_dir` replaces `tempfile.mkdtemp()` entirely (fixed
  `excel_runner_runs/<yaml_stem>/` path); read-only sessions now staged too; rename-based
  commit with per-file rollback (`CommitFailure`, `needs_human`); `cleanup()` removed —
  nothing in `working_dir` is ever auto-deleted now.
- **§6.1**: `run_workflow` gains a `working_dir` param; working_dir resolution precedence
  (param > YAML field > cwd); no more `scratch.cleanup()` call.
- **§6.2.1** (new): console/application logging via stdlib `logging` — library code never
  attaches handlers, only the CLI sets a severity threshold.
- **§6.4** (new): `cli.py`'s `--working-dir`/`--logging-level` flags.
- **§8**: build order items 12 (crash/lock-safety hardening — platform-independent, no live
  Excel needed), 13 (hang safety + timeouts — needs Windows Excel), 14 (linked-workbook
  refresh — needs item 10's `write_links` first).

No source code touched this session — Specification.md only. Full test suite reconfirmed green
(285 passed) as a sanity check, unaffected since only docs changed.

**Next step:** Get the user's sign-off on this Specification.md pass, then start building item
12 first (platform-independent, TDD-buildable now, no Windows-Excel dependency) — working_dir
relocation, read-only staging, rename-based commit/rollback, console logging, CLI flags. Items
13/14 wait for a live Excel instance to build/verify against.

## Last Session (2026-08-21e, Windows)
**Status:** Design phase complete for this thread — ready to write Specification.md sections
**Working on:** Finalized §6.3.3 (commit-time file-lock handling) with the user, converging on
a simpler mechanism than first proposed: no separate upfront precheck pass, just attempt each
workbook's commit directly (rename `real_path`→`.bak` then `tmp_path`→`real_path` — the `.bak`
rename is free, since the original was already sitting there untouched pre-commit). If a later
workbook fails, roll back every already-committed workbook in the run by renaming its `.bak`
back, recording per-file whether that rollback itself succeeded. Any file where rollback also
fails is flagged by name as needing a human — the one case the engine can't self-heal, stated
plainly rather than hidden. `.bak` files are transient (deleted on full success, kept on a
human-intervention case). A partial-but-fully-rolled-back commit failure still makes the run's
overall status `"error"`.

Also added **§6.7.1** (new): console/application logging via stdlib `logging`, distinct from
the audit log (real-time human narration vs. after-the-fact evidence). **Corrected after user
feedback**: handler/stream configuration (stdout vs. stderr, formatting) is explicitly out of
scope for excel_runner entirely — both the library and the CLI only ever call `getLogger(...)`,
never attach handlers themselves; that's the responsibility of whatever wraps this (the user's
own log-viewing tool), not something to design here. The only thing the CLI owns is a standard
`--logging-level` argument (`DEBUG`/`INFO`/`WARNING`/`ERROR`, default `INFO`), matching the
user's existing convention across other tools. Content-per-level guidance stands: `INFO` covers
every step/action start+completion; `DEBUG` is a bit more detail, not a firehose;
`WARNING`/`ERROR` must be self-sufficient for a human, not just point at the audit log.

**Next step:** Every item opened in this multi-session design thread (§6.2.3, §6.2.4, §6.3.2,
§6.3.3, §6.3.4, §6.7.1) is now decided. Write the corresponding Specification.md sections next,
then build with TDD, per the project's PRD→Spec→build workflow. Still nothing has been built
yet — this entire thread has been design-only, as the user explicitly requested throughout.

## Last Session (2026-08-21d, Windows)
**Status:** Blocked — awaiting user approval on §6.3.3 (commit-failure structuring) only;
everything else in this design thread is now decided
**Working on:** Finalized §6.2.4 (timeouts/signal detection for `recalculate`/`run_macro`) with
the user. Accepted up front that the underlying signals (`CalculationState`/`Ready`) are
genuinely weak and unreliable ("we don't know what we don't know") — rather than trying to
design a complete detection scheme now, the decision is to capture whatever signal *is*
available, summarize it compactly in the audit log, and keep the surrounding code cheap to
adjust as real-world behavior is observed. Audit summary shape decided: `state_counts`
(histogram of observed signal values, richer than a flat count), `last_state`, `poll_count`,
`elapsed_seconds`, `outcome` (`completed`/`timed_out`/`no_signal_available` — the last one
covers `run_macro`, which has no progress signal at all). Timeout semantics confirmed as final:
if a timeout is specified and elapses, that's a hard failure — no retry, no partial credit,
clean up what's safely possible and stop. Soak-testing real client workbooks (desktop + other
automation contexts) to establish actual reliability is noted as a planned validation activity
once xlwings/COM (item 9) is far enough along, not a blocker for the design itself.

**Next step:** §6.3.3 (commit-time failure when the real file is locked elsewhere — points
1/2/4; point 3 already superseded by §6.3.4) is the only remaining open item before writing
Specification.md. Do not start building any of this yet.

## Last Session (2026-08-21c, Windows)
**Status:** Blocked — awaiting user approval on remaining PRD design decisions (§6.2.4 timeout
design, §6.3.3 commit-failure structuring) before Specification.md work
**Working on:** Settled the working-directory location design (§6.3.4, DECIDED) via a few
rounds of clarifying questions with the user. Supersedes §6.3.3's "expose scratch dir on
RunResult" idea entirely — a fixed, predictable path is more useful than a random path surfaced
after the fact, since external orchestration tooling can construct it itself in advance from
just the yaml's filename, with no need to read any CLI output field.

**Decided**: renamed "scratch dir" → **`working_dir`** throughout (it now holds the audit log
too, not just workbook copies). Location: `<base>/excel_runner_runs/<yaml_stem>/` — `<base>` is
cwd by default, overridable via `--working-dir` CLI flag or a new top-level `working_dir:` YAML
field (CLI > YAML > cwd default). `excel_runner_runs` (not the first-considered
`excel_runner`) was chosen specifically because `excel_runner` collides with this repo's own
package folder name when working inside the repo itself. Inside `working_dir`: `audit.jsonl` at
the root, workbook scratch copies in a `scratch/` subfolder (same internal shape as before,
just relocated) — kept together so the user's stated recovery workflow ("zip this one folder
and hand it to whoever's investigating") has everything in one place. Re-running the same yaml
overwrites the previous run's leftovers automatically, no confirmation. **Cleanup policy
changed**: nothing is ever auto-deleted now, success or failure (previously: scratch/ deleted
on success) — safe since re-running the same yaml just overwrites its own fixed folder rather
than accumulating.

**Next step:** §6.2.4 (configurable/unbounded timeouts for `recalculate`/`run_macro`) and
§6.3.3's remaining commit-failure-structuring points (points 1/2/4 — point 3 is now superseded
by §6.3.4) are still open for approval. Once everything in §6.2.3/6.2.4/6.3.2/6.3.3/6.3.4 is
signed off, write the corresponding Specification.md sections, then build with TDD. Do not
start building any of this yet.

## Last Session (2026-08-21b, Windows)
**Status:** Blocked — awaiting user approval on remaining PRD design decisions before
Specification.md work
**Working on:** Continued the crash/lock-safety design discussion. User confirmed **Option 2**
for §6.3.2 (redirect-then-restore for linked-workbook refresh) — marked DECIDED; Power Query's
exact v1 scope is still open. Two new proposed PRD sections added, neither built yet:
- **§6.2.4**: `recalculate`/`run_macro` need an optional `timeout` param that **defaults to
  waiting indefinitely, not a short default** — some client workbooks' plugin formulas linking
  to large datafiles can legitimately take hours. Researched whether `Application.
  CalculationState`/`Ready` can distinguish "still working" from "hung": both are real COM
  properties, but independently-reported forum evidence shows polling from the same
  thread/procedure that triggered calculation is unreliable (stays `xlPending` indefinitely) —
  a real liveness check would need a separate watchdog process, dovetailing with §6.2.3's
  process-isolation design. `run_macro` has no equivalent progress signal at all. Flagged as
  needing empirical testing against real Excel before finalizing, not just docs/forum research.
- **§6.3.3**: confirmed via direct code inspection (not assumed) that a commit-time failure —
  the real file open elsewhere when a successful run tries to copy scratch content back — is
  **completely unhandled today**: `ScratchManager.commit_all()` has no exception handling,
  `run_workflow()` doesn't catch around it, and the CLI only catches `ExcelRunnerError`, so a
  real `PermissionError` would surface as an unhandled traceback, not clean JSON. Also confirmed
  `RunResult` doesn't expose the scratch directory path at all — even though scratch copies
  survive a failed commit, external tooling (the user's planned automation recovery logic) has
  no way to find them today. Proposed: catch and structure commit failures per-workbook, attempt
  every workbook's commit rather than stopping at the first failure, and expose the scratch dir
  path on `RunResult` unconditionally.
- Also answered (no code change, informational): openpyxl's read-only handle-open behavior has
  the same likelihood regardless of directory, but staging read-only opens into our own scratch
  dir (§6.2.3's earlier finding) changes the *impact* from "shared-file incident" to "harmless
  orphaned temp file" — this is why it's worth doing, not because it prevents the hang itself.
  Confirmed exactly where the scratch dir is created today: `tempfile.mkdtemp(prefix=
  "excel_runner_")` under the OS temp dir, not currently configurable.

**Next step:** Still awaiting explicit approval on §6.2.4/§6.3.3's proposed directions (and
Power Query's v1 scope for §6.3.2) before writing Specification.md sections. Do not start
building any of this yet.

## Last Session (2026-08-21, Windows)
**Status:** Blocked — awaiting user approval on PRD design decisions before Specification.md work
**Working on:** User raised a real crash-safety concern: no matter how careful the scratch-copy
model is, an unplanned crash could still leave a file locked. Researched (not assumed) whether
openpyxl's `read_only` mode is part of that risk — confirmed via openpyxl's own docs plus a
known reported issue: `read_only=True` keeps a genuine OS file handle open until `.close()` or
process exit. But also confirmed a genuine crash is self-healing (the OS releases handles/locks
on process teardown) — the real risk is a **hang**, not a crash, since a stuck-but-alive process
keeps the lock indefinitely until a human kills it. Same bug class as the already-fixed
`OwnedInstanceRegistry.close_owned()` hang.

Separately, user needs robust handling for a real scenario: modify workbook A (file-backend,
scratch-only, not yet committed), then need a *different*, unmodified workbook B — which links
to A via classic cell-reference links, a data connection/Power Query, and/or needs
recalculation — to reflect A's new values before B is itself used. Drafted two proposed PRD
sections for review (not yet approved, not yet built):
- **§6.2.3** (new): live-Excel hang safety — proposes running Excel-interacting work in an
  isolated OS process with an enforced timeout, force-killing on hang, since a Python thread
  can't cleanly interrupt a blocked COM call the way a process boundary can.
- **§6.3.2** (new): refreshing a linked consumer workbook against a not-yet-committed source —
  two candidate mechanisms drafted (checkpoint-commit+rollback vs. redirect-links-then-restore),
  with a recommendation towards the latter (preserves §6.3.1's existing "real files untouched
  until final success" invariant) but explicitly **not decided** — needs the user's sign-off,
  especially on Power Query's v1 scope, before Specification.md work starts.
- Also flagged: §6.3's existing "open read-only in place… faster, safer" bullet needs revisiting
  once the hang-safety decision above is made.

**Next step:** Get user approval/decision on §6.2.3 and §6.3.2's proposed directions. Then write
the corresponding Specification.md sections (per the project's PRD→Spec→build workflow), then
build with TDD. Do not start building any of this yet — explicitly design-first per user request.

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
too, since an external orchestration step is likely to invoke a script path directly rather than
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
workflow from an external orchestration/automation workflow as an external process. `main(argv)`
parses a workflow path
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
