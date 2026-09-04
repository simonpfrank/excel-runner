# Build Plan — Backend Eligibility, Save Blockers and xlw Twins

**Status:** ready to build. Design is settled. Do not reopen it.
**Written:** 2026-09-04, at the end of an evidence-gathering session.
**Intended reader:** a fresh session with no memory of that conversation.

---

## 0. Read this before touching anything

### The one rule that must never be broken

> **openpyxl must never save a workbook that contains outbound external links.**

This was tested. An openpyxl load-and-save with **zero edits** produced a file that Excel
**refused to open at all** — not a broken link, an unopenable workbook. A no-links control file
survived the identical round-trip fine. openpyxl keeps only one relationship per external link
and discards the rest, and the one it keeps is the wrong one.

Evidence is in [docs/link_gaps_andaction_plan.md](docs/link_gaps_andaction_plan.md).

**Do not re-test this. Do not design around "maybe it's fine if...". It is not fine.**

### Already settled — do not re-litigate

| Question | Answer |
|---|---|
| Should we just refuse to write link-bearing workbooks? | **No.** PRD §2 and §6.1: backend selection is never exposed to the author. Refusing is a PRD violation. |
| Does the commit step store a wrong path? | It stores a scratch-relative path that is wrong at the destination, **and it does not matter** — Excel writes a correct absolute fallback and uses it. Verified with a decoy file. Closed. |
| Does openpyxl damage charts? | **Unknown. Untested rumour.** Do not add it as a blocker. Do not write a placeholder for it. |
| Windows vs macOS | Windows is the real target. **Never trade Windows capability for macOS parity.** Note macOS consequences; don't design around them. |

### Project rules that apply

From [.github/copilot-instructions.md](.github/copilot-instructions.md):

- **TDD.** Failing test → implement → pass. Every item below.
- **No stubs.** No `pass`, no `raise NotImplementedError`. If it isn't ready, it's absent.
- **Never claim a test result without running it.** Paste the summarised output.
- **No duplicate functionality.** Several things listed as "to build" below already exist —
  §2 says which. Check before writing.

---

## 1. The design, in full

### 1.1 Save blockers

Workbook inspection answers **"what stops openpyxl saving this file safely?"** and returns a
**set of named blockers**. Never a boolean, never a link-specific flag.

- **Empty set** → the file backend is used as normal.
- **Non-empty** → the workbook must be handled by Excel for anything that leads to a save.

One member exists today: **outbound external links**. "Outbound" means *this workbook
references others*. Inbound links (others referencing this one) are irrelevant — saving this
file doesn't touch their link records.

The set is designed to grow. Members are added **only when empirically verified**.

### 1.2 Promotion

A workbook with a non-empty blocker set is promoted to the `xlw` backend **on its first
write**, not at open.

Why first write and not open: PRD principle 1 — only pay for a live Excel session where the
work genuinely requires it. A workflow that only reads a link-bearing workbook should never
start Excel.

Useful side effect: at the moment of the first write the workbook is **provably clean**, so the
openpyxl save inside `_switch_backend` cannot fire on this path.

### 1.3 Promotion is sticky

Once a session's blocker set is non-empty and it is on `xlw`, **it never returns to `file`**.

Precisely: **`_switch_backend` refuses an `xlw` → `file` demotion whenever the session's
blocker set is non-empty**, regardless of what promoted it.

The demotion itself is harmless — it saves through Excel first, so the file on disk is correct.
The danger is the *next* file-backend write, which would then save through openpyxl and destroy
the workbook.

**Important scoping:** stickiness is conditioned on **blockers**, not on promotion. A workbook
with an empty blocker set that gets promoted by `copy` or `recalculate` (both `com` capability)
demotes normally, exactly as it does today. Do not over-apply this.

### 1.4 read_only never promotes

A session opened `read_only` is never saved, so it can never be corrupted. It stays on `file`
whatever its blockers — which is also the faster path.

### 1.5 The taxonomy for every action

Now normative in [docs/Specification.md](docs/Specification.md) §4.0. Apply it whenever an
action is added or changed:

1. No `workbook:` param → `none`. No backend implementations.
2. Always needs Excel's engine (recalc, macros, live link refresh) → `xlw` only, no file twin.
3. Otherwise → **`file` and `xlw` twins, both mandatory.**
4. Any part with no portable xlwings call → `com` for that part. Windows-only; state the macOS
   behaviour in the docstring.

Rule 3 is the expensive one and it is what this build is mostly about. Because promotion is
sticky, **any action that can land on a promoted workbook needs an xlw implementation** — or
the run dead-ends with no way forward.

### 1.6 Tier discipline

Three tiers, already correctly documented in [docs/Specification.md](docs/Specification.md) §3.
Keep them distinct — collapsing `xlw` and `com` hides which capabilities work on macOS.

- unprefixed → openpyxl
- `xlw_` → xlwings portable API, works on Windows **and** macOS
- `com_` → `.api` escape hatch, raw COM, **Windows only**

**All the new write primitives in this build are `xlw_`, not `com_`.** `range.value`,
`sheets.add()`, `sheet.name`, `sheet.delete()`, `range.column_width`, `range.insert()` are all
portable. So writing to a link-bearing workbook works on macOS. Only the link *operations*
(`LinkSources`, `ChangeLink`, `UpdateLink`) are `com_`, i.e. Windows-only.

---

## 2. What already exists — do not rebuild it

Line numbers verified 2026-09-04.

| Thing | Where | State |
|---|---|---|
| `_switch_backend` (save → close → reopen, both directions) | [engine.py:696](excel_runner/engine.py#L696) | **Built.** Spec §5.2 still says "not built" — that text is stale and flagged in the spec. |
| `xlw_open_workbook`, `xlw_close_workbook`, `xlw_save_workbook`, `xlw_calculate_all` | [backends.py:542](excel_runner/backends.py#L542), [570](excel_runner/backends.py#L570), [580](excel_runner/backends.py#L580), [650](excel_runner/backends.py#L650) | Built |
| `com_link_sources`, `com_change_link`, `com_update_link` | [backends.py:748](excel_runner/backends.py#L748), [765](excel_runner/backends.py#L765), [785](excel_runner/backends.py#L785) | Built |
| `classify_link_target`, `resolve_link_target`, `scan_external_link_targets` | [engine.py:105](excel_runner/engine.py#L105), [140](excel_runner/engine.py#L140), [171](excel_runner/engine.py#L171) | Built — `scan_external_link_targets` reads raw XML via zipfile and is the basis for blocker detection |
| `discover_write_intent_link_graph`, `compute_link_commit_order` | [engine.py:199](excel_runner/engine.py#L199), [242](excel_runner/engine.py#L242) | Built |
| `ACTION_WRITES` (per-action mutates-or-not flag) | [core.py:399](excel_runner/core.py#L399) | Built — already carried on `ActionSpec.writes` |
| `ScratchManager.stage` / `.commit` | [engine.py:503](excel_runner/engine.py#L503), [559](excel_runner/engine.py#L559) | Built |
| `OwnedInstanceRegistry` | [backends.py:802](excel_runner/backends.py#L802) | Built |

**`ACTION_WRITES` is the signal you need for promotion-on-first-write. It already exists.
Don't invent a second one.**

---

## 3. Work items, in order

Do them in this order. Each depends on the ones before it.

---

### W1 — Blocker inspection

**Why:** nothing else can work without it. It is the input to every routing decision.

**What:** a function that takes a workbook path and returns the set of blockers. Reuse
[`scan_external_link_targets`](excel_runner/engine.py#L171) — it already reads the raw XML
without opening Excel. Outbound external links present → the links blocker is in the set.

Return a real typed value (a frozenset of an enum, or equivalent). **Not a dict, not a bool.**
Project rule: no mutable dicts for state.

**Where:** `engine.py`, next to the existing link-scanning functions.

**Tests:** new file `tests/unit/test_save_blockers.py`. Cover: no links → empty; one outbound
link → populated; malformed/non-xlsx input; file that doesn't exist.

---

### W2 — Detect blockers for template-backed workbooks

**Why:** [`discover_write_intent_link_graph`](excel_runner/engine.py#L199) skips a workbook
whose real file doesn't exist yet. A workbook declared `create_if_missing: true` with a
`template:` therefore gets no inspection on its **first** run. If the template carries links,
the new workbook inherits them and we route it to openpyxl and destroy it.

This was a minor gap before. It is now load-bearing: detection gates the entire mechanism.

**What:** when the target file doesn't exist, inspect the `template:` instead.

**Where:** the `path.exists()` check in [engine.py:199](excel_runner/engine.py#L199).

**Tests:** extend `tests/unit/test_link_discovery.py`. Regression: run that whole file.

---

### W3 — Carry blockers on the session, and make promotion sticky

**Why:** the routing decision needs the blocker set at the point of every action dispatch.

**What:**

1. Add a blockers field to `WorkbookSession` ([core.py:505](excel_runner/core.py#L505)).
   Populate it when the session is opened.
2. Change `_needed_backend` ([engine.py:513](excel_runner/engine.py#L513)). Today it maps
   capability alone. It must now resolve from **capability + the session's blockers + whether
   the action writes + the session mode**. Signature change; `get_or_open`
   ([engine.py:599](excel_runner/engine.py#L599)) is the only caller.
3. Make `_switch_backend` ([engine.py:696](excel_runner/engine.py#L696)) refuse `xlw` → `file`
   when blockers are non-empty. Raise a clear, actionable error if something asks for it —
   silently ignoring the request would be worse.

**Watch out:** `_needed_backend` currently returns `"file"` for a `@file_action(writes=True)`
like `write_cell`. That is exactly why no promotion happens today. The writes flag must feed
into the decision.

**Tests:** extend `tests/unit/test_session_manager.py`. Regression: that file plus
`tests/unit/test_registry.py`.

---

### W4 — Fix the openpyxl save inside `_switch_backend`

**Why:** [engine.py:696](excel_runner/engine.py#L696) calls `backends.save_workbook` — openpyxl
— when leaving the `file` backend. On a dirty link-bearing workbook that destroys the links on
the way *into* Excel. The bug is inside the very code written to handle links.

W1–W3 mostly dodge it (the workbook is clean at first-write promotion), but the existing link-
repointing path at [engine.py:800](excel_runner/engine.py#L800) can still hit it. Fix it
properly, don't rely on the dodge.

**Tests:** extend `tests/unit/test_session_manager.py`.

---

### W5 — The xlw backend primitives

**Why:** the bulk of the work. Without these, a promoted workbook can be opened and closed and
nothing else.

**Build in `backends.py`, all `xlw_`-prefixed:**

`xlw_read_range`, `xlw_write_cell`, `xlw_write_range`, `xlw_set_column_width`,
`xlw_create_sheet`, `xlw_rename_sheet`, `xlw_delete_sheet`, `xlw_insert_range`,
`xlw_resolve_range`, `xlw_resolve_sheet_names`, `xlw_read_properties`, `xlw_read_cells`

Each mirrors the signature and return shape of its unprefixed twin. Same inputs, same outputs,
same errors.

**Two things that save real work:**

- **`write_row` needs no primitive.** The action
  ([actions.py:488](excel_runner/actions.py#L488)) is implemented as a loop over `write_cell`.
  Route the inner call and the whole action follows.
- **The four `find_*` functions are pure matching logic over cell values.** Do **not** write
  xlwings twins of the matching. Extract the matching into a backend-agnostic helper, read the
  values through whichever backend is active, and share the logic. Duplicating four matchers is
  exactly the divergence this plan is trying to avoid.

**File size:** `backends.py` is 885 lines and will land somewhere near 1300. **The 500-line
limit is waived for this build** (user decision, 2026-09-04). Use
`--max-module-lines=1600` when running pylint. Do not split the file to satisfy a limit that
has been lifted; the `xlw_`/`com_` naming convention is still the navigation aid.

**Tests:** new file `tests/unit/test_backends_xlw_write.py`. Use the existing
`requires_excel` / `requires_working_xlwings_save` markers from
[tests/unit/conftest.py](tests/unit/conftest.py) — **and explain every skip in the final
report**, per the project rules.

---

### W6 — Route actions to the right twin

**Why:** the primitives are useless until the actions can reach them.

**What:** the ~17 rule-3 actions currently call their unprefixed backend function directly.
They must call the twin matching `session.backend`.

The actions are: `open`, `save`, `close`, `read_range`, `read_metadata`, `write_cell`,
`write_range`, `write_row`, `insert_range`, `set_column_width`, `create_sheet`, `rename_sheet`,
`delete_sheet`, `find_headers_row`, `find_row`, `find_column`, `find_columns`.

Not affected: `recalculate` and `copy` (already `com`), `stop` and `dump` (`none`).

**Design note.** Spec §4 currently says backend choice "is fixed once, by the capability tag,
not decided per-call". That is no longer true for rule-3 actions and §4.0 already records the
qualification. Pick a mechanism and apply it uniformly to all 17 — a per-action `if
session.backend == ...` scattered by hand will rot. A dispatch table keyed on backend, or a
thin resolver in `backends.py`, both work. **Do not build a DI framework** (project rule).

The dispatch site is [runner.py:109](excel_runner/runner.py#L109) `_dispatch`, plus
[runner.py:86](excel_runner/runner.py#L86) `_dispatch_copy`.

**Tests:** this is where the **contract tests** go. One test body parameterised over both
backends asserting identical results — **not** two separate test files. Only one backend is
testable without Excel, so this is the only thing keeping the twins honest.

Regression: `tests/unit/actions/` (the whole directory) and `tests/unit/test_registry.py`.

---

### W7 — Reject links we cannot handle

**Why:** relative-subpath (R2) links cannot survive the flat scratch layout —
`scratch/working/<basename>` and `scratch/originals/<basename>` sit side by side, so a link
into a subfolder resolves to nothing once staged. Today this fails silently.

**What:** detect and refuse with a specific, actionable error at validation time, before any
real workbook is touched (PRD design principle 3).

**Where:** tier-1 [`validate_static`](excel_runner/engine.py#L1156) or tier-3
[`validate_existence`](excel_runner/engine.py#L1472) — tier-3 is the one that already opens real
files, so it is the natural home.

**Tests:** extend `tests/unit/test_validation.py` or `tests/unit/test_validate_existence.py`
depending on where it lands.

---

### W8 — Put blockers in the audit record

**Why:** a run that silently starts Excel and gets slower needs to be explainable. "Why did
this take 40 seconds?" should be answerable from `audit.jsonl`.

**What:** record the blocker set per workbook when a session is created, and record the
promotion event when it happens.

**Tests:** extend `tests/unit/test_audit.py`.

---

## 4. How to run tests during the build

**The full suite is slow. Do not run it repeatedly.**

Per work item, run only:

- the new test file for that item, and
- the specific regression files named in that item.

```powershell
.venv\Scripts\pytest tests\unit\test_save_blockers.py -q
.venv\Scripts\pytest tests\unit\test_session_manager.py tests\unit\test_registry.py -q
.venv\Scripts\pytest tests\unit\actions -q
```

Use `-x` to stop on first failure while iterating. Use `-q`. Don't collect coverage until the
end — it slows every run.

**Run the full suite once**, after W8, before the quality gates.

---

## 5. Quality gates — all must pass

Run these only after every work item is done and the full suite is green.

```powershell
.venv\Scripts\ruff check .
.venv\Scripts\pylint --max-line-length=119 --max-module-lines=1600 .
.venv\Scripts\vulture . --min-confidence 60
.venv\Scripts\pyright .
.venv\Scripts\mypy --strict .
.venv\Scripts\radon cc --min C .
.venv\Scripts\pytest --cov --cov-branch
```

Targets: 90% coverage overall and per module, no ruff errors at line length 119, no mypy
errors, radon ≤ C.

Security audit: no `shell=True`, external input validated, no secrets in code or logs.

**Two things to know:**

- `pylint` is **not** in `[project.optional-dependencies] dev` in
  [pyproject.toml](pyproject.toml). Install it or add it before you reach this step.
- The 500-line module limit in `.github/copilot-instructions.md` is **waived** for this build.
  1600 is the working value. If you'd rather make it permanent, change it in
  `copilot-instructions.md` rather than remembering a flag.

---

## 6. Demos — the real acceptance test

After the full suite and all quality gates pass, run every demo in order, finishing on 08.

```powershell
.venv\Scripts\python -m excel_runner demos\00_generate_demo_workbooks.yaml
.venv\Scripts\python -m excel_runner demos\01_basic_lifecycle.yaml
.venv\Scripts\python -m excel_runner demos\02_read_write.yaml
.venv\Scripts\python -m excel_runner demos\03_structure.yaml
.venv\Scripts\python -m excel_runner demos\04_lookup.yaml
.venv\Scripts\python -m excel_runner demos\05_copy_across_workbooks.yaml
.venv\Scripts\python -m excel_runner demos\06_control_flow_stop.yaml
.venv\Scripts\python -m excel_runner demos\07_full_sequence_pending_recalc.yaml
.venv\Scripts\python -m excel_runner demos\08_full_showcase.yaml
.venv\Scripts\python -m excel_runner demos\08_full_showcase_validate.yaml
```

`08_full_showcase.yaml` is the one that matters. It opens a workbook carrying a genuine
absolute-path external link back to `catalog.xlsx`, and writes to it. **That is precisely the
case this build exists to fix**, and it is currently going through openpyxl.

`08_full_showcase_validate.yaml` reads the output back and asserts the cross-workbook link and
recalculation results using `if:` + `stop`.

**The hope is that no demo YAML needs changing.** If one does, that is a signal worth pausing
on — it may mean the transparent-routing promise has leaked into user-visible syntax, which
would be a PRD violation. Raise it rather than editing the YAML to fit.

---

## 7. Done means

- [ ] W1–W8 complete, each built TDD, no stubs anywhere
- [ ] Full test suite green, every skip explained
- [ ] 90% coverage overall and per module
- [ ] All seven quality gates pass
- [ ] Demos 00 → 08 run clean, including `08_full_showcase_validate.yaml`
- [ ] `08_full_showcase.yaml` output opens in Excel with its links intact
- [ ] [docs/Progress_Tracker.md](docs/Progress_Tracker.md) updated
- [ ] [README.md](README.md) updated
- [ ] Stale text in [docs/Specification.md](docs/Specification.md) §5.2 rewritten
      (`_switch_backend` is built; the doc says it isn't)

---

## 8. Known risks

**Divergence between twins.** Two implementations of one action name, only one testable without
Excel. Contract tests are the mitigation and they are not optional.

**Excel-dependent tests can't run everywhere.** Windows is the target, so this is accepted, not
solved. Skips must be explained, never silently tolerated.

**Sticky promotion costs read performance.** Once promoted, reads go through Excel too. That is
the accepted trade — the alternative is a second concurrent openpyxl handle on a live workbook,
which is ruled out. Measure later if it hurts.

**Scope creep into charts.** The charts rumour is untested. It does not go in this build.

---

## 9. Where the background is

- [docs/link_gaps_andaction_plan.md](docs/link_gaps_andaction_plan.md) — plain-language facts,
  what's broken, measured evidence
- [docs/r4_link_discovery_gaps_proposal.md](docs/r4_link_discovery_gaps_proposal.md) — original
  audit trail and working-out
- [docs/Specification.md](docs/Specification.md) §3, §4.0, §5.2.1 — the normative rules
- [docs/PRD.md](docs/PRD.md) §2, §6.1 — why refusing the job was never an option
