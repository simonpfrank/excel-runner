# R4 External-Link Handling: Rules, Gaps, and Ordered Fix Plan

Started while extending `demos/08_full_showcase.yaml` with a real, Excel-authored R4
(absolute-path) external link. Now covers: how Excel actually persists links (measured), the
resulting hard invariant, six gaps in `excel_runner/engine.py` and the `file` backend, the
impact on action functionality, and an ordered fix plan. Companion to
`docs/recalc_and_link_refresh_plan.md`, which this does not revise.

A further issue found in the same investigation — `classify_link_target` not recognizing the
driveless-rooted absolute path form real Excel writes for same-drive links — is **already
fixed** (TDD, `tests/unit/test_link_discovery.py`, all quality gates clean).

## Excel external-link persistence rules (empirically confirmed)

Established via a controlled test matrix: one linking workbook with 5 external links (to a
same-folder file, a same-drive-different-folder file, a same-drive-but-several-folders-away
file (`C:\temp\...`), a UNC path, and a mapped-drive-letter path), each independently opened and
saved by hand in real Excel (no automation) from several different disk locations, with raw
rels XML and openpyxl's object-model view captured before/after every move. Full raw evidence
in the session transcript. These rules describe **Excel's own behavior** — how Excel itself
decides what to persist when a workbook with external links is saved. We never write links with
`openpyxl` (see note after rule 5 on why that's irrelevant here); `openpyxl` is used only to
*inspect* what Excel already wrote.

1. **Relative vs. absolute is decided solely by whether the link target is on the same drive
   letter/volume as the workbook itself, re-evaluated fresh on every save** — not by folder
   depth, not by how the link was originally authored, not by any cached/original state.
   - Target on the same drive as the workbook → stored **relative**, and correctly
     re-derived for the workbook's current location on every single save (confirmed across a
     same-folder move, a one-folder-up move, and a different-subfolder move).
   - Target on a different drive letter, or a UNC path → stored **absolute**
     (`file:///...` form), and left **byte-for-byte unchanged** across saves/moves — Excel
     never re-derives or re-validates it once it's absolute, even after the workbook itself
     moves to a different location.
   - Moving the *workbook* itself onto the same drive as a previously-absolute target flips
     that link to relative on the next save (and vice versa) — confirmed by moving the linking
     workbook onto a mapped `Y:` drive: every `C:`-drive target flipped to absolute, while the
     target that was already on `Y:` flipped from absolute to relative.

2. **Same-drive links to a target several folder levels away get a redundant second
   relationship** — Excel writes both a relative path *and* a driveless-rooted absolute
   fallback (e.g. `../../temp/_test_c_temp.xlsx` **and** `/temp/_test_c_temp.xlsx`) for the
   same link. This only happens for same-drive links; cross-drive/UNC links never get a second
   relationship, they're always a single absolute entry.

3. **A link's target-file *name*, not just its folder, must exist for Excel to resolve it** —
   moving the target to a path that doesn't match what's recorded (wrong folder depth and/or
   wrong filename) leaves the link genuinely broken; Excel does not silently repoint it to
   something plausible-looking on save.

4. **The formula bar normally does show the true absolute path (drive letter or UNC) for an
   absolute link** — e.g. it correctly displayed `Y:\...` once the linker workbook was on the
   `Y:` drive. The only case observed where the formula bar *omits* the path is when the
   *linked-to* workbook is itself open at the same time in Excel: formulas referencing it then
   temporarily display with no path at all (just `[filename]Sheet!Cell`), since Excel resolves
   the reference live against the open workbook instance instead of a path on disk. This is
   display-only and does not get written on save — saving while the target is open still
   persists the **original absolute path**, unchanged; the no-path display never reaches disk.
   It's fair to speculate, though not separately confirmed, that once a link's stored `Target`
   is an absolute `file:///...` form it simply remains that way — Excel does not re-derive it
   back down to relative even when the on-screen formula bar briefly shows no path at all.

5. **External-link relationship IDs (and therefore which `externalLinkN.xml` file backs which
   formula's `[N]` reference) can be renumbered by Excel on save**, independent of which link
   is which conceptually. Any tooling that inspects raw rels/`openpyxl` state must resolve by
   the formula's actual bracket index (or by reading the target file to confirm identity), not
   by positional/file-name ordering, or it will misattribute which link is which.

**Note on `openpyxl` (inspection-only, not a rule about Excel):** `openpyxl` keeps only the
first relationship in file order for any link with more than one — so when *reading* a file
Excel wrote per rule 2, it silently discards the redundant absolute-rooted entry, keeping the
relative fallback. This matters only for tooling that inspects link state with `openpyxl` (as
this investigation's diagnostic scripts did); it has no bearing on what real Excel actually
persists, and is moot for `excel_runner` itself since we never save a workbook containing links
via `openpyxl` — Excel (via COM) is always the one writing them.

## Hard invariant

**Never save a link-bearing workbook with `openpyxl`.** openpyxl may only *inspect* links —
`scan_external_link_targets` (`engine.py:171`) reads the rels XML out of the zip and never
saves. Every write to a link-bearing workbook must go through Excel/COM.

Two independent reasons, both from the rules above:

1. openpyxl holds one relationship per link, so it silently destroys the two-relationship form
   Excel writes (rule 2) — on any save, with zero cell edits.
2. openpyxl writes the stored Target back verbatim; it never recomputes a relative path for
   the file's location the way Excel does on every save (rule 1). A workbook openpyxl saves in
   scratch keeps scratch-relative paths when copied to its real path.

**The damage is unrecoverable.** The discarded relationship is not recorded anywhere else in
the file, so the saved workbook contains no trace of the original target. Excel opening it
afterwards sees only the surviving path and resolves that — it cannot restore what is no longer
there, and has nothing to warn about. (Inference from rule 3 plus the file contents, not a
separate measurement: the bytes are gone.) There is no repair step, which is why this is an
invariant rather than a preference.

Current code violates this invariant in three places (G2, G3, G4).

## Gaps

### G1 — link scan skips not-yet-materialized `create_if_missing` workbooks

**Where:** `discover_write_intent_link_graph` (`engine.py:199`), `path.exists()` gate at
`engine.py:227-230`. Called once from `runner.py` after `plan()`.

**What:** Only the workbook's *real* file is scanned. A `create_if_missing: true` + `template:`
workbook has no real file on its first run, so `continue` fires and its template's baked-in
links are never seen, never in the graph, never wired.

**Fix:** Fall back to the `template:` workbook's file when the real file is absent, mirroring
the existing precedent in `_resolve_check_path` (`engine.py:1317`):

```python
if direct.exists():
    return direct
if ref.template is not None:
    template_path = Path(workflow.workbooks[ref.template].file)
    if template_path.exists():
        return template_path
```

Needs `Workflow` (or a name→template-path map) threaded in — it currently receives only
`workbook_paths` and `write_intent`. Purely additive: a workbook with no `template:` is
unaffected.

**Cost of not fixing:** A template-baked R4 link is invisible on the run that creates the real
file. It is copied in as inert bytes by `create_workbook`, never repointed at staging, never
reverted at commit — Excel later resolves it against its stale template-baked path.

### G2 — per-step checkpoint saves link-bearing workbooks with openpyxl

**Where:** `_save_dirty_staged_sessions` (`engine.py:832`), called by `checkpoint()`
(`engine.py:841`) after **every step** and again at the head of `commit_all()`
(`engine.py:853`).

**What the code does today (this is the violation, not permitted behaviour):** the save is
dirty-gated, which determines only *when* the corruption fires, not whether it is safe:
- A link-bearing workbook that is only **read** is never dirty, so this path never runs for it.
  Safe by accident, not by design — nothing here checks for links.
- A link-bearing workbook that is **written** is marked dirty by the action and then saved with
  `backends.save_workbook` (`backends.py:99` → openpyxl) at the very next step boundary. That
  save destroys the link, unrecoverably, per the invariant above.

**Fix direction:** this path must never call `save_workbook` on a workbook with links —
either the workbook is rejected up front (option A) or the save routes through COM (option B).

**Cost of not fixing:** Any file-backend write action on a link-bearing workbook corrupts its
links at the next step boundary, silently. This is Issue 2's original finding, now generalised:
not only R4, and not only at commit.

### G3 — backend switch saves with openpyxl immediately before switching to COM

**Where:** `_switch_backend` (`engine.py:696`) saves a dirty session via `save_workbook`
(openpyxl) *before* closing and reopening on `xlw`. `_wire_one_r4_link` (`engine.py:766`)
calls it precisely to enable `ChangeLink`.

**Cost of not fixing:** Whenever the source session is already dirty when its link partner is
staged, R4 wiring corrupts the link on the way to repointing it — the failure mode occurs
inside the machinery that exists to prevent it.

### G4 — commit-time revert bakes a scratch-relative path into the committed file — **UNVERIFIED, test first**

**Where:** `_revert_r4_links_before_commit` (`engine.py:785`) does
`com_change_link(→ target real path)` then `xlw_save_workbook`, **while the file is still at
`scratch/working/<name>.xlsx`**. `ScratchManager.commit` (`engine.py:391`) then `shutil.copy2`s
it to the real path.

**Why this is suspect:** By rule 1, Excel stores a same-drive target relative to the file's
*current* location — scratch. By the copy rule, `copy2` changes no link bytes. The committed
workbook therefore carries a path computed for the scratch folder.

**Unknown:** rule 2 says Excel also writes a driveless-rooted absolute fallback for distant
same-drive targets. Whether Excel uses that fallback when the relative path fails to resolve is
not established. **Measure before designing a fix:** commit a real R4 pair, then dump the
committed file's rels and open it in Excel to see whether the link resolves.

**Cost if real:** Every R4 commit produces a silently wrong or broken link.

### G5 — R2 (`relative_subpath`) links are classified but neither handled nor rejected

**Where:** `classify_link_target` returns `"relative_subpath"` (`engine.py:147`);
`discover_write_intent_link_graph` skips everything that isn't `"absolute"` (`engine.py:231`).
Scratch is flat — `working_path = scratch/working/<basename>` (`engine.py:362`).

**What:** The flat layout is why R1 (bare filename) survives staging untouched with no
repointing. An R2 `sub/x.xlsx` link cannot survive it, and nothing detects that.

**Cost of not fixing:** Silently broken link, no error. "Backlog/unsupported" is documented but
unenforced.

### G6 — no validation rejects the write + links combination

**Where:** Nothing in tier 1 (`validate_static`) or tier 2 (`plan`) compares link presence
against per-workbook write mode. Established pattern is
`raise ValidationError(ErrorDetail(message=..., technical_reason=..., suggestion=...))`.

**Cost of not fixing:** G2/G3/G5 fail silently at runtime rather than loudly before any file is
opened.

**Adjacent, out of scope:** flat scratch collides two declared workbooks sharing a basename in
different real folders, which also breaks R1 resolution. Separate issue, worth its own note.

## Action impact of enforcing the invariant

`backends.py` has no COM read/write primitives. Its COM surface is open/save/close,
`com_copy_range`, the calculate family, and `com_link_sources` / `com_change_link` /
`com_update_link` (`backends.py:748-797`). Everything else is `@file_action` (openpyxl).

- **COM-capable actions:** `copy`, `recalculate`.
- **openpyxl-only writes:** `save`, `write_cell`, `write_range`, `write_row`, `insert_range`,
  `set_column_width`, `create_sheet`, `rename_sheet`, `delete_sheet`.
- **openpyxl-only reads (unaffected — reads never save):** `open`, `close`, `read_range`,
  `read_metadata`, `find_headers_row`, `find_row`, `find_column`, `find_columns`.

**Consequence: a link-bearing workbook currently cannot be written at all** without violating
the invariant. There is no COM fallback for any of the nine write actions. Reads are safe.

## Alternatives for the write path

Applies to G2, G3 and the action impact above. Pick one before implementing either.

**A — Reject at validation (stopgap).** Fail the run if a workbook with links is a write target
of any openpyxl action.
- *Pros:* Small, loud, no silent corruption, no new backend code. Ships now.
- *Cons:* Removes a real capability outright — no link-bearing workbook can be written. Blocks
  the showcase demo's own use case.

**B — Add COM implementations for the nine write actions.** Route them to `xlw` whenever the
target workbook has links.
- *Pros:* Removes the limitation properly. Reuses the existing per-action capability +
  `_switch_backend` machinery. Excel then handles link persistence itself, so rules 1 and 2 stop
  being our problem.
- *Cons:* Largest change by far — nine new backend primitives plus per-action dispatch. Requires
  a live Excel for workflows that previously needed none. Slower. Doubles the test surface for
  every write action.

**C — Strip links before an openpyxl save, reattach the rels bytes after.**
- *Pros:* Keeps all actions on openpyxl. No new COM surface.
- *Cons:* Hand-editing the zip's rels XML around every save. Fragile and outside any supported
  API. Cannot reproduce rule 1's location-dependent recomputation, so it preserves stale paths
  by construction. Rejected.

**D — Route only the *save* through COM, keep edits in openpyxl.**
- *Pros:* One new code path, not nine. Excel does the persisting, so rules 1 and 2 are handled.
- *Cons:* Requires closing the openpyxl handle, reopening in Excel and re-saving on every
  checkpoint — slow, and openpyxl's own write of the edits still happens first, which is the
  corrupting step. Does not actually work. Rejected unless openpyxl can be made to emit the
  file without touching link parts.

**Alternatives for G4** (decide after the measurement):

**G4-a — Commit first, fix the link in place afterwards.** Copy scratch → real path, then open
the committed file over COM, `ChangeLink` to the real target, save. Excel then computes the
path from the file's final location.
- *Pros:* Correct by construction under rule 1. No assumption about fallback behaviour.
- *Cons:* Changes commit ordering — the file is briefly on disk with a scratch-relative link.
  Rollback semantics need rechecking. Needs Excel open at commit time.

**G4-b — Match scratch folder layout to production layout** for any R4-linked pair, so the
relative offset is identical in both places and no repointing is needed at all.
- *Pros:* Extends R1's "nothing to rewrite" guarantee to R4. Removes wiring and reverting
  entirely.
- *Cons:* Scratch is currently flat; making it mirror production is a structural change
  affecting staging, collision handling and the recovery artifact layout. Only works when every
  linked pair can share one consistent offset.

**G4-c — Do nothing, if the measurement shows Excel's absolute fallback resolves correctly.**
- *Pros:* Zero cost.
- *Cons:* Depends on undocumented Excel fallback behaviour, and rule 2 says that fallback is
  only written for *distant* same-drive targets — so it would not cover near ones.

## Ordered plan

| # | Change | Why now | Cost of not doing it |
|---|---|---|---|
| 1 | Measure G4: commit a real R4 pair, dump the committed file's rels, open in Excel | Pure measurement, no code. Every G4 option depends on the result | Design work on G4 proceeds on a guess |
| 2 | G6 validation: reject links + openpyxl write on the same workbook | Cheapest way to convert silent corruption into a clear pre-run error. Independent of the A/B decision | G2/G3/G5 keep failing silently |
| 3 | G3 ordering fix: never openpyxl-save on the way into a COM switch | One-line-scale fix, directly inside the R4 machinery | R4 wiring keeps corrupting the link it repoints |
| 4 | G1 template fallback in `discover_write_intent_link_graph` | Small, precedented, opt-in, testable alone | Template-baked links invisible on first run |
| 5 | G5: reject R2 links at validation | Same mechanism as step 2, closes the last silent-breakage class | Silently broken links with no error |
| 6 | Decide A vs B, then implement | Restores (B) or formally removes (A) the ability to write link-bearing workbooks | Capability stays blocked by step 2's validation with no path forward |
| 7 | Implement the chosen G4 option | Correct link state in committed files | Every R4 commit possibly ships a wrong link |

Steps 1–5 are independent of each other and of the A/B decision. Steps 6 and 7 are the design
work.

## Appendix — evidence trail (narrative superseded by the rules above)

### Issue 2 — confirmed `openpyxl` bug: it keeps only one Relationship per external link

**Where:** `.venv/Lib/site-packages/openpyxl/workbook/external_link/external.py:188` (read
side) and `.venv/Lib/site-packages/openpyxl/writer/excel.py:257-269` (write side), in the
installed `openpyxl` package itself — not excel_runner code. Affects every file-backend save
(`excel_runner/backends.py:99`, `save_workbook`, used by `write_cell`, `write_range`,
`create_sheet`, etc.), since all of them go through `openpyxl`'s writer.

**Root cause, pinned down exactly:** Real Excel writes **two** `Relationship` entries in one
`externalLinkN.xml.rels` for a single same-drive external link — a driveless-rooted absolute
one and a same-folder-relative fallback — confirmed by dumping the fixture's raw rels:

```xml
<Relationship Id="rId2" Target="../../catalog.xlsx" TargetMode="External"/>
<Relationship Id="rId1" Target="/Dev/projects/excel-runner/demos/08_full_showcase/catalog.xlsx" TargetMode="External"/>
```

`openpyxl`'s reader (`read_external_link`, `external.py:181-189`) does:
```python
deps = get_dependents(archive, link_path)
book.file_link = deps[0]
```
— it keeps **only the first** relationship in file order (here, `rId2`, the relative
fallback) as the link's sole `file_link`, and simply has no field to hold a second one. The
writer (`writer/excel.py:266-268`) then only ever emits that one stored `file_link`:
```python
rels = RelationshipList()
rels.append(link.file_link)
```
So this is not a "some links get lost sometimes" fuzz result — it is a **structural
one-relationship-per-link limitation** in `openpyxl`'s external-link model. It reads whichever
relationship happens to be first in the XML and permanently discards any other, independent of
`TargetMode` or which one is the "real" absolute target. `keep_links=True` (the `load_workbook`
default, already in effect in the round-trip test) does not help — it governs whether external
links are parsed *at all* (vs. dropped en masse to save memory), not how many relationships
per link are retained; this workbook's single external link was faithfully "kept", just
collapsed from two relationship entries down to one on the way through.

**What happens:** Confirmed by an isolated test: copying the R4 fixture and doing
`openpyxl.load_workbook(path).save(path)` with **zero cell edits** drops the absolute
relationship, keeping only the relative fallback — exactly as the code above predicts, since
`rId2` (relative) sorts first in the file.

**Consequence:** Any workbook staged with an R4 link, then touched by *any* file-backend write
action before that link is reverted, has its real external link destroyed the moment
`save_workbook` runs — regardless of whether Issue 1 is fixed and wiring is scheduled
correctly. This is a genuine `openpyxl` limitation, not something `keep_links` or any other
existing `load_workbook`/`save` parameter can turn off.

### Issue 2b — reframed: this isn't fixable by switching backends either

**This supersedes the "Options considered" list originally written for Issue 2.** Before
proposing a fix, I tested whether routing the save through `xlwings`/COM instead of `openpyxl`
(Option 1 below) actually avoids the problem, since COM is Excel's own file format engine.

**Test 1 — same fixture, no edits, open+save-as via COM (`xlwings`):**
```
original:     rId2 Target="../../catalog.xlsx"                                       (relative fallback)
              rId1 Target="/Dev/projects/excel-runner/demos/08_full_showcase/catalog.xlsx"  (absolute)

after COM save-as (saved to project root): rId1 Target="demos/08_full_showcase/catalog.xlsx"
```
COM does **not** preserve either original relationship. It collapses both down to a single,
**freshly recomputed** relationship — a path relative to wherever the file was just saved.
It happened to resolve correctly in this test only because the file was saved at the project
root, and root→`demos/08_full_showcase/catalog.xlsx` genuinely is the real relative offset. If
the file had been saved to a folder at a different relative offset from the target, the same
recomputation would have produced a link pointing at the wrong place — not a crash, not an
error, just a silently-wrong resolved path.

**Test 2 — a fixture with links to two different target files** (not two relationships for
one link — two independent `externalLink` parts), open+save-as via COM, no edits:
```
before: externalLink1 -> demos/08_full_showcase/catalog.xlsx
        externalLink2 -> _two_links_report.xlsx
after:  externalLink1 -> demos/08_full_showcase/catalog.xlsx   (unchanged)
        externalLink2 -> _two_links_report.xlsx                (unchanged)
```
Distinct targets stay independent and each is correctly re-relativized to the file's current
location — no cross-contamination between them. (This fixture's links only ever had one
relationship each to begin with, since Excel apparently only writes the redundant
absolute-plus-relative pair for deeper/more-indirect relative paths — not chased further, not
relevant to the conclusion below.)

**Conclusion (superseded by Issue 2c below — kept for the audit trail):** `openpyxl` and
Excel/COM both fail to preserve an absolute link string across a save for *this specific link
form* — but for different reasons that lead to the same place. `openpyxl` keeps a *stale*
relative fallback (wrong if anything moved). COM always writes a *freshly recomputed* relative
path, correct for the file's current location. R4's design (`ChangeLink` a link's *absolute*
target to a scratch path, then `ChangeLink` it back to the *absolute* real path before commit)
is fighting a persistence model that, for driveless-rooted-absolute same-drive links
specifically, doesn't round-trip in absolute terms.

### Issue 2c — refined further: the collapse is specific to R4's own link form, not absolute links generally

Following up on a direct question ("does it depend on same drive vs. different?"), I ran two
more isolated tests to pin down exactly when Excel writes the redundant two-relationship pair
that `openpyxl` collapses, versus when a link is stored as a single relationship that survives
untouched.

**Test 3 — two genuinely different link targets, one built from a path copied from elsewhere,
one from a literal same-drive path:** produced one link stored relative
(`demos/08_full_showcase/catalog.xlsx`) and, unexpectedly, the other stored as a fully
**absolute UNC path** (`\\int\ict\eu\unify\Sandbox\Simon Frank\_two_links_report.xlsx`) —
even though both files physically live on the same drive, same folder. Both forms survived an
`openpyxl` re-save and an Excel/COM re-save completely unchanged (confirmed independently via
`book.api.LinkSources()` and `openpyxl`'s raw `file_link.Target`).

**Test 4 — isolating the real variable, controlling for provenance:**
- *Test A:* two link targets freshly created in-place (no copy from anywhere), same drive/folder
  as the linking workbook → **both stored relative.**
- *Test B:* the **same physical file**, referenced two different ways in the link formula:
  once via its ordinary `C:\Dev\...` drive-letter path, once via the UNC-equivalent
  `\\localhost\c$\Dev\...` path to the identical file → the drive-letter form was stored
  **relative**; the UNC form was stored **absolute, verbatim**, unchanged across an `openpyxl`
  round-trip.

**Conclusion:** Excel decides relative-vs-absolute purely from **the literal path string used
to build the link**, not from any resolved "true" volume/drive identity check. A link built
from a plain drive-letter path on the same drive as the workbook → single relationship, stored
relative, and this **does** survive `openpyxl` and COM saves untouched. A link built from (or
that resolves through, e.g. via a corporate storage/sync layer) a UNC-form path → single
relationship, stored absolute, and this **also** survives both backends' saves untouched.

**The one confirmed case that still breaks under `openpyxl` (Issue 2's original finding) is
narrower than first framed:** it is specific to the *driveless-rooted absolute* form
(`/Dev/projects/excel-runner/...`, no drive letter, no UNC prefix) that real Excel writes for
same-drive links belonging to R4's own classification. For that form only, Excel writes **two**
relationships (that absolute-rooted one plus a relative fallback), and `openpyxl`'s
one-relationship-per-link model silently keeps just the relative fallback and discards the
absolute one — corrupting exactly the kind of link R4 exists to handle. Genuinely
UNC-absolute or drive-letter-relative links (single relationship each) are **not** affected by
this `openpyxl` limitation at all.

**Caveat this doesn't remove:** for a workbook whose R4 link is in the driveless-rooted-absolute
form, saving it through `openpyxl` still **destroys/corrupts that link** (silently drops the
absolute relationship, keeps a stale relative one) — this is unchanged from Issue 2's original
finding and remains the concrete risk any fix needs to address. It just isn't a
blanket "absolute links never survive" problem — it's specific to R4's own link shape.

**What this changes:** Options 1 and 2 originally listed below (permanently pinning the
workbook to `xlw`, or patching rels bytes back in after an `openpyxl` save) are still not
viable for the driveless-rooted-absolute form specifically — COM recomputes it fresh (Test 1)
rather than restoring the original string, and `openpyxl` drops it outright (Issue 2). The real
fix has to be elsewhere: since Excel always resolves a link relative to the file's current disk
location, a workbook pair never actually needs an absolute link *at all* if its scratch copies
sit at the **same relative folder offset from each other as their real files do** — that's
already exactly why R1 (same-folder links) needs no repointing at all. R4 only exists because a
workbook's scratch copy can currently land at a different relative offset than production. Two
possible directions follow from this, requiring a decision:

- **(a) Match scratch's relative folder layout to production's**, for any pair connected by an
  R4 link, so the link never needs touching in the first place — closer to extending R1's
  existing "no rewrite needed" guarantee to cover this case, rather than adding recompute
  machinery.
- **(b) Accept that an R4 "revert" is really a fresh recomputation, not a literal restore** —
  at commit time, after the workbook lands at its real final path, explicitly re-resolve
  and rewrite the link (via `ChangeLink`, over COM) once more, using whatever relative form
  Excel derives for that real, final location — rather than trying to string-restore the
  original absolute value.

Both need real design work before implementation — this is not a small patch either way,
unlike Issue 1.

### Re-scoping history

1. First pass (Issue 2): "an `openpyxl` bug" — one `Relationship` per external link
   (`external.py:188`, `writer/excel.py:266-268`), confirmed by a raw rels dump plus an
   isolated round-trip test that drops the absolute relationship with zero cell edits.
2. Second pass (Issue 2b): over-generalized to "neither `openpyxl` nor COM can round-trip an
   absolute link string".
3. Third pass (Issue 2c): narrowed to the driveless-rooted-absolute form only.
4. Fourth pass (rules section at the top): the whole relative/absolute question resolved
   empirically via the 5-target manual Excel test matrix. That section is authoritative; the
   Issue 2/2b/2c narratives above are kept only as the audit trail.

No code changes made yet — this document is for review/decision before touching `engine.py`,
`backends.py`, or the showcase demo's `linked_workbook` steps.

