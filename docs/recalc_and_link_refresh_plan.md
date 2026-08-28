# Recalc and Link Refresh Plan

Requirement document for making `recalculate` link-aware, and the scratch/commit changes
that requirement depends on. Written to be checked for correctness before any code changes.

## 1. Scratch layout changes

1. Scratch gains two subfolders instead of one flat directory: `scratch/working/` and
   `scratch/originals/`. They serve two different, unrelated purposes — link resolution and
   commit-safety backup, respectively — not a "before/after" pair.
2. `scratch/working/` holds every workbook that needs to physically sit in the same folder as
   another workbook for a same-folder relative link between them to resolve unchanged (R1).
   This includes workbooks that are themselves write-intent (being edited) *and* read-only
   siblings that are never written to but are placed here purely so the relative path
   resolves. All copies here are named with the workbook's **original real basename**, not the
   workflow YAML key.
3. A workbook that is never written to, and is never the target of a same-folder relative
   link from a workbook in `scratch/working/`, gets no scratch copy at all — it is opened
   directly at its real path. Copying it would serve no purpose.
4. `scratch/originals/` holds a pre-edit backup copy, made only for workbooks that are
   write-intent (at least one step writes to them). It exists purely for commit-time rollback
   safety (sec 3) and for R4's "restore the link to its real, untouched state" step — it is
   not a link-resolution mechanism and read-only workbooks never get a copy here.
5. `scratch/working/` is therefore a mix, per workbook (sec 1.3): some entries are write-intent
   (edited, committed, backed up in `scratch/originals/`) and some are read-only siblings
   present purely so a same-folder relative link resolves (never committed, no backup). Which
   is which is known statically from the plan — the same information already used to decide
   which workbooks a step writes to. This is the primary source of truth, not a
   file-comparison heuristic. As a defensive check only, commit may additionally compare a
   working copy against its `scratch/originals/` backup (e.g. hash or mtime) and warn/skip if
   a write-intent copy is unexpectedly unchanged, or a read-only copy was unexpectedly
   modified.
6. Whether `scratch/working/` and `scratch/originals/` are deleted after a run, or kept, is
   controlled by a CLI flag (default: delete on full success, keep on failure). This is not
   automatic/implicit behaviour.
7. Running two `excel_runner` invocations concurrently against the same real workbook files is
   not supported and must not be done — this is a documented operator constraint, not something
   the tool detects or locks against. Independently, each run must use its own distinct scratch
   directory (already the case today) so that unrelated concurrent runs at least never collide
   with each other's scratch contents.

## 2. Recalc + link rules

These are the exact, final rules `recalculate`/staging/commit must implement. No case not
listed here is supported.

- **R1 — same-folder, filename-only link.** Sibling must exist, correctly named, in
  `scratch/working/` alongside the linking workbook (sec 1.2). No relinking, ever. Whether the
  sibling is itself write-intent is irrelevant — being open live in the same shared Excel
  process at some point before the parent's own final save is what makes its values fresh; the
  link itself needs no repointing because the name and folder already resolve correctly.
- **R2 — relative link with a subpath (not same folder).** Not supported. Backlog item.
  Scratch does not currently mirror arbitrary relative folder structure.
- **R3 — UNC/absolute link to an unmodified sibling.** Leave untouched. Nothing to do — its
  cached value is already correct because nothing upstream changed, and the real file's own
  path never changes.
- **R4 — UNC/absolute link to a workbook that will be modified this run.** The link is edited
  exactly twice for the whole run, not once per `recalculate` call:
  1. **At staging time** (once, when both workbooks are first staged): `ChangeLink` the
     linking workbook's link from its real path to the target's scratch path
     (`scratch/working/<target>`). Confirmed safe (probe7/probe9): the target's scratch copy
     already exists on disk at this point, so this does not blank the cached value.
  2. **For the rest of the run**, nothing link-specific happens. `recalculate` needs no
     link-awareness at all — it is the plain, already-built action. Normal Excel mechanics
     (live update if both are open in the same shared process, or a plain recalculation
     against the on-disk scratch path otherwise) keep the linking workbook's values correct
     as the target workbook is edited.
  3. **At commit time** (once, per sec 3): after the target workbook has already been
     committed (its real path overwritten with its final scratch content), the linking
     workbook's link is switched back from the target's scratch path to the target's real
     path, and saved. Because the real path now holds the exact same content as the scratch
     copy did, this revert produces the correct, fresh value — confirmed empirically
     (probe10) that this works correctly on save with **no** explicit `UpdateLink`/`calculate()`
     call needed, including under manual calculation mode.
- **R5 — commit order is a dependency, not a list.** A workbook that is the target of another
  write-intent workbook's R4 link must be fully committed (sec 3) before that other workbook's
  link-revert-and-save step runs. This is a topological order over the "which write-intent
  workbook links to which other write-intent workbook" graph, computed once, not an arbitrary
  iteration order.
- **R6 — cyclical R4 links between two write-intent workbooks are rejected.** If workbook A
  has an R4 link to workbook B, and B also has an R4 link to A, there is no valid commit order.
  This must be detected during planning and raised as a clear error, not silently misordered.
- **R7 — link depth beyond one hop is out of scope (backlog).** If a workbook that is itself
  the target of an R4 link is *also* found to have its own outbound R4 link to a different
  write-intent workbook (a chain, e.g. A→B→C), this is not resolved. It must be detected and
  raised as a clear error rather than silently committed in a possibly-wrong order.

## 3. Commit process changes

1. Compute a commit order over all write-intent workbooks: any workbook with no outbound R4
   link to another write-intent workbook can be committed first; a workbook with such a link
   must come after its target(s). Reject with a clear error on any cycle (R6) or chain (R7).
2. For each write-intent workbook, in that order:
   1. If it has any outbound R4 link(s) (i.e. to a target workbook, which by this point in the
      order is already committed and holds fresh content at its real path): open it via the
      `xlw` backend if it isn't already open there (temporary backend switch regardless of
      which backend was used during the run — `ChangeLink`/save requires COM), `ChangeLink`
      each such link from the target's scratch path back to the target's real path, then save.
      No extra recalculation step is required (probe10).
   2. Copy the workbook's (now-correct) scratch file directly onto its real path
      (`shutil.copy2`), after first copying the existing real file (if any) to a `.bak`
      sibling. The original is copied, never moved or deleted, before the overwrite.
3. On full success (every staged, write-intent workbook committed), delete every `.bak`
   created this run.
4. On failure partway through, roll back every already-committed workbook in this call by
   copying its `.bak` back over `real_path` (copy, not rename). If that rollback copy itself
   fails, the `.bak` is left in place and that workbook is flagged as needing manual
   intervention. Because commit order follows the link dependency graph (R5), a partial
   failure can legitimately leave a target workbook committed while a workbook that links to
   it is not — this is an accepted, self-consistent state, not a corruption.
5. `scratch/originals/` (sec 1.4) remains a second, independent safety net alongside the
   `.bak` files, made only for write-intent workbooks. Since it is kept for the whole run
   regardless of commit outcome (unless the CLI cleanup flag says otherwise, sec 1.6), it
   remains available as an untouched pre-edit reference even if both the commit's own `.bak`
   and the real file end up in an inconsistent state.
