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
5. Which `scratch/working/` copies are actually write-intent (and therefore need committing
   and have an `scratch/originals/` backup) is known statically from the plan — the same
   information already used to decide which workbooks a step writes to. This is the primary
   source of truth, not a file-comparison heuristic. As a defensive check only, commit may
   additionally compare a working copy against its `scratch/originals/` backup (e.g. hash or
   mtime) and warn/skip if a write-intent copy is unexpectedly unchanged, or a read-only copy
   was unexpectedly modified.
6. Whether `scratch/working/` and `scratch/originals/` are deleted after a run, or kept, is
   controlled by a CLI flag (default: delete on full success, keep on failure). This is not
   automatic/implicit behaviour.
7. Running two `excel_runner` invocations concurrently against the same real workbook files is
   not supported and must not be done — this is a documented operator constraint, not something
   the tool detects or locks against. Independently, each run must use its own distinct scratch
   directory (already the case today) so that unrelated concurrent runs at least never collide
   with each other's scratch contents.

## 2. Recalc + link rules

These are the exact, final rules `recalculate` must implement. No case not listed here is
supported.

- **R1 — same-folder, filename-only link.** Sibling must exist, correctly named, in
  `scratch/working/` alongside the linking workbook (sec 1.2). No relinking, ever. Whether the
  sibling is itself write-intent is irrelevant — being open live in the same shared Excel
  process at some point before the parent's own final save is what makes its values fresh; the
  link itself needs no repointing because the name and folder already resolve correctly.
- **R2 — relative link with a subpath (not same folder).** Not supported. Backlog item.
  Scratch does not currently mirror arbitrary relative folder structure.
- **R3 — UNC/absolute link to an unmodified sibling.** Leave untouched. Nothing to do — its
  cached value is already correct because nothing upstream changed.
- **R4 — UNC/absolute link to a modified sibling.** Every time `recalculate` runs against the
  linking workbook, it performs this exact cycle in full, unconditionally (never skipped,
  never cached from a previous run of the same step):
  1. Find the link entry (via `LinkSources`) whose target matches the sibling's known real
     path or basename.
  2. `ChangeLink` that entry to the sibling's *current* scratch path
     (`scratch/working/<sibling>`).
  3. `UpdateLink` (Type=1) on that same link name, to force a fresh read.
  4. Perform the workbook's normal recalculation.
  5. `ChangeLink` that entry back to the sibling's original real path (this must exist on
     disk — the real file is never touched until commit, so it is always there unchanged).
  6. Do **not** call `UpdateLink` again after step 5. Doing so would silently overwrite the
     just-computed fresh value with whatever the original (stale, pre-run) file contains.
  7. Save.
- **R5 — the R4 cycle repeats every call.** If `recalculate` runs on the same workbook twice
  in one workflow, the full R4 cycle happens twice, independently. This is what makes the
  design safe without needing to reason about what state the link was left in between calls.

## 3. Commit process changes

1. Replace the current rename-based commit (`real_path.rename(bak_path)`, then
   `tmp_path.rename(real_path)`) with copy-based overwrite: copy
   `scratch/working/<name>` directly onto `real_path` (`shutil.copy2`), after first copying
   the existing `real_path` (if it exists) to a `.bak` sibling. The original is copied, never
   moved or deleted, before the overwrite.
2. On full success (every staged, write-intent workbook copied over), delete every `.bak`
   created this run.
3. On failure partway through, roll back every already-committed workbook in this call by
   copying its `.bak` back over `real_path` (copy, not rename). If that rollback copy itself
   fails, the `.bak` is left in place and that workbook is flagged as needing manual
   intervention — same semantics as today's `_rollback()`, just copy-based throughout instead
   of rename-based.
4. `scratch/originals/` (sec 1.4) is a second, independent safety net alongside the `.bak`
   files, made only for write-intent workbooks. Since it is kept for the whole run regardless
   of commit outcome (unless the CLI cleanup flag says otherwise, sec 1.6), it remains
   available as an untouched pre-edit reference even if both the commit's own `.bak` and the
   real file end up in an inconsistent state.
