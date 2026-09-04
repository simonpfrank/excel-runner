# External Links: What We Know, What's Broken, What To Do

## About this document

Excel workbooks can contain formulas that point at *other* workbooks. Those are external
links. This document covers how Excel stores them, where `excel_runner` gets it wrong, and
what to do about it.

Everything is labelled:

- **FACT** — measured directly. Real Excel, by hand or by script, results dumped from the saved
  file.

There are no unverified claims left in this document. The one that used to be here (U1) was
tested on 2026-09-04; the result is in Part 2.

---

# Part 1 — How Excel stores links (FACT)

Measured with a test workbook containing five links, one to each of: a file in the same
folder, a file elsewhere on the same drive, a file in `C:\temp`, a file on a network share
(`\\server\share\...`), and a file on a mapped network drive (`Y:`). Each was opened and saved
by hand in real Excel from several different locations, and the saved file's contents dumped
each time.

### F1 — Excel stores either a relative or an absolute path, and picks fresh on every save

The rule is simply: **is the linked file on the same drive letter as the workbook?**

- **Same drive** → Excel stores a *relative* path, like `..\..\data\prices.xlsx`.
- **Different drive, or a network path** → Excel stores an *absolute* path, like
  `file:///Y:\prices.xlsx`.

It does not matter how the link was originally typed, or what was stored last time. Excel
decides again from scratch every single time you save.

### F2 — Relative paths are recalculated correctly whenever the workbook moves

Move the workbook to a different folder on the same drive, open it, save it, and Excel
rewrites the relative path so it still points at the right file. This works reliably.

### F3 — Absolute paths are never revisited

Once Excel has stored an absolute path, it leaves it exactly as it is on subsequent saves. It
does not re-check it or re-derive it, even after the workbook moves.

### F4 — Moving the workbook to another drive flips every link

We moved the linking workbook onto the `Y:` drive and saved. Every link to a `C:` drive file
switched from relative to absolute. The one link that pointed at a file already on `Y:`
switched the other way, from absolute to relative. This follows directly from F1.

### F5 — Some links are stored twice

When the linked file is on the same drive but several folders away, Excel writes **two**
entries for that one link: the relative path, and a second absolute-style path with the drive
letter stripped off:

```
Target="../../temp/prices.xlsx"
Target="/temp/prices.xlsx"
```

This only happens for same-drive links that are some distance away. Links to another drive or
a network share always get exactly one entry.

### F6 — openpyxl throws away the second entry, and the file will not open afterwards

openpyxl can only hold one entry per link. When it reads a file that has two, it keeps the
first and discards the other — immediately on load, before you do anything.

Measured, with zero cell edits, just load and save:

| File | Opens in Excel? |
|------|-----------------|
| No links, openpyxl round-trip | Yes |
| Has links, not touched by openpyxl | Yes |
| Has links, openpyxl round-trip | **No — Excel refuses to open it** |

Excel fails with `Open method of Workbooks class failed`. This is not a broken link that still
lets you use the rest of the workbook. It is a file Excel will not load at all.

And it is not recoverable. The discarded entry is not recorded anywhere else in the file, so
there is nothing left to repair it from.

Note which entry survives: openpyxl keeps the **first** one, which is the *relative* path. The
correct absolute one is the one thrown away.

### F7 — openpyxl never recalculates paths

openpyxl writes the stored path back exactly as it read it. It has no equivalent of F2. So a
workbook that openpyxl saves while it is sitting in a temporary folder keeps paths that were
calculated for that temporary folder.

### F8 — Copying a file does not change its links

A plain file copy (`shutil.copy2`, Windows Explorer, anything) leaves the stored link paths
byte-for-byte identical. Only Excel saving the file, or a library rewriting it, changes them.

### F9 — What the formula bar shows is not always what is stored

Normally the formula bar shows the real full path. The one exception: if the *linked* workbook
happens to also be open in Excel at the same time, the formula bar drops the path entirely and
shows just `[prices.xlsx]Sheet1!A1`. That is display only. Saving in that state still writes
the original full path to the file.

### F10 — A moved or renamed target just breaks the link

Excel does not search for a plausible replacement. If the stored path no longer matches a real
file, the link is broken and stays broken.

### F11 — Excel renumbers links internally

The internal numbering of link entries can change when Excel saves. Any tool that inspects
link data must match links up by the reference in the formula, not by file order, or it will
report the wrong link.

---

# Part 2 — Our commit step, checked

### F12 — The commit step stores a path calculated for the scratch folder, and that is fine

When `excel_runner` finishes with a workbook it does this:

1. The workbook is sitting in the scratch folder: `scratch\working\report.xlsx`.
2. We point its link back at the real target file and tell Excel to save. Excel is saving a
   file that lives in the scratch folder, so by F1 it calculates the path from there.
3. We copy the file from the scratch folder to its real home. By F8 the copy changes nothing.

That looked like a bug. It was tested by reproducing the exact sequence.

**What is stored after step 2:**

```
Target="../../../../demos/showcase/target.xlsx"                    <- wrong once moved
Target="/Dev/projects/excel-runner/demos/showcase/target.xlsx"     <- correct
```

The relative path is indeed wrong for the file's final home. But Excel also wrote the absolute
fallback (F5), and **Excel uses the absolute one**. The committed file resolved correctly and
read the right value.

This was pushed harder: a decoy file was planted at exactly the location the wrong relative
path points to. Excel still used the absolute path and read the correct file. It does not
silently prefer the relative path.

**Conclusion: the commit step is not broken.** No fix needed.

**The one thing to keep in mind:** this works only because both stored paths survive. Anything
that drops one of them leaves the wrong one behind — which is exactly what openpyxl does
(F6).

---

# Part 3 — The rule

> **A workbook that contains links is never saved by openpyxl. Excel saves it, or nothing
> does.**

This is settled, not a topic for discussion. F6 measured it: an openpyxl save leaves a
workbook Excel will not open, and there is nothing left in the file to repair it from.

Reading with openpyxl is fine. Reading never saves.

---

# Part 4 — Where our code breaks this rule

### B1 — We save after every step, with openpyxl

**Code:** `_save_dirty_staged_sessions` ([engine.py:832](excel_runner/engine.py#L832)), called
by `checkpoint()` after every step and again when committing.

Any workbook that has been written to gets saved with openpyxl at the end of that step. If it
has links, they are destroyed then and there. Nothing in this code looks at whether the
workbook has links.

Workbooks that are only *read* are never saved here, so they are safe — but only by accident.
Nothing is checking.

**If we don't fix it:** any write to a link-bearing workbook destroys its links, silently, on
the very next step.

### B2 — We save with openpyxl immediately before handing the file to Excel

**Code:** `_switch_backend` ([engine.py:696](excel_runner/engine.py#L696)).

To repoint a link we need Excel, so we switch the workbook from openpyxl to Excel. That switch
saves the file with openpyxl first, then closes and reopens it in Excel.

So if the workbook has unsaved changes at that moment, we destroy its link on the way to
repointing it. The damage happens inside the very code written to handle links.

**If we don't fix it:** link repointing corrupts the link it is repointing.

### B3 — Commit-time path calculation — CHECKED, NOT A PROBLEM

**Code:** `_revert_r4_links_before_commit` ([engine.py:785](excel_runner/engine.py#L785)) then
`ScratchManager.commit` ([engine.py:391](excel_runner/engine.py#L391)).

Tested and cleared — see F12. Excel's absolute fallback covers the wrong relative path. No
change needed.

### B4 — Links in template workbooks are missed on the first run

**Code:** `discover_write_intent_link_graph`
([engine.py:199](excel_runner/engine.py#L199)), the `path.exists()` check.

We scan a workbook's real file for links. A workbook declared `create_if_missing: true` with a
`template:` has no real file yet on its first run, so we skip it. Any link built into the
template is never noticed, never repointed, never restored.

**If we don't fix it:** on a fresh checkout, template links point at whatever path was baked
into the template.

### B5 — Links like `subfolder\prices.xlsx` break silently

**Code:** `classify_link_target` ([engine.py:147](excel_runner/engine.py#L147)) recognises this
kind of link, but `discover_write_intent_link_graph`
([engine.py:231](excel_runner/engine.py#L231)) ignores everything except absolute paths.

Our scratch folder is flat — every workbook is copied into one folder,
`scratch\working\<filename>` ([engine.py:362](excel_runner/engine.py#L362)). A link that is
just a filename still works there, because both files land side by side. A link with a
subfolder in it cannot work, because the subfolder does not exist in scratch.

**If we don't fix it:** the link breaks and nobody is told.

### B6 — Nothing warns the user

There is no check anywhere that notices "this workbook has links and we are about to write to
it with openpyxl".

**If we don't fix it:** B1, B2 and B5 all fail silently at runtime instead of stopping the run
with a clear message before anything is touched.

### Also worth noting (separate issue)

Because scratch is flat, two workbooks with the same filename in different real folders would
collide there. Not a link problem as such, but it would break same-folder link resolution too.

---

# Part 5 — What this means for what users can do today

Our Excel/COM code can open, save, close, copy ranges, recalculate, and repoint links. That is
all. Everything else is openpyxl only.

**Safe today — reading never saves:**
`open`, `close`, `read_range`, `read_metadata`, `find_headers_row`, `find_row`, `find_column`,
`find_columns`

**Unsafe today — these save via openpyxl:**
`save`, `write_cell`, `write_range`, `write_row`, `insert_range`, `set_column_width`,
`create_sheet`, `rename_sheet`, `delete_sheet`

**Already safe — these use Excel:**
`copy`, `recalculate`

So, plainly: **a workbook with links can be read, but cannot currently be written to safely.**
There is no Excel-based alternative for any of the nine write actions.

---

# Part 6 — The choice to make

Two real options. Pick one.

### Option A — Refuse the job

Check for links before the run starts. If a workbook has links and the workflow writes to it,
stop with a clear error.

**Good:** small, quick, honest. No more silent corruption. No new Excel dependency.

**Bad:** removes the capability entirely. You cannot write to a linked workbook at all. This
blocks the showcase demo's own scenario.

### Option B — Write through Excel instead

Add Excel/COM versions of the nine write actions. Use them whenever the target workbook has
links.

**Good:** actually fixes it. Excel handles link storage itself, so F1 and F5 stop being our
problem. Fits the existing backend-switching design.

**Bad:** much bigger job — nine new operations plus the routing. Needs Excel installed and
running for workflows that previously needed neither. Slower. Roughly doubles the testing for
every write action.

### Two ideas that were considered and rejected

- **Strip the links out before saving, put them back afterwards.** Requires hand-editing the
  file's internal XML around every save, and it preserves the *old* path rather than
  recalculating it — so it reintroduces F7's problem by design.
- **Let openpyxl edit, then re-save through Excel.** The openpyxl save still happens first, and
  that is the step that does the damage.

---

# Part 7 — The plan, in order

| # | Do this | Why now | If we skip it |
|---|---------|---------|---------------|
| 1 | Add the check from B6 — refuse to write to a linked workbook | Turns silent corruption into a clear error before anything is opened. Cheap, and doesn't depend on the A/B decision | B1, B2 and B5 keep producing workbooks Excel cannot open |
| 2 | Fix B2 — never save with openpyxl on the way into Excel | Very small fix, and it's inside the link-handling code itself | Repointing keeps destroying the link it repoints |
| 3 | Fix B4 — fall back to the template file when scanning | Small, self-contained, easy to test on its own | Template links stay invisible on first run |
| 4 | Reject subfolder links (B5) | Same mechanism as step 1. Closes the last silent failure | Broken links with no error |
| 5 | Decide Option A or Option B, then build it | Either restores writing to linked workbooks, or formally drops it | Step 1's check blocks it with no way forward |

Steps 1 to 4 are independent of each other and of the A/B decision. Step 5 is the real design
work.

B3 needed no step — it was tested and cleared (F12).

---

## Where the evidence lives

The measurements behind Part 1, and the earlier working-out that led to them, are in
`docs/r4_link_discovery_gaps_proposal.md`. That document is the audit trail. This one is the
conclusion.
