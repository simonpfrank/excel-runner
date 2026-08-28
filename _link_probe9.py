"""Full R4 cycle, end to end, as recalculate would actually perform it:

  1. B links (absolute path) to A_real (exists, stale value).
  2. ChangeLink(B's link -> A_scratch's absolute path)   [A_scratch is CLOSED on disk]
  3. UpdateLink(B's link, Type=1)                        [force fresh read from disk]
  4. Recalculate B
  5. ChangeLink(B's link -> A_real's absolute path)      [A_real exists, stale]
  6. Do NOT UpdateLink again
  7. Save B, close B
  8. Fresh reopen (new App instance, update_links=False) - what actually persisted?

This is the thing that must be proven, not just each ChangeLink/UpdateLink fact in isolation.
"""

import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe9")
tmp.mkdir(parents=True, exist_ok=True)

a_real_path = str((tmp / "A_real.xlsx").resolve())
a_scratch_path = str((tmp / "A_scratch.xlsx").resolve())
b_path = str((tmp / "B.xlsx").resolve())

# --- Set up A_real (stale, pre-run data) ---
app1 = xw.App(visible=True)
app1.display_alerts = False
try:
    a_real = app1.books.add()
    a_real.sheets[0].range("A1").value = 999
    a_real.save(a_real_path)
    a_real.close()

    # --- Set up A_scratch (fresh, this-run data), then CLOSE it ---
    a_scratch = app1.books.add()
    a_scratch.sheets[0].range("A1").value = 42
    a_scratch.save(a_scratch_path)
    a_scratch.close()

    # --- Set up B, initially linked to A_real (the "before this run touched anything" state) ---
    b = app1.books.add()
    b.sheets[0].range("A1").formula = f"='[A_real.xlsx]Sheet1'!A1*2"
    b.save(b_path)
    b.close()
finally:
    app1.quit()

# --- Now simulate `recalculate` running on B, in a *separate* app instance ---
app2 = xw.App(visible=True)
app2.display_alerts = False
app2.api.AskToUpdateLinks = False
try:
    b = app2.books.open(b_path, update_links=False)
    print("B.A1 on open (linked to real, stale):", b.sheets[0].range("A1").value)
    print("LinkSources before:", list(b.api.LinkSources(1) or []))

    # Step 2: repoint to scratch (A_scratch is closed on disk right now)
    b.api.ChangeLink(Name=a_real_path, NewName=a_scratch_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink -> scratch:", b.sheets[0].range("A1").value
    )

    # Step 3: force fresh read from the (closed) scratch file
    b.api.UpdateLink(Name=a_scratch_path, Type=1)
    print("B.A1 after UpdateLink from scratch:", b.sheets[0].range("A1").value)

    # Step 4: recalc
    app2.calculate()
    print("B.A1 after recalc:", b.sheets[0].range("A1").value)

    # Step 5: repoint back to the real (existing, stale) path
    b.api.ChangeLink(Name=a_scratch_path, NewName=a_real_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink -> real (no UpdateLink after):",
        b.sheets[0].range("A1").value,
    )

    # Step 7: save + close, no further UpdateLink
    b.save()
    b.close()
finally:
    app2.quit()

# --- Step 8: fresh reopen, totally separate app instance, to see what actually persisted ---
# Left open (not quit) so the final state can be inspected visually.
app3 = xw.App(visible=True)
app3.display_alerts = False
b2 = app3.books.open(b_path, update_links=False)
print("B.A1 on fresh reopen (persisted value):", b2.sheets[0].range("A1").value)
print("LinkSources on fresh reopen:", list(b2.api.LinkSources(1) or []))
print("DONE - B.xlsx left open in Excel for inspection.")
