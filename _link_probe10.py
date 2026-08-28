"""Does ChangeLink's instant auto-refresh (proven in probe9) still happen:
  (a) with NO explicit calculate()/UpdateLink call at all afterward, and
  (b) under manual calculation mode?

This is the exact commit-time revert step: A_real has just been overwritten with the fresh
scratch content (simulating "commit A first"), then B's link is switched from A_scratch back
to A_real, and B is saved immediately - no calculate(), no UpdateLink. If the cached value is
still correct after that, the new commit-ordering design is safe without any extra recalc step.
"""

import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe10")
tmp.mkdir(parents=True, exist_ok=True)

a_real_path = str((tmp / "A_real.xlsx").resolve())
a_scratch_path = str((tmp / "A_scratch.xlsx").resolve())
b_path = str((tmp / "B.xlsx").resolve())

xlCalculationManual = -4135
xlCalculationAutomatic = -4105

# --- Set up: A_real starts stale (999), A_scratch has the fresh value (42) ---
app1 = xw.App(visible=False)
app1.display_alerts = False
try:
    a_real = app1.books.add()
    a_real.sheets[0].range("A1").value = 999
    a_real.save(a_real_path)
    a_real.close()

    a_scratch = app1.books.add()
    a_scratch.sheets[0].range("A1").value = 42
    a_scratch.save(a_scratch_path)
    a_scratch.close()

    # B links to A_scratch already (as it would mid-run, per the new staging design)
    b = app1.books.add()
    b.sheets[0].range("A1").formula = "='[A_scratch.xlsx]Sheet1'!A1*2"
    b.save(b_path)
    b.close()
finally:
    app1.quit()

# --- Simulate: "A has just been committed" -> overwrite A_real with A_scratch's content ---
import shutil

shutil.copy2(a_scratch_path, a_real_path)

# --- Now open B in a fresh app, set MANUAL calc, revert its link, save - no calculate() call ---
app2 = xw.App(visible=False)
app2.display_alerts = False
app2.api.AskToUpdateLinks = False
try:
    app2.api.Calculation = xlCalculationManual
    b = app2.books.open(b_path, update_links=False)
    print("Calc mode is manual:", app2.api.Calculation == xlCalculationManual)
    print(
        "B.A1 on open (linked to scratch, mid-run state):",
        b.sheets[0].range("A1").value,
    )

    # The commit-time revert: ChangeLink back to the NOW-fresh real path. No UpdateLink, no calculate().
    b.api.ChangeLink(Name=a_scratch_path, NewName=a_real_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink -> real (now fresh), still manual calc, no recalc call:",
        b.sheets[0].range("A1").value,
    )

    b.save()
    b.close()
finally:
    app2.quit()

# --- Fresh reopen, separate app, to see what actually persisted ---
app3 = xw.App(visible=False)
app3.display_alerts = False
try:
    b2 = app3.books.open(b_path, update_links=False)
    print("B.A1 on fresh reopen (persisted value):", b2.sheets[0].range("A1").value)
    print("LinkSources on fresh reopen:", list(b2.api.LinkSources(1) or []))
finally:
    app3.quit()

print("DONE")
