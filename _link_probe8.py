import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe8")
tmp.mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
try:
    # The "real" A already exists (pre-run), with OLD stale data.
    a_real = app.books.add()
    a_real.sheets[0].range("A1").value = 999  # stale, pre-run value
    a_real.save(str(tmp / "A_real_name.xlsx"))
    a_real.close()

    # A scratch copy of A gets the fresh, this-run edit.
    a_scratch = app.books.add()
    a_scratch.sheets[0].range("A1").value = 25
    a_scratch.save(str(tmp / "A_scratch.xlsx"))

    # Parent, currently linked to scratch A (mid-run state), gets the live fresh value.
    b = app.books.add()
    b.sheets[0].range("A1").formula = "='[A_scratch.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "B.xlsx"))
    print("B.A1 linked to scratch (fresh, live):", b.sheets[0].range("A1").value)

    # Now: "put it back" - ChangeLink to the REAL path, which exists but has STALE data.
    # Do NOT call UpdateLink afterward - just change the pointer and save.
    real_path = str((tmp / "A_real_name.xlsx").resolve())
    b.api.ChangeLink(Name="A_scratch.xlsx", NewName=real_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink to the (existing, stale) real path:",
        b.sheets[0].range("A1").value,
    )

    b.save()
    b.close()

    # Fresh reopen, nothing else open - what actually got persisted?
    b2 = app.books.open(str(tmp / "B.xlsx"), update_links=False)
    print("B.A1 on fresh reopen:", b2.sheets[0].range("A1").value)
    print("B LinkSources on fresh reopen:", list(b2.api.LinkSources(1) or []))
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
