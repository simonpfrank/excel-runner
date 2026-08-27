import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe5")
(tmp / "orig").mkdir(parents=True, exist_ok=True)
(tmp / "other").mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
try:
    a = app.books.add()
    a.sheets[0].range("A1").value = 10
    a.save(str(tmp / "orig" / "A.xlsx"))

    b = app.books.add()
    b.sheets[0].range("A1").formula = "='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "orig" / "B.xlsx"))
    print("B.A1 (A open from orig):", b.sheets[0].range("A1").value)

    # Force B's stored link to a path that will NEVER actually exist - simulating a
    # UNC/absolute path stored in the file that points somewhere other than where
    # we're actually going to open A from.
    fake_unc_like_path = str((tmp / "nonexistent_share" / "A.xlsx").resolve())
    b.api.ChangeLink(Name="A.xlsx", NewName=fake_unc_like_path, Type=1)
    print(
        "B.A1 after ChangeLink to a path that will never exist:",
        b.sheets[0].range("A1").value,
    )
    a.close()  # close A so it's no longer open anywhere
    print(
        "B.A1 with A fully closed, link points nowhere real:",
        b.sheets[0].range("A1").value,
    )

    # Now re-open A, but from a COMPLETELY DIFFERENT location than either the
    # original save path or the fake stored link path.
    a2 = app.books.open(
        str(
            tmp / "other" / "A.xlsx"
            if (tmp / "other" / "A.xlsx").exists()
            else tmp / "orig" / "A.xlsx"
        )
    )
    print("B LinkSources after A reopened elsewhere:", list(b.api.LinkSources(1) or []))
    print(
        "B.A1 after A (same filename) reopened from a different folder than the stored link path:",
        b.sheets[0].range("A1").value,
    )
    app.calculate()
    print("B.A1 after forcing app.calculate():", b.sheets[0].range("A1").value)
    b.api.Application.CalculateFull()
    print("B.A1 after CalculateFull():", b.sheets[0].range("A1").value)
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
