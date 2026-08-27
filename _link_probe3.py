import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe3")
tmp.mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
try:
    a = app.books.add()
    a.sheets[0].range("A1").value = 10
    a.save(str(tmp / "A.xlsx"))

    b = app.books.add()
    b.sheets[0].range("A1").formula = "='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "B.xlsx"))
    print("B.A1 initial (A open, live):", b.sheets[0].range("A1").value)

    a.sheets[0].range("A1").value = 50
    print("B.A1 after A edit, no calc call:", b.sheets[0].range("A1").value)

    # Repoint the link to a path that is NOT open / doesn't exist yet - simulating
    # "point back to the real committed path before A has actually been saved there".
    fake_real_path = str((tmp / "A_real_name.xlsx").resolve())
    print("B LinkSources before repoint:", list(b.api.LinkSources(1) or []))
    b.api.ChangeLink(Name="A.xlsx", NewName=fake_real_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink to a non-open/non-existent path:",
        b.sheets[0].range("A1").value,
    )
    print("B LinkSources after repoint:", list(b.api.LinkSources(1) or []))

    b.save()
    print("B.A1 after save:", b.sheets[0].range("A1").value)
    b.close()

    # Reopen B alone, with neither A nor the fake real path open/existing, to see what
    # the saved cached value looks like.
    b2 = app.books.open(str(tmp / "B.xlsx"), update_links=False)
    print(
        "B.A1 on fresh reopen (fake real path still doesn't exist):",
        b2.sheets[0].range("A1").value,
    )
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
