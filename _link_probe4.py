import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe4")
tmp.mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
app.calculation = "manual"
try:
    a = app.books.add()
    a.sheets[0].range("A1").value = 10
    a.save(str(tmp / "A.xlsx"))

    b = app.books.add()
    b.sheets[0].range("A1").formula = "='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "B.xlsx"))
    print("calc mode:", app.calculation)
    print("B.A1 initial:", b.sheets[0].range("A1").value)

    fake_real_path = str((tmp / "A_real_name.xlsx").resolve())
    b.api.ChangeLink(Name="A.xlsx", NewName=fake_real_path, Type=1)
    print(
        "B.A1 immediately after ChangeLink (manual calc mode):",
        b.sheets[0].range("A1").value,
    )

    # also check an unrelated plain local cell to see if ChangeLink triggered a full recalc
    b.sheets[0].range("B1").formula = "=1+1"
    print(
        "B.B1 (unrelated formula) after ChangeLink, before any calc:",
        b.sheets[0].range("B1").value,
    )
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
