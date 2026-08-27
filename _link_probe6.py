import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe6")
tmp.mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
try:
    # Build A and B together once, so the link is established, then close A fully.
    a = app.books.add()
    a.sheets[0].range("A1").value = 10
    a.save(str(tmp / "A.xlsx"))

    b = app.books.add()
    b.sheets[0].range("A1").formula = "='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "B.xlsx"))
    print("B.A1 (A open):", b.sheets[0].range("A1").value)

    a.close()  # A is now fully closed - not open anywhere in this app instance
    print("B.A1 with A closed, no update yet:", b.sheets[0].range("A1").value)

    # Now, separately (as if in a totally different run), change A's value on disk
    # via a brand-new headless open/save, fully independent of B's session.
    a2 = app.books.open(str(tmp / "A.xlsx"))
    a2.sheets[0].range("A1").value = 99
    a2.save()
    a2.close()

    print(
        "B.A1 still stale (A was never reopened in B's session):",
        b.sheets[0].range("A1").value,
    )

    link_name = list(b.api.LinkSources(1) or [])[0]
    print("Link name to use for UpdateLink:", link_name)
    b.api.UpdateLink(Name=link_name, Type=1)
    print(
        "B.A1 after UpdateLink(), A never opened via our own Book object:",
        b.sheets[0].range("A1").value,
    )
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
