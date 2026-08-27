import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe6b")
tmp.mkdir(parents=True, exist_ok=True)

# --- App #1: build A and B together, establish the link, then close everything. ---
app1 = xw.App(visible=False)
app1.display_alerts = False
app1.api.AskToUpdateLinks = False
try:
    a = app1.books.add()
    a.sheets[0].range("A1").value = 10
    a.save(str(tmp / "A.xlsx"))

    b = app1.books.add()
    b.sheets[0].range("A1").formula = "='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "B.xlsx"))
    print("B.A1 (built together):", b.sheets[0].range("A1").value)
finally:
    for bk in list(app1.books):
        bk.close()
    app1.quit()

# --- App #2: a totally separate Excel process modifies A.xlsx and closes. B is never
#     opened here at all. ---
app2 = xw.App(visible=False)
app2.display_alerts = False
app2.api.AskToUpdateLinks = False
try:
    a2 = app2.books.open(str(tmp / "A.xlsx"))
    a2.sheets[0].range("A1").value = 99
    a2.save()
    a2.close()
finally:
    app2.quit()

# --- App #3: a fresh, third Excel process opens ONLY B.xlsx. A.xlsx is closed and not
#     open anywhere on the system at this point. ---
app3 = xw.App(visible=False)
app3.display_alerts = False
app3.api.AskToUpdateLinks = False
try:
    b3 = app3.books.open(str(tmp / "B.xlsx"), update_links=False)
    print(
        "B.A1 on fresh open, A fully closed elsewhere:", b3.sheets[0].range("A1").value
    )
    link_name = list(b3.api.LinkSources(1) or [])[0]
    print("Link name:", link_name)
    b3.api.UpdateLink(Name=link_name, Type=1)
    print(
        "B.A1 after UpdateLink(), A never opened in this process at all:",
        b3.sheets[0].range("A1").value,
    )
finally:
    for bk in list(app3.books):
        bk.close()
    app3.quit()
print("DONE")
