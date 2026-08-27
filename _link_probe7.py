import xlwings as xw
from pathlib import Path

tmp = Path("excel_runner_runs/_link_probe7")
(tmp / "scratch").mkdir(parents=True, exist_ok=True)
(tmp / "real").mkdir(parents=True, exist_ok=True)

app = xw.App(visible=False)
app.display_alerts = False
app.api.AskToUpdateLinks = False
try:
    # Source built and edited at its "scratch" location, with fresh data.
    a_scratch = app.books.add()
    a_scratch.sheets[0].range("A1").value = 10
    a_scratch.save(str(tmp / "scratch" / "A.xlsx"))

    # Parent links to the source at its scratch location (simulating mid-run state).
    b = app.books.add()
    b.sheets[0].range("A1").formula = f"='[A.xlsx]Sheet1'!A1*2"
    b.save(str(tmp / "scratch" / "B.xlsx"))
    print(
        "B.A1 (linked to scratch A, both open together):", b.sheets[0].range("A1").value
    )

    # Simulate: source gets its final edit, recalculated, saved - fresh value now 10*... say 25.
    a_scratch.sheets[0].range("A1").value = 25
    print("B.A1 live after source edit:", b.sheets[0].range("A1").value)
    a_scratch.save()

    # Commit: close source, "rename" scratch -> real (copy since same filesystem/simplicity).
    a_scratch.close()
    real_path = tmp / "real" / "A_real_name.xlsx"
    import shutil

    shutil.copy2(tmp / "scratch" / "A.xlsx", real_path)

    # Now fix the parent's link: ChangeLink to the real path (blanks), then UpdateLink
    # immediately to pull the value back from the (closed, but now up-to-date) real file.
    link_name_before = list(b.api.LinkSources(1) or [])[0]
    print("Link name before fix:", link_name_before)
    b.api.ChangeLink(Name=link_name_before, NewName=str(real_path.resolve()), Type=1)
    print(
        "B.A1 immediately after ChangeLink to real path (expect blank):",
        b.sheets[0].range("A1").value,
    )

    link_name_after = list(b.api.LinkSources(1) or [])[0]
    b.api.UpdateLink(Name=link_name_after, Type=1)
    print(
        "B.A1 after UpdateLink() pulling from the closed, now-committed real file:",
        b.sheets[0].range("A1").value,
    )

    b.save()
    b.close()

    # Reopen fresh, with nothing else open, to confirm it's genuinely persisted.
    b2 = app.books.open(str(tmp / "scratch" / "B.xlsx"), update_links=False)
    print("B.A1 on totally fresh reopen:", b2.sheets[0].range("A1").value)
    print("B LinkSources on fresh reopen:", list(b2.api.LinkSources(1) or []))
finally:
    for bk in list(app.books):
        bk.close()
    app.quit()
print("DONE")
