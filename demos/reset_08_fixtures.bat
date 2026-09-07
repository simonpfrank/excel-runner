@echo off
REM ---------------------------------------------------------------------------
REM Resets demo 08's working copies from their originals\ fixtures.
REM
REM RUN THIS BEFORE demos\08_full_showcase.yaml.
REM
REM Why it's needed: 08 writes to its working copies and commits them, so a
REM second run starts from the previous run's output instead of a clean fixture
REM and fails. This puts the three working copies back to their originals.
REM
REM If this repo folder is ever MOVED or CLONED elsewhere, the linked workbook
REM fixture goes stale first - it stores an absolute path to this checkout's
REM catalog.xlsx. Regenerate it with:
REM     .venv\Scripts\python demos\08_full_showcase\generate_linked_workbook_fixture.py
REM ---------------------------------------------------------------------------

setlocal
set HERE=%~dp0
set DEMO=%HERE%08_full_showcase
set RUNDIR=%HERE%..\excel_runner_runs\08_full_showcase

copy /Y "%DEMO%\originals\catalog.xlsx" "%DEMO%\catalog.xlsx" >nul
if errorlevel 1 goto failed

copy /Y "%DEMO%\originals\report.xlsx" "%DEMO%\report.xlsx" >nul
if errorlevel 1 goto failed

copy /Y "%DEMO%\linked_workbook\originals\linked.xlsx" "%DEMO%\linked_workbook\linked.xlsx" >nul
if errorlevel 1 goto failed

if exist "%RUNDIR%" rmdir /S /Q "%RUNDIR%"

echo Demo 08 fixtures reset. Now run:
echo     .venv\Scripts\python -m excel_runner demos\08_full_showcase.yaml
exit /b 0

:failed
echo.
echo FAILED to reset demo 08 fixtures - check that these exist:
echo     %DEMO%\originals\catalog.xlsx
echo     %DEMO%\originals\report.xlsx
echo     %DEMO%\linked_workbook\originals\linked.xlsx
exit /b 1
