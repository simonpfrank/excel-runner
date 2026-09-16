@echo off
REM Creates dist\excel-runner-unify.zip for source-directory deployment in Unify.
REM Extract the archive contents into the target working directory, then run:
REM     python run_excel_runner.py unify_smoke_test.yaml

setlocal
set "ROOT=%~dp0"
set "DIST=%ROOT%dist"
set "STAGING=%DIST%\excel-runner-unify"
set "ARCHIVE=%DIST%\excel-runner-unify.zip"

if not exist "%ROOT%excel_runner\__main__.py" goto missing_source
if not exist "%ROOT%run_excel_runner.py" goto missing_source
if not exist "%ROOT%requirements.txt" goto missing_source
if not exist "%ROOT%unify_smoke_test.yaml" goto missing_source

if exist "%STAGING%" rmdir /S /Q "%STAGING%"
if exist "%ARCHIVE%" del /Q "%ARCHIVE%"
mkdir "%STAGING%"
if errorlevel 1 goto failed

robocopy "%ROOT%excel_runner" "%STAGING%\excel_runner" /E /XD __pycache__ /NJH /NJS /NDL /NFL >nul
if errorlevel 8 goto failed
copy /Y "%ROOT%run_excel_runner.py" "%STAGING%\run_excel_runner.py" >nul
if errorlevel 1 goto failed
copy /Y "%ROOT%requirements.txt" "%STAGING%\requirements.txt" >nul
if errorlevel 1 goto failed
copy /Y "%ROOT%unify_smoke_test.yaml" "%STAGING%\unify_smoke_test.yaml" >nul
if errorlevel 1 goto failed

powershell -NoProfile -Command "Compress-Archive -Path '%STAGING%\*' -DestinationPath '%ARCHIVE%' -Force"
if errorlevel 1 goto failed

echo Created %ARCHIVE%
exit /b 0

:missing_source
echo FAILED: expected deployment source files are missing under %ROOT%
exit /b 1

:failed
echo FAILED: could not create the Unify deployment archive.
exit /b 1