@echo off
REM Launches Word, Excel, and PowerPoint with WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0
REM so each picks a different free port automatically. Use together with
REM dynamic-port-discovery.js to validate the architecture in
REM docs/OFFICE-EXPANSION-PLAN.md section 3.5.
REM
REM Usage:
REM   1. Close ALL Word/Excel/PowerPoint windows first.
REM   2. Run this .bat.
REM   3. In each Office app: open the Claude add-in pane.
REM   4. In another terminal: node dynamic-port-discovery.js
REM
REM This .bat does NOT modify HKCU. The env var is set only for the launched processes.

setlocal EnableDelayedExpansion

REM Pre-flight: refuse to run if any Office app is already open. A running
REM Office process inherits the OLD env, so 'start "" WINWORD.EXE' would
REM reuse it and silently ignore --remote-debugging-port=0.
for %%E in (WINWORD.EXE EXCEL.EXE POWERPNT.EXE) do (
    tasklist /FI "IMAGENAME eq %%E" 2>NUL | find /I "%%E" >NUL
    if not errorlevel 1 (
        echo [ERROR] %%E is already running. Close all Office apps before running this POC.
        echo         Otherwise the dynamic port flag will not take effect.
        exit /b 1
    )
)

set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Click-to-Run paths (most common modern installs):
set "WORD=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WORD%" set "WORD=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
REM MSI fallbacks (older deployments, some volume-licensed):
if not exist "%WORD%" set "WORD=C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"
if not exist "%WORD%" set "WORD=C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"

set "EXCEL=C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"

set "PPT=C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%PPT%" set "PPT=C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%PPT%" set "PPT=C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE"
if not exist "%PPT%" set "PPT=C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE"

set "ANY_LAUNCHED=0"

if exist "%WORD%" (
    echo Launching Word...
    start "" "%WORD%"
    set "ANY_LAUNCHED=1"
) else (
    echo [WARN] WINWORD.EXE not found, skipping Word.
)

if exist "%EXCEL%" (
    echo Launching Excel...
    start "" "%EXCEL%"
    set "ANY_LAUNCHED=1"
) else (
    echo [WARN] EXCEL.EXE not found, skipping Excel.
)

if exist "%PPT%" (
    echo Launching PowerPoint...
    start "" "%PPT%"
    set "ANY_LAUNCHED=1"
) else (
    echo [WARN] POWERPNT.EXE not found, skipping PowerPoint.
)

if "%ANY_LAUNCHED%"=="0" (
    echo [ERROR] No Office app could be launched.
    exit /b 1
)

echo.
echo Office apps launched with --remote-debugging-port=0.
echo Open the Claude add-in in each, then run:
echo   node dynamic-port-discovery.js

endlocal
exit /b 0
