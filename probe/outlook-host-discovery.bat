@echo off
REM Outlook host discovery probe (M0 for v0.3.0).
REM Launches Outlook classic (OUTLOOK.EXE) with --remote-debugging-port=0
REM so its WebView2 picks a free port, then prints next-step instructions.
REM Does NOT modify HKCU. Env var lives only in this process.
REM
REM Usage:
REM   1. Close ALL Outlook windows (classic AND New Outlook) first.
REM   2. Run this .bat.
REM   3. In Outlook: open a mail, click the Apps / Add-ins menu, open Claude.
REM   4. In another terminal: node outlook-host-discovery.js
REM
REM If New Outlook (olk.exe) is also installed and you want to probe it later,
REM the env var route likely does NOT propagate into its Appx container -
REM see Q5 of the probe output for the evidence-based answer.

setlocal EnableDelayedExpansion

REM Pre-flight: refuse if any Outlook flavor is already running. An already-running
REM Outlook will not pick up the new env var and the probe will report false negatives.
for %%E in (OUTLOOK.EXE olk.exe) do (
    tasklist /FI "IMAGENAME eq %%E" 2>NUL | find /I "%%E" >NUL
    if not errorlevel 1 (
        echo [ERROR] %%E is already running. Close it before running this probe.
        echo         Otherwise the dynamic port flag will not take effect.
        exit /b 1
    )
)

set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Outlook classic Click-to-Run paths (most common modern installs):
set "OUTLOOK=C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not exist "%OUTLOOK%" set "OUTLOOK=C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
REM MSI fallbacks (older deployments, some volume-licensed):
if not exist "%OUTLOOK%" set "OUTLOOK=C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
if not exist "%OUTLOOK%" set "OUTLOOK=C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"

if not exist "%OUTLOOK%" (
    echo [ERROR] OUTLOOK.EXE not found in any standard Office16 path.
    echo         If you only have New Outlook ^(Appx^), this probe path does not apply.
    echo         See Q5 in outlook-host-discovery.js for the New Outlook story.
    exit /b 1
)

echo Launching classic Outlook with --remote-debugging-port=0 ...
echo   %OUTLOOK%
start "" "%OUTLOOK%"

echo.
echo Outlook is starting. Now do this manually:
echo   1. Open a mail message (or any item that exposes the reading pane).
echo   2. Click the Apps / Add-ins button and open Claude.
echo   3. Wait a few seconds for the panel to load.
echo.
echo Then, in another terminal:
echo   node outlook-host-discovery.js
echo.
echo Optional: to answer Q6 (shared host pool with Word/Excel/PowerPoint),
echo also launch one of those via probe\launch-office-dynamic.bat in advance
echo and open Claude there too, then run the .js probe.

endlocal
exit /b 0
