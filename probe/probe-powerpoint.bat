@echo off
REM Probe wrapper for PowerPoint - launches PowerPoint with WebView2 debug port 9224.
REM Use this to check whether the Claude add-in loads inside PowerPoint and what
REM URL/DOM it exposes. Does NOT run the injector.
REM
REM Usage:
REM   1. Close all PowerPoint windows.
REM   2. Run this .bat.
REM   3. In PowerPoint, open the Claude add-in pane.
REM   4. In another terminal, run: node probe.js 9224
REM
REM Port choice: 9222=Word, 9223=Excel, 9224=PowerPoint.

setlocal EnableDelayedExpansion

set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9224"

set "PPT=C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%PPT%" set "PPT=C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%PPT%" (
    echo [ERROR] POWERPNT.EXE not found in standard locations.
    exit /b 1
)

echo Launching PowerPoint with WebView2 debug port 9224...
echo Once PowerPoint is open, activate the Claude add-in, then run: node probe.js 9224
start "" "%PPT%"

endlocal
exit /b 0
