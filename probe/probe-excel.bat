@echo off
REM Probe wrapper for Excel - launches Excel with WebView2 debug port 9223.
REM Use this to check whether the Claude add-in loads inside Excel and what
REM URL/DOM it exposes. Does NOT run the injector.
REM
REM Usage:
REM   1. Close all Excel windows.
REM   2. Run this .bat.
REM   3. In Excel, open the Claude add-in pane (Home > Claude, or wherever it lives).
REM   4. In another terminal, run: node probe.js 9223
REM
REM Port choice: 9222 is used by Word. Excel gets 9223 so both can coexist.

setlocal EnableDelayedExpansion

set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9223"

set "EXCEL=C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" (
    echo [ERROR] EXCEL.EXE not found in standard locations.
    exit /b 1
)

echo Launching Excel with WebView2 debug port 9223...
echo Once Excel is open, activate the Claude add-in, then run: node probe.js 9223
start "" "%EXCEL%"

endlocal
exit /b 0
