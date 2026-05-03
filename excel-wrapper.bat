@echo off
REM Claude for Office RTL - transparent wrapper for Excel.
REM Called via shortcut or file association. Sets the WebView2 debug flag,
REM ensures the hidden injector is running, then launches Excel (optionally
REM with a workbook argument).

setlocal EnableDelayedExpansion

REM Set the WebView2 debug flag in this process scope only. The flag is
REM inherited by Excel (started below via 'start ""') so its WebView2 host
REM exposes the Chrome DevTools Protocol on a free dynamic port, but it
REM is NOT inherited by Teams, Outlook, Edge or any other WebView2 host -
REM those run under separate process trees that never see this variable.
REM
REM Port=0 (dynamic) lets each Office process pick its own free port,
REM which is required when more than one Office app runs at the same time
REM (Word + Excel, Excel + PowerPoint, all three together).
set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Locate Excel. Click-to-Run paths first (most modern installs), then MSI fallback.
set "EXCEL=C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
if not exist "%EXCEL%" set "EXCEL=C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"
if not exist "%EXCEL%" (
    REM Excel not found - bail silently; user will see the system error anyway
    exit /b 1
)

REM Anti-duplicate injector logic.
REM
REM Lock + PID files are SHARED across all three Office wrappers (word, excel,
REM powerpoint). One injector process serves all three apps; it discovers any
REM Office WebView2 host via tasklist + netstat in scripts/port-discovery.js.
REM So a Word session that already started the injector means a subsequent
REM Excel launch via this wrapper will skip the injector spawn and just
REM launch Excel directly. The injector picks up the new Excel target on its
REM next 2s tick.
REM
REM File names retain the historical "claude-word-rtl" prefix for backward
REM compatibility with v0.1.x installs that may still have these files in
REM %TEMP% during an upgrade window. The prefix is a name, not a constraint.

set "LOCK=%TEMP%\claude-word-rtl.lock"
set "PIDFILE=%TEMP%\claude-word-rtl.pid"
set "SKIP_INJECTOR="

if exist "%LOCK%" (
    REM Check if OUR injector is actually alive (via PID file, not any node.exe).
    if exist "%PIDFILE%" (
        set "PID_ALIVE="
        for /f "usebackq delims=" %%P in ("%PIDFILE%") do (
            tasklist /FI "PID eq %%P" 2>nul | find "%%P" >nul && set "PID_ALIVE=1"
        )
        if defined PID_ALIVE (
            set "SKIP_INJECTOR=1"
        ) else (
            REM Stale: PID no longer alive
            del /q "%LOCK%"    >nul 2>&1
            del /q "%PIDFILE%" >nul 2>&1
        )
    ) else (
        REM Lock exists but no PID file: stale
        del /q "%LOCK%" >nul 2>&1
    )
)

if not defined SKIP_INJECTOR (
    REM First-run install of dependencies if missing
    if not exist "%~dp0scripts\node_modules\chrome-remote-interface" (
        pushd "%~dp0scripts"
        call npm install --silent >nul 2>&1
        popd
    )
    REM Mark injector as launched, then start it hidden via VBS wrapper
    echo %DATE% %TIME%> "%LOCK%"
    wscript "%~dp0inject-hidden.vbs"
)

REM Launch Excel (pass through any workbook argument)
if "%~1"=="" (
    start "" "%EXCEL%"
) else (
    start "" "%EXCEL%" "%~1"
)

endlocal
exit /b 0
