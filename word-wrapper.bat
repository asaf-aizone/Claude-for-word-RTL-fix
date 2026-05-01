@echo off
REM Claude for Word RTL - transparent wrapper for Word.
REM Called via shortcut or file association. Sets the WebView2 debug flag,
REM ensures the hidden injector is running, then launches Word (optionally
REM with a document argument).

setlocal EnableDelayedExpansion

REM Set the WebView2 debug flag in this process scope only. The flag is
REM inherited by Word (started below via 'start ""') so its WebView2 host
REM exposes the Chrome DevTools Protocol on a free dynamic port, but it
REM is NOT inherited by Teams, Outlook, Edge or any other WebView2 host -
REM those run under separate process trees that never see this variable.
REM
REM Port=0 (dynamic) was chosen over a fixed port (legacy 9222) so each
REM Office process picks its own free port. Required for v0.2.0+ Excel
REM and PowerPoint support, where multiple Office apps may run together.
set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Locate Word
set "WINWORD=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" set "WINWORD=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" (
    REM Word not found - bail silently; user will see the system error anyway
    exit /b 1
)

REM Anti-duplicate injector logic.
REM
REM We use a lock file at %TEMP%\claude-word-rtl.lock as a marker that an
REM injector has been launched recently. Rationale:
REM   - Port 9222 is the WebView2 debug port owned by Word, not the injector.
REM     A LISTENING 9222 only means Word is up, not that the injector is.
REM   - We cannot reliably identify "our" node.exe among all node processes
REM     on the system without extra tooling.
REM   - The injector is idempotent - re-injecting is harmless but wasteful.
REM     A lock file avoids spawning a second injector when the user opens
REM     additional .docx files while Word is already running.
REM
REM Staleness: if the lock file is older than ~2 minutes (forfiles /m /d -0
REM semantics are coarse, so we use a simpler existence check plus manual
REM cleanup in cleanup.bat). Opening Word after a long pause still re-launches
REM because the injector process exits when Word closes and the lock will be
REM removed by cleanup.bat or manually.

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

REM Launch Word (pass through any file argument)
if "%~1"=="" (
    start "" "%WINWORD%"
) else (
    start "" "%WINWORD%" "%~1"
)

endlocal
exit /b 0
