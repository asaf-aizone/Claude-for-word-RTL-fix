@echo off
REM Claude for Office RTL - transparent wrapper for PowerPoint.
REM Called via shortcut or file association. Sets the WebView2 debug flag,
REM ensures the hidden injector is running, then launches PowerPoint
REM (optionally with a presentation argument).

setlocal EnableDelayedExpansion

REM Set the WebView2 debug flag in this process scope only. The flag is
REM inherited by PowerPoint (started below via 'start ""') so its WebView2
REM host exposes the Chrome DevTools Protocol on a free dynamic port, but
REM it is NOT inherited by Teams, Outlook, Edge or any other WebView2 host -
REM those run under separate process trees that never see this variable.
REM
REM Port=0 (dynamic) lets each Office process pick its own free port,
REM which is required when more than one Office app runs at the same time.
REM Note: Word and PowerPoint sometimes share a WebView2 host process, so
REM both apps may end up on the same port; the injector handles both shapes
REM via the _host_Info= URL parameter that disambiguates per-target.
set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Locate PowerPoint. Click-to-Run paths first, then MSI fallback.
set "POWERPNT=C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%POWERPNT%" set "POWERPNT=C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
if not exist "%POWERPNT%" set "POWERPNT=C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE"
if not exist "%POWERPNT%" set "POWERPNT=C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE"
if not exist "%POWERPNT%" (
    REM PowerPoint not found - bail silently; user will see the system error anyway
    exit /b 1
)

REM Anti-duplicate injector logic. See excel-wrapper.bat / word-wrapper.bat
REM for the full rationale. Same lock + PID files are shared across all
REM three wrappers; one injector serves all three Office apps.

set "LOCK=%TEMP%\claude-word-rtl.lock"
set "PIDFILE=%TEMP%\claude-word-rtl.pid"
set "SKIP_INJECTOR="

if exist "%LOCK%" (
    if exist "%PIDFILE%" (
        set "PID_ALIVE="
        for /f "usebackq delims=" %%P in ("%PIDFILE%") do (
            tasklist /FI "PID eq %%P" 2>nul | find "%%P" >nul && set "PID_ALIVE=1"
        )
        if defined PID_ALIVE (
            set "SKIP_INJECTOR=1"
        ) else (
            del /q "%LOCK%"    >nul 2>&1
            del /q "%PIDFILE%" >nul 2>&1
        )
    ) else (
        del /q "%LOCK%" >nul 2>&1
    )
)

if not defined SKIP_INJECTOR (
    if not exist "%~dp0scripts\node_modules\chrome-remote-interface" (
        pushd "%~dp0scripts"
        call npm install --silent >nul 2>&1
        popd
    )
    echo %DATE% %TIME%> "%LOCK%"
    wscript "%~dp0inject-hidden.vbs"
)

REM Mutex-settle wait. PowerPoint uses a per-user named mutex for single-
REM instance coordination. The mutex can outlive the visible POWERPNT.EXE
REM process by 1-3 seconds while WebView2 and COM unwind. If the tray's
REM "Connect PowerPoint" flow relaunches before the mutex is released, the
REM new POWERPNT sees an existing instance, hands off to the dying one, and
REM silently exits - looking like "PowerPoint never came back" to the user.
REM Poll tasklist up to ~5s (5 x 1s) for POWERPNT.EXE to disappear,
REM then proceed. Fail-open: if still present after 5s, launch anyway.
REM
REM We use `timeout` not `ping -w 250` for the sleep. Loopback replies in
REM ~0ms and `-w` only controls the reply-timeout, so `ping -n 1 -w 250`
REM returns instantly and a 20-iteration loop would complete in <50ms -
REM effectively no wait at all.
set /a _WAIT_ITER=0
:wait_powerpnt_gone
tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if errorlevel 1 goto :proceed_launch
set /a _WAIT_ITER+=1
if %_WAIT_ITER% GEQ 5 goto :proceed_launch
timeout /t 1 /nobreak >nul 2>&1
goto :wait_powerpnt_gone

:proceed_launch
REM Launch PowerPoint (pass through any presentation argument)
if "%~1"=="" (
    start "" "%POWERPNT%"
) else (
    start "" "%POWERPNT%" "%~1"
)

endlocal
exit /b 0
