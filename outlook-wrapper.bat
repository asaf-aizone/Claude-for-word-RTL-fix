@echo off
REM Claude for Office RTL - opt-in wrapper for classic Outlook.
REM
REM Unlike word/excel/powerpoint wrappers, this one writes a per-launch opt-in
REM flag (%TEMP%\claude-office-rtl.outlook-optin) before launching Outlook. The
REM injector reads that flag on each tick and only then attaches to the Outlook
REM CDP target. Without the flag the M1a gate in scripts/inject.js refuses to
REM attach. See docs/OUTLOOK-EXPANSION-PLAN.md section 4.1 and probe/README.md
REM "silent CDP attach" for the M0 evidence that motivated this opt-in model.
REM
REM Usage:
REM   1. Close any running Outlook flavor. The WebView2 debug env var is
REM      inherited at process start time, so an already-live Outlook will not
REM      pick it up and the injector would never see a debug port.
REM   2. Run this wrapper. It writes the flag, ensures the injector is up,
REM      then launches Outlook.
REM   3. Open mail, open the Claude add-in. RTL should appear in the panel.
REM
REM Error blocks are implemented via goto labels rather than inline () blocks
REM because cmd's parser treats stray parens inside echo and REM text as block
REM delimiters, which can break the surrounding if/for parens block.

setlocal EnableDelayedExpansion

REM Pre-flight: refuse if any Outlook flavor is already running. An already-running
REM Outlook will not inherit WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS, so the wrapper
REM would silently fail to enable RTL.
tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>nul | find /I "OUTLOOK.EXE" >nul
if not errorlevel 1 goto :err_outlook_running

tasklist /FI "IMAGENAME eq olk.exe" 2>nul | find /I "olk.exe" >nul
if not errorlevel 1 goto :err_olk_running

REM Set the WebView2 debug flag in this process scope only. Inherited by the
REM OUTLOOK.EXE child started below via "start"; not inherited by Word, Excel,
REM PowerPoint, Edge, Teams, or any other WebView2 host running under separate
REM process trees. Port=0 picks a free port at bind time, matching the v0.2.x
REM dynamic-ports model.
set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0"

REM Locate classic Outlook. Click-to-Run is the common modern install path; the
REM x86 variant catches 32-bit-on-64-bit deployments. MSI fallbacks are not
REM included - they tend to be retired Office versions where the Claude add-in
REM is not supported anyway.
set "OUTLOOK=C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not exist "%OUTLOOK%" set "OUTLOOK=C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not exist "%OUTLOOK%" goto :err_not_installed

REM Anti-duplicate injector logic, identical to word-wrapper.bat. The lock and
REM PID file paths keep the legacy claude-word-rtl prefix so wrappers from
REM different apps coordinate on the same singleton. A stale lock whose PID is
REM dead is cleared so a crash does not lock out the next legitimate launch.
set "LOCK=%TEMP%\claude-word-rtl.lock"
set "PIDFILE=%TEMP%\claude-word-rtl.pid"
set "OPTIN=%TEMP%\claude-office-rtl.outlook-optin"
set "SKIP_INJECTOR="

if exist "%LOCK%" call :check_pid_alive

if defined SKIP_INJECTOR (
    REM Existing injector. Its startup stale-clear ran in the past, so writing
    REM the flag now is safe; the next 2-second tick picks it up and the
    REM wrapper-started Outlook will be attached on first appearance.
    echo %DATE% %TIME%> "%OPTIN%"
) else (
    if not exist "%~dp0scripts\node_modules\chrome-remote-interface" (
        pushd "%~dp0scripts"
        call npm install --silent >nul 2>&1
        popd
    )
    echo %DATE% %TIME%> "%LOCK%"
    wscript "%~dp0inject-hidden.vbs"
    REM Race-avoidance: a freshly-spawned injector clears every opt-in flag
    REM file during synchronous startup. If we wrote the Outlook flag before
    REM that ran, it would be deleted and the next tick would refuse to
    REM attach. 3 seconds is comfortably longer than the injector startup
    REM phase but short enough to not delay the user noticeably.
    timeout /t 3 /nobreak >nul 2>&1
    echo %DATE% %TIME%> "%OPTIN%"
)

REM Launch Outlook. No document argument - Outlook opens to its default inbox.
start "" "%OUTLOOK%"

endlocal
exit /b 0

:check_pid_alive
REM Sub-routine. Sets SKIP_INJECTOR=1 if the PID file points at a live process.
REM Otherwise clears the stale lock and PID files so the caller will spawn a
REM fresh injector.
if not exist "%PIDFILE%" (
    del /q "%LOCK%" >nul 2>&1
    goto :eof
)
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
goto :eof

:err_outlook_running
echo [ERROR] OUTLOOK.EXE is already running. Close it, then re-run this wrapper.
echo         The WebView2 debug env var is inherited at process start time
echo         only, so an already-live Outlook will not pick it up.
endlocal
exit /b 2

:err_olk_running
echo [ERROR] New Outlook olk.exe is running. Close it, then re-run this wrapper.
echo         Even if you do not intend to use New Outlook, its presence usually
echo         means a classic-Outlook launch will collide on shared state.
endlocal
exit /b 2

:err_not_installed
echo [ERROR] OUTLOOK.EXE not found in any standard Office16 path.
echo         Classic Outlook Microsoft 365 is required. New Outlook Appx is
echo         not yet supported by this wrapper.
endlocal
exit /b 3
