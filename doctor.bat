@echo off
REM Claude for Word RTL - diagnostic script.
REM Runs a series of environment checks and writes a report to doctor.log.
REM Share doctor.log when reporting issues.

setlocal EnableDelayedExpansion

set "HERE=%~dp0"
if "%HERE:~-1%"=="\" set "HERE=%HERE:~0,-1%"
set "LOG=%HERE%\doctor.log"
echo Doctor started %DATE% %TIME% > "%LOG%"

set /a OK_COUNT=0
set /a WARN_COUNT=0
set /a FAIL_COUNT=0

call :log "================================================================"
call :log " Claude for Word RTL - Doctor"
call :log "================================================================"
call :log ""

REM ---------------------------------------------------------------
REM [1] Node.js
REM ---------------------------------------------------------------
where node >nul 2>&1
if errorlevel 1 (
    call :fail "[1] Node.js: not found in PATH"
) else (
    for /f "delims=" %%v in ('node --version 2^>^&1') do set "NODE_VER=%%v"
    call :ok "[1] Node.js: !NODE_VER!"
)

REM ---------------------------------------------------------------
REM [2] npm
REM ---------------------------------------------------------------
where npm >nul 2>&1
if errorlevel 1 (
    call :fail "[2] npm: not found in PATH"
) else (
    for /f "delims=" %%v in ('npm --version 2^>^&1') do set "NPM_VER=%%v"
    call :ok "[2] npm: !NPM_VER!"
)

REM ---------------------------------------------------------------
REM [3] chrome-remote-interface
REM ---------------------------------------------------------------
if exist "%HERE%\scripts\node_modules\chrome-remote-interface" (
    call :ok "[3] chrome-remote-interface: installed"
) else (
    call :fail "[3] chrome-remote-interface: missing - run install.bat or 'npm install' in scripts\"
)

REM ---------------------------------------------------------------
REM [4] WINWORD.EXE location
REM ---------------------------------------------------------------
set "WINWORD1=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
set "WINWORD2=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
set "WINWORD_FOUND="
if exist "%WINWORD1%" set "WINWORD_FOUND=%WINWORD1%"
if not defined WINWORD_FOUND if exist "%WINWORD2%" set "WINWORD_FOUND=%WINWORD2%"
if defined WINWORD_FOUND (
    call :ok "[4] WINWORD.EXE: !WINWORD_FOUND!"
) else (
    call :fail "[4] WINWORD.EXE: not found in default Office 16 locations"
)

REM ---------------------------------------------------------------
REM [5] Word currently running
REM ---------------------------------------------------------------
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 (
    call :ok "[5] Word running: no"
) else (
    call :warn "[5] Word running: yes (close Word for a clean test)"
)

REM ---------------------------------------------------------------
REM [6] Debug port 9222 state
REM ---------------------------------------------------------------
netstat -an | findstr ":9222" >nul 2>&1
if errorlevel 1 (
    call :warn "[6] Port 9222: not bound (Word not launched via wrapper, or not yet ready)"
) else (
    netstat -an | findstr ":9222" | findstr /I "LISTENING" >nul 2>&1
    if errorlevel 1 (
        call :warn "[6] Port 9222: bound but not LISTENING"
    ) else (
        call :ok "[6] Port 9222: LISTENING"
    )
)

REM ---------------------------------------------------------------
REM [7] Injector process (via PID file written by inject.js)
REM ---------------------------------------------------------------
set "INJ_PIDFILE=%TEMP%\claude-word-rtl.pid"
set "INJ_PID="
if exist "%INJ_PIDFILE%" (
    set /p INJ_PID=<"%INJ_PIDFILE%"
)
if defined INJ_PID (
    tasklist /FI "PID eq !INJ_PID!" 2>nul | "%SystemRoot%\System32\find.exe" "!INJ_PID!" >nul
    if errorlevel 1 (
        call :warn "[7] Injector: stale PID file at %INJ_PIDFILE% (node.exe not actually running)"
    ) else (
        call :ok "[7] Injector: running (PID !INJ_PID!)"
    )
) else (
    call :warn "[7] Injector: not running (Connect via tray, or launch word-wrapper.bat)"
)

REM ---------------------------------------------------------------
REM [8] Startup folder entry for the tray
REM ---------------------------------------------------------------
set "STARTUP_LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk"
if exist "%STARTUP_LNK%" (
    call :ok "[8] Startup entry 'Claude for Word RTL Tray': present (tray auto-launches at login)"
) else (
    call :warn "[8] Startup entry 'Claude for Word RTL Tray': missing (run install.bat to create it)"
)

REM ---------------------------------------------------------------
REM [9] Tray process running (via tray PID file)
REM ---------------------------------------------------------------
set "TRAY_PIDFILE=%TEMP%\claude-word-rtl.tray.pid"
set "TRAY_PID="
if exist "%TRAY_PIDFILE%" (
    set /p TRAY_PID=<"%TRAY_PIDFILE%"
)
if defined TRAY_PID (
    tasklist /FI "PID eq !TRAY_PID!" 2>nul | "%SystemRoot%\System32\find.exe" "!TRAY_PID!" >nul
    if errorlevel 1 (
        call :warn "[9] Tray process: stale PID file at !TRAY_PIDFILE! (tray not actually running)"
    ) else (
        call :ok "[9] Tray process: running (PID !TRAY_PID!)"
    )
) else (
    call :warn "[9] Tray process: not running (double-click scripts\start-tray.vbs or log out and back in)"
)

REM ---------------------------------------------------------------
REM [10] WebView2 runtime
REM ---------------------------------------------------------------
set "WV2_FOUND="
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" >nul 2>&1
if not errorlevel 1 set "WV2_FOUND=WOW6432Node"
if not defined WV2_FOUND (
    reg query "HKLM\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" >nul 2>&1
    if not errorlevel 1 set "WV2_FOUND=HKLM\SOFTWARE"
)
if defined WV2_FOUND (
    call :ok "[10] WebView2 runtime: installed (!WV2_FOUND!)"
) else (
    call :fail "[10] WebView2 runtime: not detected - Claude add-in will not render"
)

REM ---------------------------------------------------------------
REM [11] Office version
REM ---------------------------------------------------------------
set "OFFICE_VER="
for /f "tokens=2,*" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\Common\ProductVersion" /v LastProduct 2^>nul ^| find "LastProduct"') do set "OFFICE_VER=%%b"
if defined OFFICE_VER (
    call :ok "[11] Office 16 version: !OFFICE_VER!"
) else (
    call :warn "[11] Office 16 version: could not read from registry"
)

REM ---------------------------------------------------------------
REM [12] Apps and Features registration
REM ---------------------------------------------------------------
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL" >nul 2>&1
if errorlevel 1 (
    call :warn "[12] Apps and Features registration: missing (tool will not show in Settings; tray Uninstall still works)"
) else (
    call :ok "[12] Apps and Features registration: present (visible in Settings, Apps, Installed apps)"
)

call :log ""
call :log "--- Summary: !OK_COUNT! OK, !WARN_COUNT! WARN, !FAIL_COUNT! FAIL ---"
call :log ""
call :log "Share this doctor.log when reporting issues at:"
call :log "  https://github.com/asaf-aizone/Claude-for-word-RTL-fix/issues"
call :log ""
call :log "Full log saved to: %LOG%"

pause
endlocal
goto :eof

:log
echo %~1
echo %~1>> "%LOG%"
goto :eof

:ok
set /a OK_COUNT+=1
call :log "  [OK]   %~1"
goto :eof

:warn
set /a WARN_COUNT+=1
call :log "  [WARN] %~1"
goto :eof

:fail
set /a FAIL_COUNT+=1
call :log "  [FAIL] %~1"
goto :eof
