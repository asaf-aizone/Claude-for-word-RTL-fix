@echo off
REM Claude for Office RTL - diagnostic script.
REM Runs a series of environment checks and writes a report to doctor.log.
REM Share doctor.log when reporting issues.
REM
REM v0.2.0 multi-app architecture: Word + Excel + PowerPoint, dynamic CDP
REM ports per Office WebView2 host (--remote-debugging-port=0). Port
REM discovery is performed at runtime by scripts/port-discovery.js, which
REM walks tasklist for msedgewebview2.exe + netstat for LISTENING sockets.
REM
REM Style note: this script avoids if/else paren blocks for any branch that
REM logs a dynamic string. cmd's parser treats a literal ')' inside a paren
REM block as the block terminator, even when quoted, which corrupts logged
REM messages that contain Office paths like 'Program Files (x86)' or list
REM markers like '(none)'. Branches are flattened with goto/labels instead -
REM same pattern v0.1.x used for the port-row helpers.

setlocal EnableDelayedExpansion

set "HERE=%~dp0"
if "%HERE:~-1%"=="\" set "HERE=%HERE:~0,-1%"
set "LOG=%HERE%\doctor.log"
echo Doctor started %DATE% %TIME% > "%LOG%"

set /a OK_COUNT=0
set /a INFO_COUNT=0
set /a WARN_COUNT=0
set /a FAIL_COUNT=0

call :log "================================================================"
call :log " Claude for Office RTL - Doctor"
call :log "================================================================"
call :log ""

REM ---------------------------------------------------------------
REM [1/15] Node.js
REM ---------------------------------------------------------------
where node >nul 2>&1
if errorlevel 1 goto :node_missing
for /f "delims=" %%v in ('node --version 2^>^&1') do set "NODE_VER=%%v"
call :ok "[1/15] Node.js: !NODE_VER!"
goto :step_2
:node_missing
call :fail "[1/15] Node.js: not found in PATH"
:step_2

REM ---------------------------------------------------------------
REM [2/15] npm
REM ---------------------------------------------------------------
where npm >nul 2>&1
if errorlevel 1 goto :npm_missing
for /f "delims=" %%v in ('npm --version 2^>^&1') do set "NPM_VER=%%v"
call :ok "[2/15] npm: !NPM_VER!"
goto :step_3
:npm_missing
call :fail "[2/15] npm: not found in PATH"
:step_3

REM ---------------------------------------------------------------
REM [3/15] chrome-remote-interface
REM ---------------------------------------------------------------
set "CRI_INSTALLED="
if not exist "%HERE%\scripts\node_modules\chrome-remote-interface" goto :cri_missing
set "CRI_INSTALLED=1"
call :ok "[3/15] chrome-remote-interface: installed"
goto :step_4
:cri_missing
call :fail "[3/15] chrome-remote-interface: missing - run install.bat or 'npm install' in scripts\"
:step_4

REM ---------------------------------------------------------------
REM [4/15] Office apps installed (Word, Excel, PowerPoint)
REM ---------------------------------------------------------------
REM Click-to-Run paths first (modern installs), then MSI fallback (no 'root\').
REM Word is the headline app: missing -> FAIL. Excel/PowerPoint optional -> INFO.
call :find_office_app WINWORD.EXE
if not defined OFFICE_FOUND goto :word_missing
call :ok "[4/15] WINWORD.EXE: !OFFICE_FOUND!"
goto :step_4_excel
:word_missing
call :fail "[4/15] WINWORD.EXE: not found in default Office 16 locations"
:step_4_excel
call :find_office_app EXCEL.EXE
if not defined OFFICE_FOUND goto :excel_missing
call :ok "[4/15] EXCEL.EXE: !OFFICE_FOUND!"
goto :step_4_pptx
:excel_missing
call :info "[4/15] EXCEL.EXE: not found - Excel support optional, OK to skip if you do not use Excel"
:step_4_pptx
call :find_office_app POWERPNT.EXE
if not defined OFFICE_FOUND goto :pptx_missing
call :ok "[4/15] POWERPNT.EXE: !OFFICE_FOUND!"
goto :step_5
:pptx_missing
call :info "[4/15] POWERPNT.EXE: not found - PowerPoint support optional, OK to skip if you do not use PowerPoint"
:step_5

REM ---------------------------------------------------------------
REM [5/15] Office apps currently running
REM ---------------------------------------------------------------
call :running_check WINWORD.EXE Word
call :running_check EXCEL.EXE Excel
call :running_check POWERPNT.EXE PowerPoint

REM ---------------------------------------------------------------
REM [6/15] Dynamic CDP ports discovered
REM ---------------------------------------------------------------
if not defined CRI_INSTALLED goto :ports_no_deps
set "PORTS_OUT=%TEMP%\claude-office-rtl.doctor.ports.txt"
if exist "%PORTS_OUT%" del /q "%PORTS_OUT%" >nul 2>&1
pushd "%HERE%"
node -e "const p=require('./scripts/port-discovery'); p.discoverPorts().then(s=>console.log([...s].sort().join(',')||'(none)')).catch(e=>console.log('ERR: '+e.message))" > "%PORTS_OUT%" 2>&1
popd
set "PORTS_LINE="
for /f "usebackq delims=" %%L in ("%PORTS_OUT%") do set "PORTS_LINE=%%L"
if not defined PORTS_LINE set "PORTS_LINE=(no output)"
call :info "[6/15] Dynamic CDP ports discovered: !PORTS_LINE!"
if exist "%PORTS_OUT%" del /q "%PORTS_OUT%" >nul 2>&1
goto :step_7
:ports_no_deps
call :info "[6/15] Dynamic CDP ports: deps not installed - run install.bat to enable port discovery"
:step_7

REM ---------------------------------------------------------------
REM [7/15] Active Claude targets (per Office app)
REM ---------------------------------------------------------------
if not defined CRI_INSTALLED goto :targets_no_deps
set "TARGETS_OUT=%TEMP%\claude-office-rtl.doctor.targets.txt"
if exist "%TARGETS_OUT%" del /q "%TARGETS_OUT%" >nul 2>&1
pushd "%HERE%"
node -e "const p=require('./scripts/port-discovery'); p.discoverActiveTargets().then(arr=>console.log(arr.length===0?'(none)':arr.map(r=>(r.app?r.app.name:'unknown')+'@'+r.port).join(', '))).catch(e=>console.log('ERR: '+e.message))" > "%TARGETS_OUT%" 2>&1
popd
set "TARGETS_LINE="
for /f "usebackq delims=" %%L in ("%TARGETS_OUT%") do set "TARGETS_LINE=%%L"
if not defined TARGETS_LINE set "TARGETS_LINE=(no output)"
call :info "[7/15] Active Claude targets: !TARGETS_LINE!"
if exist "%TARGETS_OUT%" del /q "%TARGETS_OUT%" >nul 2>&1
goto :step_8
:targets_no_deps
call :info "[7/15] Active Claude targets: deps not installed - run install.bat to enable target discovery"
:step_8

REM ---------------------------------------------------------------
REM [8/15] Injector PID file + alive
REM ---------------------------------------------------------------
set "INJ_PIDFILE=%TEMP%\claude-word-rtl.pid"
set "INJ_PID="
if exist "%INJ_PIDFILE%" set /p INJ_PID=<"%INJ_PIDFILE%"
if not defined INJ_PID goto :inj_no_pidfile
tasklist /FI "PID eq %INJ_PID%" 2>nul | "%SystemRoot%\System32\find.exe" "%INJ_PID%" >nul
if errorlevel 1 goto :inj_stale
call :ok "[8/15] Injector: running, PID %INJ_PID%"
goto :step_9
:inj_stale
call :warn "[8/15] Injector: stale PID file at %INJ_PIDFILE% - node.exe with that PID is not running"
goto :step_9
:inj_no_pidfile
call :warn "[8/15] Injector: not running - Connect via tray, or launch a wrapper"
:step_9

REM ---------------------------------------------------------------
REM [9/15] Injector aggregate status file
REM ---------------------------------------------------------------
set "STATUS_FILE=%TEMP%\claude-word-rtl.status"
if not exist "%STATUS_FILE%" goto :status_missing
set "STATUS_LINE="
for /f "usebackq delims=" %%L in ("%STATUS_FILE%") do if not defined STATUS_LINE set "STATUS_LINE=%%L"
if not defined STATUS_LINE set "STATUS_LINE=(empty)"
call :report_status "!STATUS_LINE!"
goto :step_10
:status_missing
call :info "[9/15] Aggregate status file: missing - injector has not written %STATUS_FILE% yet"
:step_10

REM ---------------------------------------------------------------
REM [10/15] Per-app injector status (v0.2.0+)
REM ---------------------------------------------------------------
set "APPS_STATUS=%TEMP%\claude-office-rtl.apps.json"
if not exist "%APPS_STATUS%" goto :apps_status_missing
set "APPS_OUT=%TEMP%\claude-office-rtl.doctor.apps.txt"
if exist "%APPS_OUT%" del /q "%APPS_OUT%" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $j = Get-Content -Raw -LiteralPath $env:APPS_STATUS | ConvertFrom-Json; $parts = @(); foreach ($n in 'Word','Excel','PowerPoint') { $v = $j.$n; if (-not $v) { $v = '(missing)' }; $parts += ('{0}: {1}' -f $n, $v) }; ($parts -join ', ') } catch { 'ERR: ' + $_.Exception.Message }" > "%APPS_OUT%" 2>&1
set "APPS_LINE="
for /f "usebackq delims=" %%L in ("%APPS_OUT%") do set "APPS_LINE=%%L"
if not defined APPS_LINE set "APPS_LINE=(no output)"
call :info "[10/15] Per-app status: !APPS_LINE!"
if exist "%APPS_OUT%" del /q "%APPS_OUT%" >nul 2>&1
goto :step_11
:apps_status_missing
call :info "[10/15] Per-app status: file missing at %APPS_STATUS% - injector may be from older version, or not running"
:step_11

REM ---------------------------------------------------------------
REM [11/15] Tray PID file + alive
REM ---------------------------------------------------------------
set "TRAY_PIDFILE=%TEMP%\claude-word-rtl.tray.pid"
set "TRAY_PID="
if exist "%TRAY_PIDFILE%" set /p TRAY_PID=<"%TRAY_PIDFILE%"
if not defined TRAY_PID goto :tray_no_pidfile
tasklist /FI "PID eq %TRAY_PID%" 2>nul | "%SystemRoot%\System32\find.exe" "%TRAY_PID%" >nul
if errorlevel 1 goto :tray_stale
call :ok "[11/15] Tray process: running, PID %TRAY_PID%"
goto :step_12
:tray_stale
call :warn "[11/15] Tray process: stale PID file at %TRAY_PIDFILE% - tray not actually running"
goto :step_12
:tray_no_pidfile
call :warn "[11/15] Tray process: not running - double-click scripts\start-tray.vbs or log out and back in"
:step_12

REM ---------------------------------------------------------------
REM [12/15] Startup folder shortcut
REM ---------------------------------------------------------------
set "STARTUP_LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk"
if not exist "%STARTUP_LNK%" goto :startup_missing
call :ok "[12/15] Startup entry 'Claude for Word RTL Tray': present, tray auto-launches at login"
goto :step_13
:startup_missing
call :warn "[12/15] Startup entry 'Claude for Word RTL Tray': missing - run install.bat to create it"
:step_13

REM ---------------------------------------------------------------
REM [13/15] Apps and Features registration
REM ---------------------------------------------------------------
set "REGKEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL"
reg query "%REGKEY%" >nul 2>&1
if errorlevel 1 goto :appsfeat_missing
set "DISPLAY_VER="
for /f "tokens=2,*" %%a in ('reg query "%REGKEY%" /v DisplayVersion 2^>nul ^| "%SystemRoot%\System32\find.exe" /I "DisplayVersion"') do set "DISPLAY_VER=%%b"
if not defined DISPLAY_VER goto :appsfeat_no_ver
call :ok "[13/15] Apps and Features registration: present, DisplayVersion=!DISPLAY_VER!"
goto :step_14
:appsfeat_no_ver
call :ok "[13/15] Apps and Features registration: present, DisplayVersion not readable"
goto :step_14
:appsfeat_missing
call :warn "[13/15] Apps and Features registration: missing - tool will not show in Settings, tray Uninstall still works"
:step_14

REM ---------------------------------------------------------------
REM [14/15] Legacy env var must NOT be persisted (CRITICAL)
REM ---------------------------------------------------------------
REM v0.2.0+ NEVER writes HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS.
REM v0.1.0 - v0.1.3 wrote it as the Auto-enable mechanism. EDR products treat
REM modifications of WebView2 browser arguments as token-theft signals and have
REM triggered host isolation in the field. The wrappers now set the flag in
REM their own process scope only.
set "LEGACY_VAR_NAME=WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"
set "LEGACY_VAL="
for /f "tokens=2,*" %%a in ('reg query "HKCU\Environment" /v %LEGACY_VAR_NAME% 2^>nul ^| "%SystemRoot%\System32\find.exe" /I "%LEGACY_VAR_NAME%"') do set "LEGACY_VAL=%%b"
if not defined LEGACY_VAL goto :legacy_ok
set "LEGACY_HIT="
if /I "%LEGACY_VAL%"=="--remote-debugging-port=9222" set "LEGACY_HIT=1"
if /I "%LEGACY_VAL%"=="--remote-debugging-port=0" set "LEGACY_HIT=1"
if defined LEGACY_HIT goto :legacy_fail
call :info "[14/15] Legacy env var %LEGACY_VAR_NAME% is set to a non-default value '%LEGACY_VAL%' - we did not write this and we will not touch it"
goto :step_15
:legacy_fail
call :fail "[14/15] Legacy env var %LEGACY_VAR_NAME% IS SET to '%LEGACY_VAL%' - this is an EDR trigger, v0.2.0+ never writes it"
call :log "         Fix: run uninstall.bat - clears it automatically. Or manually:"
call :log "           reg delete HKCU\Environment /v %LEGACY_VAR_NAME% /f"
goto :step_15
:legacy_ok
call :ok "[14/15] Legacy env var %LEGACY_VAR_NAME%: not set, correct for v0.2.0+"
:step_15

REM ---------------------------------------------------------------
REM [15/15] WebView2 runtime
REM ---------------------------------------------------------------
set "WV2_FOUND="
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" >nul 2>&1
if not errorlevel 1 set "WV2_FOUND=WOW6432Node"
if defined WV2_FOUND goto :wv2_ok
reg query "HKLM\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" >nul 2>&1
if not errorlevel 1 set "WV2_FOUND=HKLM\SOFTWARE"
if not defined WV2_FOUND goto :wv2_missing
:wv2_ok
call :ok "[15/15] WebView2 runtime: installed, key %WV2_FOUND%"
goto :summary
:wv2_missing
call :fail "[15/15] WebView2 runtime: not detected - Claude add-in will not render"
:summary

call :log ""
call :log "--- Summary: !OK_COUNT! OK, !INFO_COUNT! INFO, !WARN_COUNT! WARN, !FAIL_COUNT! FAIL ---"
call :log ""
call :log "Attach doctor.log when reporting issues at:"
call :log "  https://github.com/asaf-aizone/Claude-for-word-RTL-fix/issues"
call :log ""
call :log "Full log saved to: %LOG%"

pause
endlocal
goto :eof

REM ===============================================================
REM Helpers
REM ===============================================================

:log
REM Print to stdout and append to %LOG%. Empty string prints a blank line.
REM "echo." with the trailing dot is the documented cmd idiom for a blank
REM line that does not print "ECHO is on/off." status.
if "%~1"=="" goto :log_blank
echo %~1
echo %~1>> "%LOG%"
goto :eof
:log_blank
echo.
echo.>> "%LOG%"
goto :eof

:ok
set /a OK_COUNT+=1
call :log "  [OK]   %~1"
goto :eof

:info
set /a INFO_COUNT+=1
call :log "  [INFO] %~1"
goto :eof

:warn
set /a WARN_COUNT+=1
call :log "  [WARN] %~1"
goto :eof

:fail
set /a FAIL_COUNT+=1
call :log "  [FAIL] %~1"
goto :eof

:find_office_app
REM In: %1 = exe name (WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE)
REM Out: OFFICE_FOUND = full path or empty
set "OFFICE_FOUND="
set "OFFICE_EXE=%~1"
set "P1=C:\Program Files\Microsoft Office\root\Office16\%OFFICE_EXE%"
set "P2=C:\Program Files (x86)\Microsoft Office\root\Office16\%OFFICE_EXE%"
set "P3=C:\Program Files\Microsoft Office\Office16\%OFFICE_EXE%"
set "P4=C:\Program Files (x86)\Microsoft Office\Office16\%OFFICE_EXE%"
if exist "%P1%" set "OFFICE_FOUND=%P1%"
if not defined OFFICE_FOUND if exist "%P2%" set "OFFICE_FOUND=%P2%"
if not defined OFFICE_FOUND if exist "%P3%" set "OFFICE_FOUND=%P3%"
if not defined OFFICE_FOUND if exist "%P4%" set "OFFICE_FOUND=%P4%"
goto :eof

:running_check
REM In: %1 = exe name, %2 = friendly name
tasklist /FI "IMAGENAME eq %~1" 2>nul | "%SystemRoot%\System32\find.exe" /I "%~1" >nul
if errorlevel 1 goto :running_no
call :info "[5/15] %~2 running: yes"
goto :eof
:running_no
call :info "[5/15] %~2 running: no"
goto :eof

:report_status
REM In: %1 = status string from %TEMP%\claude-word-rtl.status. May contain
REM ')' (e.g. 'ERROR:foo (bar)'), so we never enter a paren block here.
set "STATUS_VAL=%~1"
if /I "%STATUS_VAL%"=="CONNECTED" call :ok "[9/15] Aggregate status: CONNECTED"
if /I "%STATUS_VAL%"=="CONNECTED" goto :eof
if /I "%STATUS_VAL%"=="DISCONNECTED" call :info "[9/15] Aggregate status: DISCONNECTED"
if /I "%STATUS_VAL%"=="DISCONNECTED" goto :eof
echo %STATUS_VAL% | "%SystemRoot%\System32\findstr.exe" /I /B "ERROR:" >nul 2>&1
if errorlevel 1 goto :report_status_other
call :warn "[9/15] Aggregate status: %STATUS_VAL%"
goto :eof
:report_status_other
call :info "[9/15] Aggregate status: %STATUS_VAL%"
goto :eof
