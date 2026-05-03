@echo off
REM Claude for Word RTL - installer.
REM 1. Verifies prerequisites (Node.js, WINWORD).
REM 2. Installs npm dependencies if missing.
REM 3. Creates a Startup-folder entry that launches the tray icon at login.
REM
REM Tray-only model:
REM   - No HKCU file-association overrides.
REM   - No Start Menu shortcut.
REM   - The tray icon is the single entry point: it launches at login from
REM     the user's Startup folder, and the user right-clicks it to Connect
REM     (which relaunches Word via the wrapper with the RTL fix enabled).
REM
REM Requires no admin rights (all changes are per-user files).

setlocal EnableDelayedExpansion

set "HERE=%~dp0"
if "%HERE:~-1%"=="\" set "HERE=%HERE:~0,-1%"
set "LOG=%HERE%\install.log"
echo Install started %DATE% %TIME% > "%LOG%"

call :log "================================================================"
call :log " Claude for Word RTL - Installer"
call :log "================================================================"
call :log ""

set "WRAPPER=%HERE%\word-wrapper.bat"
set "TRAYVBS=%HERE%\scripts\start-tray.vbs"
if not exist "%WRAPPER%" (
    call :log "[ERROR] word-wrapper.bat not found next to this installer."
    pause
    exit /b 1
)
if not exist "%TRAYVBS%" (
    call :log "[ERROR] scripts\start-tray.vbs not found. Installation cannot proceed."
    pause
    exit /b 1
)

REM ---------------------------------------------------------------
call :log "[1/4] Checking prerequisites..."
REM ---------------------------------------------------------------
where node >nul 2>&1
if errorlevel 1 (
    call :log ""
    call :log "  ================================================================"
    call :log "  [ERROR] Node.js is NOT installed."
    call :log "  ================================================================"
    call :log ""
    call :log "  This tool needs Node.js 16 or newer to run. Install it, then try"
    call :log "  install.bat again."
    call :log ""
    call :log "  How to install Node.js:"
    call :log "    1. Go to  https://nodejs.org/"
    call :log "    2. Download the LTS installer for Windows."
    call :log "    3. Run it - the defaults (Next-Next-Next) are fine. No admin"
    call :log "       rights required if you use the official installer."
    call :log "    4. Close this window, open a NEW Command Prompt, and verify:"
    call :log "         node --version"
    call :log "       You should see v16 or higher."
    call :log "    5. Run install.bat again."
    call :log ""
    pause
    exit /b 1
)
REM Node is on PATH. Also warn if it is below v16 (still runs but unsupported).
REM Delayed expansion (!var!) is required here because we assign and read in
REM the same parens block. Pre-clearing NODE_MAJOR protects against the
REM for-loop never firing (e.g. node exits non-zero before printing).
set "NODE_MAJOR="
for /f "tokens=1 delims=." %%V in ('node --version 2^>nul') do set "NODE_MAJOR=%%V"
if defined NODE_MAJOR (
    set "NODE_MAJOR=!NODE_MAJOR:v=!"
    if !NODE_MAJOR! LSS 16 (
        call :log "  [WARN] Node.js !NODE_MAJOR!.x detected. Version 16+ is recommended."
        call :log "         Proceeding, but please upgrade if anything misbehaves."
    ) else (
        call :log "  [OK] Node.js v!NODE_MAJOR!.x found."
    )
) else (
    call :log "  [OK] Node.js found (version check unavailable)."
)

REM --- Check that Word is not running ---
:check_word
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 goto word_closed
call :log "  [WAIT] Microsoft Word is running. Please close all Word windows, then press any key to continue."
pause >nul
goto check_word
:word_closed
call :log "  [OK] Word is not running."

set "WINWORD=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" set "WINWORD=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" (
    call :log "  [ERROR] WINWORD.EXE not found in default Office 16 locations."
    pause
    exit /b 1
)
call :log "  [OK] Word found."
call :log ""

REM ---------------------------------------------------------------
call :log "[2/4] Installing npm dependencies..."
REM ---------------------------------------------------------------
if not exist "%HERE%\scripts\node_modules\chrome-remote-interface" (
    call :log "  Installing dependencies, ~15 seconds..."
    pushd "%HERE%\scripts"
    call npm install --silent >> "%LOG%" 2>&1
    if errorlevel 1 (
        call :log "  [ERROR] npm install failed. See install.log for details."
        popd
        pause
        exit /b 1
    )
    popd
    call :log "  [OK] Dependencies installed."
) else (
    call :log "  [OK] Dependencies already present."
)
call :log ""

REM ---------------------------------------------------------------
call :log "[3/4] Creating Startup-folder entry for the tray icon..."
REM ---------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%\scripts\create-shortcut.ps1" -TrayLauncher "%TRAYVBS%" -WorkingDir "%HERE%" -IconPath "%WINWORD%" >> "%LOG%" 2>&1
if errorlevel 1 (
    call :log "  [WARN] Could not create Startup entry. Tray will not auto-launch at login."
    call :log "         You can still launch the tray manually by double-clicking scripts\start-tray.vbs."
) else (
    call :log "  [OK] Startup entry placed. Tray will launch automatically at next login."
)
call :log ""

REM ---------------------------------------------------------------
call :log "[4/4] Registering in Windows 'Apps and Features'..."
REM ---------------------------------------------------------------
REM Per-user HKCU Uninstall key so the tool appears in Settings > Apps >
REM Installed apps. No admin rights required. UninstallString points at
REM uninstall.bat, which is what both this key and the tray's Uninstall
REM menu item invoke.
set "REGKEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL"
reg add "%REGKEY%" /v DisplayName     /t REG_SZ /d "Claude for Word RTL Fix" /f >nul 2>&1
reg add "%REGKEY%" /v DisplayVersion  /t REG_SZ /d "0.2.1" /f >nul 2>&1
reg add "%REGKEY%" /v Publisher       /t REG_SZ /d "Asaf Abramzon" /f >nul 2>&1
reg add "%REGKEY%" /v InstallLocation /t REG_SZ /d "%HERE%" /f >nul 2>&1
reg add "%REGKEY%" /v UninstallString /t REG_SZ /d "\"%HERE%\uninstall.bat\"" /f >nul 2>&1
reg add "%REGKEY%" /v DisplayIcon     /t REG_SZ /d "%WINWORD%,0" /f >nul 2>&1
reg add "%REGKEY%" /v URLInfoAbout    /t REG_SZ /d "https://github.com/asaf-aizone/Claude-for-Office-RTL-fix" /f >nul 2>&1
reg add "%REGKEY%" /v NoModify        /t REG_DWORD /d 1 /f >nul 2>&1
reg add "%REGKEY%" /v NoRepair        /t REG_DWORD /d 1 /f >nul 2>&1
if errorlevel 1 (
    call :log "  [WARN] Could not write Apps and Features registration. Uninstall via tray still works."
) else (
    call :log "  [OK] Registered. You can now uninstall via the Windows Settings Apps page, or via the tray."
)
call :log ""

REM ---------------------------------------------------------------
REM Cleanup: legacy Auto-Enable env var from v0.1.x installs.
REM
REM v0.1.x persisted WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222
REM as a USER env var so Word would expose CDP on every launch without
REM the user having to click Connect. That value is also read by every
REM other WebView2 host on the account (Teams, Outlook, Edge WebView,
REM OneDrive UI). Enterprise EDR products (Microsoft Defender for
REM Endpoint, CrowdStrike, SentinelOne) treat unexpected modifications
REM of WebView2 browser arguments as a token-theft signal and may trigger
REM host isolation on managed devices.
REM
REM v0.2.0+ no longer persists this env var. The wrappers (word-wrapper.bat
REM and equivalents for Excel/PowerPoint) set the flag in their own
REM process scope only, which Word inherits when launched through them
REM but Teams/Outlook/Edge do not.
REM
REM On install we silently remove the legacy value if (and ONLY if) it
REM matches one of our known strings. A user-modified value is preserved.
REM Goto/labels on purpose: log strings may include "(legacy ...)"  with
REM a closing paren that would terminate a parens-block early.
REM ---------------------------------------------------------------
set "AE_LEGACY_9222=--remote-debugging-port=9222"
set "AE_LEGACY_0=--remote-debugging-port=0"
set "AE_CHECK=%TEMP%\cwr_ae_check.txt"
reg query "HKCU\Environment" /v "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS" > "%AE_CHECK%" 2>nul
if errorlevel 1 goto :ae_cleanup_done
findstr /C:"%AE_LEGACY_9222%" "%AE_CHECK%" >nul 2>&1
if not errorlevel 1 goto :ae_match_legacy
findstr /C:"%AE_LEGACY_0%" "%AE_CHECK%" >nul 2>&1
if not errorlevel 1 goto :ae_match_legacy
goto :ae_cleanup_done
:ae_match_legacy
REM Defensive: only auto-clean when the value is EXACTLY one of our
REM legacy strings, no trailing additions the user may have appended.
findstr /C:"%AE_LEGACY_9222% " "%AE_CHECK%" >nul 2>&1
if not errorlevel 1 goto :ae_cleanup_done
findstr /C:"%AE_LEGACY_0% " "%AE_CHECK%" >nul 2>&1
if not errorlevel 1 goto :ae_cleanup_done
reg delete "HKCU\Environment" /v "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS" /f >nul 2>&1
if errorlevel 1 goto :ae_cleanup_done
call :log "  [INFO] Removed legacy Auto-Enable env var from older install."
call :log "         RTL is now activated per-process via the tray Connect menu only."
:ae_cleanup_done
if exist "%AE_CHECK%" del /q "%AE_CHECK%" >nul 2>&1
call :log ""

REM ---------------------------------------------------------------
REM Start the tray now so the user can use the tool immediately
REM without having to log out and back in.
REM
REM Upgrade path: if a previous tray is already running (reinstall over
REM the top of an older version), it holds the singleton mutex and the
REM new tray would exit silently, leaving the OLD ps1 code loaded. Kill
REM the old tray via its PID file first so the fresh tray-icon.ps1 gets
REM loaded. Same for the injector - if an older version's injector is
REM still alive, the new tray's auto-launch path will see it as alive
REM and never restart it with whatever changes the new inject.js has.
REM ---------------------------------------------------------------
set "TRAYPIDFILE=%TEMP%\claude-word-rtl.tray.pid"
set "INJPIDFILE=%TEMP%\claude-word-rtl.pid"
if exist "%TRAYPIDFILE%" (
    for /f "usebackq delims=" %%P in ("%TRAYPIDFILE%") do taskkill /F /PID %%P >nul 2>&1
    del /q "%TRAYPIDFILE%" >nul 2>&1
    call :log "  Stopped previous tray so the new code is loaded."
)
if exist "%INJPIDFILE%" (
    for /f "usebackq delims=" %%P in ("%INJPIDFILE%") do taskkill /F /PID %%P >nul 2>&1
    del /q "%INJPIDFILE%" >nul 2>&1
    call :log "  Stopped previous injector - the new tray will relaunch it."
)
if exist "%TEMP%\claude-word-rtl.lock"   del /q "%TEMP%\claude-word-rtl.lock"   >nul 2>&1
if exist "%TEMP%\claude-word-rtl.status" del /q "%TEMP%\claude-word-rtl.status" >nul 2>&1

call :log "  Starting tray icon..."
start "" wscript.exe "%TRAYVBS%"
call :log "  [OK] Tray icon started. Right-click it near the clock to Connect."
call :log ""

call :log "================================================================"
call :log " Installation complete."
call :log "================================================================"
call :log ""
call :log "How to use:"
call :log "  1. Open Word normally, or keep it open."
call :log "  2. Right-click the tray icon near the clock and pick 'Connect'."
call :log "     The tray will relaunch Word via the wrapper with RTL enabled,"
call :log "     reopening the documents you had open."
call :log ""
call :log "  Connect must be used once per Word session. The RTL flag is set"
call :log "  per-process by the wrapper, not as a persistent user setting,"
call :log "  so other apps on your account are never affected."
call :log ""
call :log "Security notice:"
call :log "  While Word is running via Connect, the WebView2 debug interface"
call :log "  is exposed to local processes on this machine (localhost only,"
call :log "  not the network). Only Word is affected, not Teams/Outlook/Edge."
call :log "  Close Word when you are not using the add-in."
call :log ""
call :log "Full log saved to: %LOG%"
call :log "To remove: run uninstall.bat"
call :log ""
pause
endlocal
goto :eof

:log
REM Empty-arg case: echo a real blank line rather than letting cmd print
REM "ECHO is off." when %~1 is an empty quoted string.
REM
REM Goto-based (not if-else-parens) on purpose: when the logged string
REM contains a closing paren ")" - e.g. "Outlook panes, Edge WebView
REM hosts)." - cmd inside a parens block treats that ")" as the end of
REM the block and anything after as a new statement, which then fails
REM with "xxx was unexpected at this time.". Keeping :log's body at
REM top-level statement scope avoids the trap entirely.
if not "%~1"=="" goto :log_msg
echo.
echo.>> "%LOG%"
goto :eof
:log_msg
echo %~1
echo %~1>> "%LOG%"
goto :eof
