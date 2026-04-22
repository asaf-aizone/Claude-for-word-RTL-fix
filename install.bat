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
    call :log "  [ERROR] Node.js is not installed. Install from https://nodejs.org first."
    pause
    exit /b 1
)
call :log "  [OK] Node.js found."

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
reg add "%REGKEY%" /v DisplayVersion  /t REG_SZ /d "0.1.0" /f >nul 2>&1
reg add "%REGKEY%" /v Publisher       /t REG_SZ /d "Asaf Abramzon" /f >nul 2>&1
reg add "%REGKEY%" /v InstallLocation /t REG_SZ /d "%HERE%" /f >nul 2>&1
reg add "%REGKEY%" /v UninstallString /t REG_SZ /d "\"%HERE%\uninstall.bat\"" /f >nul 2>&1
reg add "%REGKEY%" /v DisplayIcon     /t REG_SZ /d "%WINWORD%,0" /f >nul 2>&1
reg add "%REGKEY%" /v URLInfoAbout    /t REG_SZ /d "https://github.com/asaf-aizone/Claude-for-word-RTL-fix" /f >nul 2>&1
reg add "%REGKEY%" /v NoModify        /t REG_DWORD /d 1 /f >nul 2>&1
reg add "%REGKEY%" /v NoRepair        /t REG_DWORD /d 1 /f >nul 2>&1
if errorlevel 1 (
    call :log "  [WARN] Could not write Apps and Features registration. Uninstall via tray still works."
) else (
    call :log "  [OK] Registered. You can now uninstall via the Windows Settings Apps page, or via the tray."
)
call :log ""

REM ---------------------------------------------------------------
REM Start the tray now so the user can use the tool immediately
REM without having to log out and back in.
REM ---------------------------------------------------------------
call :log "  Starting tray icon..."
start "" wscript.exe "%TRAYVBS%"
call :log "  [OK] Tray icon started. Right-click it near the clock to Connect."
call :log ""

call :log "================================================================"
call :log " Installation complete."
call :log "================================================================"
call :log ""
call :log "How to use:"
call :log "  1. Open Word normally (or keep it open)."
call :log "  2. Right-click the tray icon near the clock and pick 'Connect'."
call :log "     The tray will relaunch Word via the wrapper with RTL enabled,"
call :log "     reopening the documents you had open."
call :log ""
call :log "Security notice:"
call :log "  While Word is running via this tool, WebView2 debug port 9222"
call :log "  is exposed to local processes on this machine (localhost only,"
call :log "  not the network). Close Word when you are not using the add-in."
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
if "%~1"=="" (
    echo.
    echo.>> "%LOG%"
) else (
    echo %~1
    echo %~1>> "%LOG%"
)
goto :eof
