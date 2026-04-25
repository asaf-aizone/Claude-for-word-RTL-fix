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
call :log "[1/5] Checking prerequisites..."
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
call :log "[2/5] Installing npm dependencies..."
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
call :log "[3/5] Creating Startup-folder entry for the tray icon..."
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
call :log "[4/5] Registering in Windows 'Apps and Features'..."
REM ---------------------------------------------------------------
REM Per-user HKCU Uninstall key so the tool appears in Settings > Apps >
REM Installed apps. No admin rights required. UninstallString points at
REM uninstall.bat, which is what both this key and the tray's Uninstall
REM menu item invoke.
set "REGKEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL"
reg add "%REGKEY%" /v DisplayName     /t REG_SZ /d "Claude for Word RTL Fix" /f >nul 2>&1
reg add "%REGKEY%" /v DisplayVersion  /t REG_SZ /d "0.1.3" /f >nul 2>&1
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
REM Migration from v0.1.x: legacy Auto-Enable used a fixed port 9222,
REM which fails silently when more than one Office app is open at once
REM (only the first to start grabs 9222; Excel/PowerPoint started after
REM Word get no debug surface). v0.2.0 switches to dynamic ports via
REM --remote-debugging-port=0 so each Office process picks its own free
REM port. If the existing value is exactly our old string, update it
REM silently - it is our value, we own the migration. Anything else
REM (already =0, or user-customized) is left untouched.
REM
REM Goto/labels on purpose: log strings include "(was fixed port 9222)"
REM with a closing paren, which would terminate a parens-block early.
REM ---------------------------------------------------------------
set "MIG_LEGACY=--remote-debugging-port=9222"
set "MIG_NEW=--remote-debugging-port=0"
set "MIG_CHECK=%TEMP%\cwr_mig_check.txt"
reg query "HKCU\Environment" /v "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS" > "%MIG_CHECK%" 2>nul
if errorlevel 1 goto :migration_done
findstr /C:"%MIG_LEGACY%" "%MIG_CHECK%" >nul 2>&1
if errorlevel 1 goto :migration_done
REM Defensive: legacy string matched, but make sure the user has not
REM tacked on extra args after it (e.g. "--remote-debugging-port=9222
REM --some-other-flag"). Only auto-migrate when the value is EXACTLY
REM our legacy string with no trailing additions.
findstr /C:"%MIG_LEGACY% " "%MIG_CHECK%" >nul 2>&1
if not errorlevel 1 goto :migration_done
reg add "HKCU\Environment" /v "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS" /t REG_SZ /d "%MIG_NEW%" /f >nul 2>&1
if errorlevel 1 goto :migration_done
call :log "  [INFO] Auto-Enable updated to dynamic ports (was fixed port 9222) for Excel + PowerPoint support."
REM Broadcast WM_SETTINGCHANGE so running processes pick up the new
REM value without requiring a logout. Same approach as the Auto-Enable
REM ON path below; temp .ps1 file avoids cmd/PS quoting traps.
set "MIG_BCAST_PS=%TEMP%\cwr_mig_bcast.ps1"
>  "%MIG_BCAST_PS%" echo Add-Type -Namespace W -Name E -MemberDefinition '[System.Runtime.InteropServices.DllImport(^"user32.dll^")] public static extern System.IntPtr SendMessageTimeout(System.IntPtr h, uint m, System.UIntPtr w, string l, uint f, uint t, out System.UIntPtr r);'
>> "%MIG_BCAST_PS%" echo $r=[System.UIntPtr]::Zero
>> "%MIG_BCAST_PS%" echo [W.E]::SendMessageTimeout([IntPtr]0xFFFF,0x1A,[System.UIntPtr]::Zero,'Environment',2,5000,[ref]$r) ^| Out-Null
powershell -NoProfile -ExecutionPolicy Bypass -File "%MIG_BCAST_PS%" >nul 2>&1
del /q "%MIG_BCAST_PS%" >nul 2>&1
:migration_done
if exist "%MIG_CHECK%" del /q "%MIG_CHECK%" >nul 2>&1
call :log ""

REM ---------------------------------------------------------------
call :log "[5/5] Auto-enable RTL on every Word launch?"
REM ---------------------------------------------------------------
REM Sets HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS so every
REM Word process this user launches (taskbar, Recent, docx double-click,
REM email attachment) picks up --remote-debugging-port=9222 at startup.
REM Without this, the user has to right-click the tray and pick Connect
REM every session to relaunch Word with the debug port enabled, which is
REM clunky when Word is already open with documents.
REM
REM Prompted (not defaulted) because the variable is user-scoped and
REM also read by other WebView2 apps on the account (Teams, Outlook
REM panes, Edge WebView hosts). In practice they don't bind debug port
REM 9222, but the user deserves a heads-up before we persist user-level
REM environment state.
REM
REM Skipped silently on reinstall when the variable is already set to
REM our exact value - no reason to re-prompt.

set "AE_NAME=WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"
set "AE_VALUE=--remote-debugging-port=0"
set "AE_CHECK=%TEMP%\cwr_ae_check.txt"
set "AE_STATE=unset"
reg query "HKCU\Environment" /v "%AE_NAME%" > "%AE_CHECK%" 2>nul
if exist "%AE_CHECK%" (
    findstr /I /C:"%AE_NAME%" "%AE_CHECK%" >nul && set "AE_STATE=conflict"
    findstr /I /C:"%AE_VALUE%" "%AE_CHECK%" >nul && set "AE_STATE=ours"
    del /q "%AE_CHECK%" >nul 2>&1
)

set "AUTO_ENABLE_ON="
if "%AE_STATE%"=="ours" (
    call :log "  [OK] Auto-enable is already on. Nothing to change."
    set "AUTO_ENABLE_ON=1"
    goto :auto_enable_done
)

call :log "  With Auto-enable ON, every new Word launch picks up the RTL debug"
call :log "  flag automatically. You will NOT need to right-click the tray and"
call :log "  pick Connect each time - Word just opens with RTL ready."
call :log ""
call :log "  Heads-up: this sets a USER environment variable that is also read"
call :log "  by other WebView2 apps on your account (Teams, Outlook panes,"
call :log "  Edge WebView hosts). In practice they do not bind debug port 9222,"
call :log "  but you should know this affects them too. Uninstall removes it,"
call :log "  and you can toggle it any time from the tray icon."
call :log ""
if "%AE_STATE%"=="conflict" (
    call :log "  WARNING: the variable is currently set to a different value. If"
    call :log "  you answer Y, that value will be replaced."
    call :log ""
)

choice /c yn /n /m "  Turn Auto-enable ON now? [Y/N] "
set "CHOICE_ERR=%errorlevel%"
if "%CHOICE_ERR%"=="2" (
    call :log "  [SKIP] Auto-enable left off. Use the tray 'Connect' option each time,"
    call :log "         or enable later from the tray menu."
    goto :auto_enable_done
)

reg add "HKCU\Environment" /v "%AE_NAME%" /t REG_SZ /d "%AE_VALUE%" /f >nul 2>&1
if errorlevel 1 (
    call :log "  [WARN] Could not write the environment variable. You can toggle"
    call :log "         Auto-enable later from the tray icon."
    goto :auto_enable_done
)

REM Broadcast WM_SETTINGCHANGE so newly-started processes (Explorer, the
REM tray we're about to launch, any Word launched next) see the new env
REM var without requiring a logout. Temp .ps1 file avoids cmd/PS quoting
REM traps when embedding DllImport signatures inline.
set "BCAST_PS=%TEMP%\cwr_bcast.ps1"
>  "%BCAST_PS%" echo Add-Type -Namespace W -Name E -MemberDefinition '[System.Runtime.InteropServices.DllImport(^"user32.dll^")] public static extern System.IntPtr SendMessageTimeout(System.IntPtr h, uint m, System.UIntPtr w, string l, uint f, uint t, out System.UIntPtr r);'
>> "%BCAST_PS%" echo $r=[System.UIntPtr]::Zero
>> "%BCAST_PS%" echo [W.E]::SendMessageTimeout([IntPtr]0xFFFF,0x1A,[System.UIntPtr]::Zero,'Environment',2,5000,[ref]$r) ^| Out-Null
powershell -NoProfile -ExecutionPolicy Bypass -File "%BCAST_PS%" >nul 2>&1
del /q "%BCAST_PS%" >nul 2>&1

call :log "  [OK] Auto-enable turned on. Future Word launches have RTL ready."
set "AUTO_ENABLE_ON=1"

:auto_enable_done
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
if defined AUTO_ENABLE_ON (
    call :log "How to use:"
    call :log "  Auto-enable is ON. Just open Word normally - the RTL fix is available"
    call :log "  as soon as you open the Claude panel. The tray icon turns green when"
    call :log "  the injector attaches."
    call :log ""
    call :log "  Word windows that were already running before this install still need"
    call :log "  one Connect to switch (or close + reopen Word)."
) else (
    call :log "How to use:"
    call :log "  1. Open Word normally (or keep it open)."
    call :log "  2. Right-click the tray icon near the clock and pick 'Connect'."
    call :log "     The tray will relaunch Word via the wrapper with RTL enabled,"
    call :log "     reopening the documents you had open."
    call :log ""
    call :log "  Tip: enable 'Auto-enable at every Word launch' from the tray menu"
    call :log "  to skip Connect entirely in future sessions."
)
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
