@echo off
REM Claude for Word RTL - uninstaller.
REM
REM Tray-only model: removes the Startup-folder entry and stops any running
REM tray/injector, clears the Auto-enable environment variable (only when
REM it matches our exact value), removes the Apps and Features entry, and
REM cleans node_modules plus temp status files.

setlocal EnableDelayedExpansion

set "HERE=%~dp0"
if "%HERE:~-1%"=="\" set "HERE=%HERE:~0,-1%"

echo ================================================================
echo  Claude for Word RTL - Uninstaller
echo ================================================================
echo.

REM --- Check that Word is not running ---
:check_word
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 goto word_closed
echo   [WAIT] Microsoft Word is running. Please close all Word windows, then press any key to continue.
pause >nul
goto check_word
:word_closed
echo   [OK] Word is not running.
echo.

REM ---------------------------------------------------------------
REM [1/4] Remove Startup folder entry
REM ---------------------------------------------------------------
set "STARTUP_LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk"
if exist "%STARTUP_LNK%" (
    del "%STARTUP_LNK%" >nul 2>&1
    echo [1/4] Removed Startup entry ^(tray will no longer auto-launch at login^).
) else (
    echo [1/4] No Startup entry found ^(nothing to remove^).
)

REM ---------------------------------------------------------------
REM [2/4] Stop the tray and the injector.
REM ---------------------------------------------------------------
echo [2/4] Stopping tray and injector...

REM Stop only OUR injector via PID file (do not kill unrelated node.exe).
set "PIDFILE=%TEMP%\claude-word-rtl.pid"
if exist "%PIDFILE%" (
    for /f "usebackq delims=" %%P in ("%PIDFILE%") do taskkill /F /PID %%P >nul 2>&1
    del /q "%PIDFILE%" >nul 2>&1
    echo   Stopped running injector.
)

REM Stop the tray. Preferred path: its PID file.
REM Fallback: match by command line via WMI so we do not touch unrelated
REM PowerShell sessions. Kept for safety - the tray always writes a PID
REM file on start, but if that write ever failed, this path recovers.
set "TRAYPIDFILE=%TEMP%\claude-word-rtl.tray.pid"
if exist "%TRAYPIDFILE%" (
    for /f "usebackq delims=" %%P in ("%TRAYPIDFILE%") do taskkill /F /PID %%P >nul 2>&1
    del /q "%TRAYPIDFILE%" >nul 2>&1
    echo   Stopped running tray.
) else (
    REM Fallback when no PID file was written (recovery path).
    REM Use PowerShell Stop-Process so the pipe stays inside PS and does not
    REM get re-parsed by cmd in a for /f subshell.
    powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | Where-Object { $_.CommandLine -like '*tray-icon.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Write-Host ('  Stopped legacy tray PID ' + $_.ProcessId) }"
)

if exist "%TEMP%\claude-word-rtl.lock"   del /q "%TEMP%\claude-word-rtl.lock"   >nul 2>&1
if exist "%TEMP%\claude-word-rtl.status" del /q "%TEMP%\claude-word-rtl.status" >nul 2>&1
REM v0.2.0+ per-app status file. Always remove on uninstall so a stale
REM JSON cannot mislead a future tray (from a separate install) into
REM rendering wrong status labels.
if exist "%TEMP%\claude-office-rtl.apps.json" del /q "%TEMP%\claude-office-rtl.apps.json" >nul 2>&1
REM v0.3.0+ Outlook opt-in flag. The injector clears it at startup too,
REM but a clean uninstall should leave nothing behind in case a future
REM install path differs. Same for the Disconnect-only IPC request file
REM - if the user uninstalls mid-request, do not leave it around to
REM trigger a phantom detach on the next injector run.
if exist "%TEMP%\claude-office-rtl.outlook-optin"             del /q "%TEMP%\claude-office-rtl.outlook-optin"             >nul 2>&1
if exist "%TEMP%\claude-office-rtl.disconnect-outlook.request" del /q "%TEMP%\claude-office-rtl.disconnect-outlook.request" >nul 2>&1
echo   [OK] Tray and injector stopped.

REM ---------------------------------------------------------------
REM Remove legacy Auto-enable env var (only if its value matches one of
REM our known strings, so we never clobber a variable the user set for
REM a different purpose). v0.1.x persisted '--remote-debugging-port=9222',
REM v0.2.0 development builds briefly persisted '--remote-debugging-port=0'.
REM Both are cleaned here. v0.2.0+ release builds do not write this key
REM at all; it is removed only as cleanup of older state.
REM ---------------------------------------------------------------
powershell -NoProfile -Command "try { $v = (Get-ItemProperty -Path 'HKCU:\Environment' -Name 'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS' -ErrorAction Stop).'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'; if ($v -eq '--remote-debugging-port=9222' -or $v -eq '--remote-debugging-port=0') { Remove-ItemProperty -Path 'HKCU:\Environment' -Name 'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS' -ErrorAction SilentlyContinue; Add-Type -Namespace W -Name E -MemberDefinition '[System.Runtime.InteropServices.DllImport(\"user32.dll\")] public static extern System.IntPtr SendMessageTimeout(System.IntPtr h,uint m,System.UIntPtr w,string l,uint f,uint t,out System.UIntPtr r);'; $r = [System.UIntPtr]::Zero; [W.E]::SendMessageTimeout([System.IntPtr]0xFFFF, 0x1A, [System.UIntPtr]::Zero, 'Environment', 2, 5000, [ref]$r) | Out-Null; Write-Host '  Removed legacy Auto-enable env var.' } } catch {}"

REM ---------------------------------------------------------------
REM [3/4] Remove Apps and Features registration
REM ---------------------------------------------------------------
set "REGKEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL"
reg query "%REGKEY%" >nul 2>&1
if not errorlevel 1 (
    reg delete "%REGKEY%" /f >nul 2>&1
    echo [3/4] Removed Apps and Features registration.
) else (
    echo [3/4] No Apps and Features registration found ^(nothing to remove^).
)

REM ---------------------------------------------------------------
REM [4/4] Remove node_modules
REM ---------------------------------------------------------------
echo [4/4] Removing installed dependencies...
if exist "%HERE%\scripts\node_modules" (
    rmdir /S /Q "%HERE%\scripts\node_modules"
    echo   [OK] node_modules removed.
) else (
    echo   [OK] No node_modules to remove.
)

echo.
echo ================================================================
echo  Uninstall complete.
echo ================================================================
echo.
echo You can now safely delete this folder.
echo Word itself was not modified.
echo.
pause
endlocal
