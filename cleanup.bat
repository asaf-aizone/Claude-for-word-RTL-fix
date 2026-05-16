@echo off
REM Claude for Office RTL Fix - cleanup helper.
REM Stops all node.exe processes that could be running inject.js and verifies
REM no Office WebView2 debug ports remain exposed. Does NOT close any
REM Office app.

setlocal

echo ================================================================
echo  Claude for Office RTL Fix - Cleanup
echo ================================================================
echo.

REM Stop OUR current injector via PID file.
set "PIDFILE=%TEMP%\claude-word-rtl.pid"
if exist "%PIDFILE%" (
    echo Stopping current injector...
    for /f "usebackq delims=" %%P in ("%PIDFILE%") do taskkill /F /PID %%P >nul 2>&1
    del /q "%PIDFILE%" >nul 2>&1
)

REM Also scan for orphan node.exe processes running inject.js from prior
REM sessions. The tray used to be able to leak these if it was killed
REM before the injector could clean up its own PID file. We match by the
REM CommandLine containing "inject.js" so we never touch unrelated Node
REM processes (user may have other Node work going on concurrently).
echo Scanning for orphan injector processes...
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" | Where-Object { $_.CommandLine -like '*inject.js*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Write-Host ('  Stopped orphan injector PID ' + $_.ProcessId) }"

REM Kill tray-icon.ps1 (powershell processes whose command line contains the
REM script name). Leaves other powershell sessions untouched.
REM Uses PowerShell CIM because wmic is deprecated on Windows 11 24H2+.
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | Where-Object { $_.CommandLine -like '*tray-icon.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>&1

REM Clear transient files so next launch starts clean
if exist "%TEMP%\claude-word-rtl.status" del /q "%TEMP%\claude-word-rtl.status" >nul 2>&1
if exist "%TEMP%\claude-word-rtl.lock"   del /q "%TEMP%\claude-word-rtl.lock"   >nul 2>&1

echo.
REM Check if any Office WebView2 host still has a debug port open. v0.2.0
REM uses dynamic ports (--remote-debugging-port=0), so we look for
REM msedgewebview2.exe processes spawned by Word, Excel, or PowerPoint and
REM report any LISTENING TCP socket they own. The user does not need to
REM see specific port numbers; the message just says "still exposed" or
REM "all closed".
powershell -NoProfile -Command "$wv=Get-CimInstance Win32_Process -Filter \"Name='msedgewebview2.exe'\" -ErrorAction SilentlyContinue; if ($wv) { $offices=Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE' OR Name='POWERPNT.EXE' OR Name='OUTLOOK.EXE'\" -ErrorAction SilentlyContinue; if ($offices) { Write-Host '[WARN] One or more Office WebView2 hosts are still running with the debug flag enabled.'; Write-Host '       To fully close the debug ports, close Word, Excel, PowerPoint, and Outlook.' } else { Write-Host 'No Office app is running. Any WebView2 hosts found belong to other apps and are unaffected by this tool.' } } else { Write-Host 'No WebView2 hosts running. Debug surface is no longer exposed.' }"

echo.
echo Cleanup complete.
pause
endlocal
