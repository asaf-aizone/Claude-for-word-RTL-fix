@echo off
REM Claude for Word RTL Fix - cleanup helper.
REM Stops all node.exe processes that could be running inject.js and verifies
REM port 9222 is no longer listening. Does NOT close Word.

setlocal

echo ================================================================
echo  Claude for Word RTL Fix - Cleanup
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
REM Check if anything still listens on 9222
netstat -an | find "9222" | find "LISTENING" >nul
if %ERRORLEVEL% equ 0 (
    echo [WARN] Port 9222 is still listening.
    echo This means Word is still running with the debug flag enabled.
    echo To fully close the debug port, close Word.
) else (
    echo Port 9222 is closed. Debug surface is no longer exposed.
)

echo.
echo Cleanup complete.
pause
endlocal
