@echo off
REM Claude for Word RTL Fix - single-step launcher.
REM Closes existing Word (warns first), launches Word with WebView2 debug port,
REM then runs the injector. Keep this window open while using Word.

setlocal EnableDelayedExpansion

echo ================================================================
echo  Claude for Word RTL Fix
echo ================================================================
echo.

REM Check Node.js is installed
where node >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js is not installed or not in PATH.
    echo Please install Node.js from https://nodejs.org/ ^(version 16 or later^).
    echo.
    pause
    exit /b 1
)

REM Warn and offer to close Word if running
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if !ERRORLEVEL! equ 0 (
    echo [WARNING] Word is currently running.
    echo Word must be closed and reopened for the debug port to take effect.
    echo.
    set /p CLOSE="Close Word now? [y/N] "
    if /I "!CLOSE!"=="y" (
        taskkill /F /IM WINWORD.EXE >nul 2>&1
        timeout /t 2 /nobreak >nul
        echo Word closed.
        echo.
    ) else (
        echo Aborting. Please close Word manually and run again.
        pause
        exit /b 1
    )
)

REM Locate WINWORD.EXE
set "WINWORD=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" set "WINWORD=C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
if not exist "%WINWORD%" (
    echo [ERROR] Could not find WINWORD.EXE in the default Office 16 locations.
    echo Please edit start.bat and set WINWORD to the correct path.
    pause
    exit /b 1
)

REM Install dependencies if needed
if not exist "%~dp0scripts\node_modules\chrome-remote-interface" (
    echo [INFO] Installing dependencies ^(first run only^)...
    pushd "%~dp0scripts"
    call npm install --silent
    if errorlevel 1 (
        echo [ERROR] npm install failed.
        popd
        pause
        exit /b 1
    )
    popd
    echo.
)

REM Enable WebView2 debug port for this session
set "WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222"

echo [1/2] Launching Word with debug port enabled...
start "" "%WINWORD%"

echo [2/2] Starting RTL injector...
echo.
echo ----------------------------------------------------------------
echo Keep this window open. Minimize it if it's in your way.
echo Close it ^(or Ctrl+C^) when you are done using Word.
echo ----------------------------------------------------------------
echo.
echo In Word:
echo   1. Open or create a document ^(save it if new^).
echo   2. Open the Claude add-in panel.
echo   3. The RTL fix will apply automatically within 2 seconds.
echo.

cd /d "%~dp0scripts"
node inject.js

endlocal
