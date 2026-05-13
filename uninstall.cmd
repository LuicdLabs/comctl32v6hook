@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SHOULD_PAUSE=1"
if defined COMCTL32V6HOOK_NO_PAUSE set "SHOULD_PAUSE=0"

net session >nul 2>nul
if not "%ERRORLEVEL%"=="0" (
    echo Requesting administrator permission...
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath $env:ComSpec -ArgumentList '/d /c ""%~f0"" %*' -Verb RunAs"
    exit /b 0
)

pushd "%SCRIPT_DIR%" >nul
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\uninstall.ps1" %*
set "CODE=%ERRORLEVEL%"
popd >nul

if not "%CODE%"=="0" (
    echo.
    echo Uninstall failed with exit code %CODE%.
)

if "%SHOULD_PAUSE%"=="1" (
    echo.
    pause
)

exit /b %CODE%
