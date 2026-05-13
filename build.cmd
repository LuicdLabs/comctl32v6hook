@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SHOULD_PAUSE=1"
if defined COMCTL32V6HOOK_NO_PAUSE set "SHOULD_PAUSE=0"

pushd "%SCRIPT_DIR%" >nul
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\build.ps1" %*
set "CODE=%ERRORLEVEL%"
popd >nul

if not "%CODE%"=="0" (
    echo.
    echo Build failed with exit code %CODE%.
)

if "%SHOULD_PAUSE%"=="1" (
    echo.
    pause
)

exit /b %CODE%
