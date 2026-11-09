@echo off
setlocal EnableExtensions

rem One-click Excel check of K-PRICING (tools\Check-KPR.ps1). Double-click it
rem to check main, or pass a branch name: tools\Check-KPR.cmd my-branch
rem The modules always come from GitHub at the branch's exact commit, so the
rem local clone needs no update.

set "SCRIPT=%~dp0Check-KPR.ps1"
if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)
set "BRANCH=%~1"
if not defined BRANCH set "BRANCH=main"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Branch "%BRANCH%"
exit /b %ERRORLEVEL%
