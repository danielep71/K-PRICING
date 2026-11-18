@echo off
setlocal EnableExtensions

rem Builds the K-PRICING workbook from this clone (tools\Build-KPR.ps1): the
rem macro-free template, the repository's VBA modules, the commit on the Home
rem sheet and the demo sheet, saved in the clone's ignored build folder.
rem Double-click it for the Production workbook, or pass Development to also
rem import the regression and test modules: tools\Build-KPR.cmd Development
rem The clone is only read; no GitHub download and no regression run.
rem (tools\Check-KPR.cmd checks a commit of GitHub and runs the regression.)

set "SCRIPT=%~dp0Build-KPR.ps1"
if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

set "MODE=%~1"
if not defined MODE set "MODE=Production"
if /i not "%MODE%"=="Production" if /i not "%MODE%"=="Development" (
    echo ERROR: unknown mode "%MODE%". Use Production or Development.
    echo.
    pause
    exit /b 1
)

set "REPO_ROOT=%~dp0.."

echo K-PRICING workbook build
echo.
echo Clone:
echo   %REPO_ROOT%
echo Mode:
echo   %MODE%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Mode "%MODE%" -RepoRoot "%REPO_ROOT%" -NoPause
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: the K-PRICING workbook was built in the build folder of the clone.
) else (
    echo FAILED: the build returned exit code %RC%. Copy the text above and send it.
)
echo.
pause
exit /b %RC%
