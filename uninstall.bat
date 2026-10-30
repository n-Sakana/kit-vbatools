@echo off >nul 2>&1
@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0common\uninstall.ps1" %*
set "VBATOOLS_EXIT=%ERRORLEVEL%"
if /i not "%~1"=="-PlanOnly" if not defined VBATOOLS_HEADLESS pause
exit /b %VBATOOLS_EXIT%
