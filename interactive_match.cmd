@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
title Activity Matcher Interactive Test
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0interactive_match.ps1" %*
set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" echo Interactive tester stopped with an error.
pause
exit /b %RESULT%
