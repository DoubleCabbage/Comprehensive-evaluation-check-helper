@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_all_tests.ps1" %*
set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" echo Tests failed. Review the error above.
if "%RESULT%"=="0" echo Tests completed successfully.
pause
exit /b %RESULT%
