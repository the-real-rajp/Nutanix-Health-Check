@echo off
setlocal

cd /d "%~dp0"

if not exist "output" mkdir "output"
if not exist "output\logs" mkdir "output\logs"

"NutanixHealthCheck.exe" --output-dir "%~dp0output"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo Nutanix Health Check completed.
) else (
    echo Nutanix Health Check exited with code %EXIT_CODE%.
)
echo Reports, JSON files, and logs are located in:
echo %~dp0output
echo.
pause

exit /b %EXIT_CODE%
