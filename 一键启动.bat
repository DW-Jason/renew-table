@echo off
setlocal
cd /d "%~dp0"

set "APP_SCRIPT="
for %%F in (*.py) do (
    set "APP_SCRIPT=%%~fF"
    goto :found_script
)

echo Cannot find the Python app file in this folder.
pause
exit /b 1

:found_script
set "PYTHON_CMD="

where py >nul 2>nul
if not errorlevel 1 (
    py -3.8 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3.8"
        goto :run_app
    )

    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py -3"
        goto :run_app
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :run_app
)

echo Python was not found.
echo Please install Python 3.8, then run this file again.
pause
exit /b 1

:run_app
%PYTHON_CMD% "%APP_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo The app exited with error code %EXIT_CODE%.
    echo If dependencies are missing, run:
    echo   %PYTHON_CMD% -m pip install -r requirements.txt
    pause
)
exit /b %EXIT_CODE%
