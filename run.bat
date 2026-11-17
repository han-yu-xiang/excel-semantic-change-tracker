@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_LAUNCH="
where py >nul 2>nul
if not errorlevel 1 (
    py -3.12 --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_LAUNCH=py -3.12"
    if "%PYTHON_LAUNCH%"=="" (
        py -3.11 --version >nul 2>nul
        if not errorlevel 1 set "PYTHON_LAUNCH=py -3.11"
    )
)
if "%PYTHON_LAUNCH%"=="" (
    python --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_LAUNCH=python"
)
if "%PYTHON_LAUNCH%"=="" (
    echo Python 3.11 or 3.12 was not found.
    echo Install 64-bit Python 3.11 or 3.12 and enable the Python Launcher or PATH option.
    pause
    exit /b 1
)
%PYTHON_LAUNCH% -c "import sys; raise SystemExit(0 if sys.version_info[:2] in ((3, 11), (3, 12)) else 1)"
if errorlevel 1 (
    echo Python 3.11 or 3.12 is required. The selected interpreter is not supported.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating a private project environment...
    %PYTHON_LAUNCH% -m venv .venv
    if errorlevel 1 goto :install_failed
    echo Installing required Python packages. This may take several minutes...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    if errorlevel 1 goto :install_failed
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :install_failed
)

".venv\Scripts\python.exe" dependency_check.py
if errorlevel 1 goto :repair_dependencies
goto :dependencies_ready

:repair_dependencies
echo The project environment is incomplete. Repairing required packages...
".venv\Scripts\python.exe" -m ensurepip --upgrade
if errorlevel 1 goto :install_failed
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :install_failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :install_failed
".venv\Scripts\python.exe" dependency_check.py
if errorlevel 1 goto :install_failed

:dependencies_ready

echo Starting Excel Semantic Change Tracker...
".venv\Scripts\python.exe" main.py
if errorlevel 1 (
    echo.
    echo The application closed with an error. Review the message above.
    pause
    exit /b 1
)
exit /b 0

:install_failed
echo.
echo Dependency installation failed. Check your internet connection and try again.
pause
exit /b 1

