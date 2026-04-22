@echo off
setlocal EnableDelayedExpansion

set "APP_DIR=%~dp0"
set "VENV_DIR=%APP_DIR%.venv"
set "PYTHON=%VENV_DIR%\Scripts\python.exe"
set "PIP=%VENV_DIR%\Scripts\pip.exe"

:: -----------------------------------------------------------------------
:: UPDATE  (git-based for now — replace this block when moving off Git)
:: -----------------------------------------------------------------------
echo [launcher] Checking for updates...
cd /d "%APP_DIR%"

where git >nul 2>&1
if !errorlevel! neq 0 (
    echo [launcher] git not found - skipping update step.
    goto :setup_venv
)

git reset --hard
git pull

:: -----------------------------------------------------------------------
:: VENV SETUP
:: -----------------------------------------------------------------------
:setup_venv
if exist "%PYTHON%" goto :install_deps

:: Find a usable Python: try Windows Launcher (py) first, then python
set "SYSTEM_PYTHON="
where py >nul 2>&1
if !errorlevel! equ 0 set "SYSTEM_PYTHON=py"

if "!SYSTEM_PYTHON!"=="" (
    where python >nul 2>&1
    if !errorlevel! equ 0 set "SYSTEM_PYTHON=python"
)

if "!SYSTEM_PYTHON!"=="" (
    echo [launcher] ERROR: Python not found on PATH.
    echo             Install Python from https://www.python.org
    echo             Tick "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo [launcher] Creating virtual environment using !SYSTEM_PYTHON!...
!SYSTEM_PYTHON! -m venv "%VENV_DIR%"
if !errorlevel! neq 0 (
    echo [launcher] ERROR: Could not create virtual environment.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: DEPENDENCIES
:: -----------------------------------------------------------------------
:install_deps
echo [launcher] Installing / verifying dependencies...
"%PIP%" install -q -r "%APP_DIR%requirements.txt"
if !errorlevel! neq 0 (
    echo [launcher] ERROR: pip install failed.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: LAUNCH
:: -----------------------------------------------------------------------
echo [launcher] Starting app...
"%PYTHON%" "%APP_DIR%main.py"

set "EXIT_CODE=!errorlevel!"
if !EXIT_CODE! neq 0 (
    echo.
    echo [launcher] App exited with an error - exit code !EXIT_CODE!
    pause
)

endlocal
