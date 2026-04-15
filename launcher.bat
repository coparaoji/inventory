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
if %errorlevel% neq 0 (
    echo [launcher] git not found — skipping update step.
    goto :setup_venv
)

git reset --hard
if %errorlevel% neq 0 (
    echo [launcher] WARNING: git reset --hard failed. Continuing anyway.
)

git pull
if %errorlevel% neq 0 (
    echo [launcher] WARNING: git pull failed. Running with local files.
)

:: -----------------------------------------------------------------------
:: VENV SETUP
:: -----------------------------------------------------------------------
:setup_venv
if exist "%PYTHON%" goto :install_deps

echo [launcher] Creating virtual environment...
python -m venv "%VENV_DIR%"
if %errorlevel% neq 0 (
    echo [launcher] ERROR: Could not create virtual environment.
    echo             Make sure Python is installed and on PATH.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: DEPENDENCIES
:: -----------------------------------------------------------------------
:install_deps
echo [launcher] Installing / verifying dependencies...
"%PIP%" install -q -r "%APP_DIR%requirements.txt"
if %errorlevel% neq 0 (
    echo [launcher] ERROR: pip install failed.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: LAUNCH
:: -----------------------------------------------------------------------
echo [launcher] Starting app...
"%PYTHON%" "%APP_DIR%main.py"

if %errorlevel% neq 0 (
    echo.
    echo [launcher] App exited with an error (code %errorlevel%).
    pause
)

endlocal
