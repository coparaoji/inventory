@echo off
setlocal EnableDelayedExpansion

set "APP_DIR=%~dp0"
set "VENV_DIR=%APP_DIR%.venv"
set "PYTHON=%VENV_DIR%\Scripts\python.exe"
set "PIP=%VENV_DIR%\Scripts\pip.exe"
set "LOG=%APP_DIR%launcher.log"

echo [launcher] Starting — %DATE% %TIME% > "%LOG%"

:: -----------------------------------------------------------------------
:: UPDATE  (git-based for now — replace this block when moving off Git)
:: -----------------------------------------------------------------------
echo [launcher] Checking for updates...
echo [launcher] Checking for updates... >> "%LOG%"
cd /d "%APP_DIR%"

where git >nul 2>&1
if %errorlevel% neq 0 (
    echo [launcher] git not found — skipping update step.
    echo [launcher] git not found >> "%LOG%"
    goto :setup_venv
)

git reset --hard >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [launcher] WARNING: git reset --hard failed.
    echo [launcher] WARNING: git reset --hard failed >> "%LOG%"
)

git pull >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [launcher] WARNING: git pull failed. Running with local files.
    echo [launcher] WARNING: git pull failed >> "%LOG%"
)

:: -----------------------------------------------------------------------
:: VENV SETUP
:: -----------------------------------------------------------------------
:setup_venv
if exist "%PYTHON%" goto :install_deps

:: Find a usable Python — prefer the Windows Launcher (py), fall back to python
set "SYSTEM_PYTHON="
where py >nul 2>&1
if %errorlevel% equ 0 (
    set "SYSTEM_PYTHON=py"
) else (
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        set "SYSTEM_PYTHON=python"
    )
)

if "%SYSTEM_PYTHON%"=="" (
    echo [launcher] ERROR: Python not found on PATH.
    echo [launcher] ERROR: Python not found on PATH >> "%LOG%"
    echo             Install Python from https://www.python.org and re-run.
    echo             (make sure to tick "Add Python to PATH" during install)
    pause
    exit /b 1
)

echo [launcher] Using Python command: %SYSTEM_PYTHON%
echo [launcher] Using Python command: %SYSTEM_PYTHON% >> "%LOG%"

echo [launcher] Creating virtual environment...
echo [launcher] Creating virtual environment... >> "%LOG%"
%SYSTEM_PYTHON% -m venv "%VENV_DIR%" >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [launcher] ERROR: Could not create virtual environment.
    echo [launcher] ERROR: Could not create venv >> "%LOG%"
    echo             See launcher.log for details.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: DEPENDENCIES
:: -----------------------------------------------------------------------
:install_deps
echo [launcher] Installing / verifying dependencies...
echo [launcher] Installing dependencies... >> "%LOG%"
"%PIP%" install -q -r "%APP_DIR%requirements.txt" >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [launcher] ERROR: pip install failed.
    echo [launcher] ERROR: pip install failed >> "%LOG%"
    echo             See launcher.log for details.
    pause
    exit /b 1
)

:: -----------------------------------------------------------------------
:: LAUNCH
:: -----------------------------------------------------------------------
echo [launcher] Starting app...
echo [launcher] Starting app... >> "%LOG%"
"%PYTHON%" "%APP_DIR%main.py" >> "%LOG%" 2>&1

if %errorlevel% neq 0 (
    echo.
    echo [launcher] App exited with an error (code %errorlevel%).
    echo [launcher] App crashed (exit %errorlevel%) >> "%LOG%"
    echo             See launcher.log next to this file for the full traceback.
    pause
)

endlocal
