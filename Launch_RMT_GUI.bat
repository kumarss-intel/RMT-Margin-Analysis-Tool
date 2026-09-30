@echo off
REM ===================================================================
REM  MarginIQ - Intel CCG CVE DDR5 RMT Margin Analysis Tool - GUI Launcher
REM  Double-click this file to open the graphical front-end.
REM
REM  Windows cannot attach an icon to a .bat file itself, so the first run
REM  creates "MarginIQ.lnk" next to this file: the same launcher carrying the
REM  MarginIQ icon (assets\marginiq.ico). Pin / copy that shortcut anywhere.
REM  setup.bat also creates the Desktop / Start menu / taskbar shortcuts.
REM
REM  Interpreter priority:
REM    1. .venv created by setup.bat   (recommended)
REM    2. C:\Program Files\Python314   (pinned system install)
REM    3. whatever python is on PATH
REM ===================================================================
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "GUI=%SCRIPT_DIR%rmt_gui.py"
set "VENV_SCRIPTS=%SCRIPT_DIR%.venv\Scripts"

set "PYTHON_EXE="
set "PYTHONW_EXE="

REM 1) Tool-local virtual environment.
if exist "%VENV_SCRIPTS%\python.exe" (
    set "PYTHON_EXE=%VENV_SCRIPTS%\python.exe"
    if exist "%VENV_SCRIPTS%\pythonw.exe" set "PYTHONW_EXE=%VENV_SCRIPTS%\pythonw.exe"
)

REM 2) Pinned system install.
if not defined PYTHON_EXE (
    if exist "C:\Program Files\Python314\python.exe" (
        set "PYTHON_EXE=C:\Program Files\Python314\python.exe"
        set "PYTHONW_EXE=C:\Program Files\Python314\pythonw.exe"
    )
)

REM 3) Anything on PATH.
if not defined PYTHON_EXE (
    set "PYTHON_EXE=python"
    set "PYTHONW_EXE=pythonw"
    echo [WARN] No .venv found. Run setup.bat once to create the virtual
    echo        environment with all required dependencies.
)

REM Icon'd twin of this launcher (created once; see header).
if not exist "%SCRIPT_DIR%MarginIQ.lnk" if exist "%SCRIPT_DIR%scripts\create_shortcuts.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%scripts\create_shortcuts.ps1" -Local >nul 2>&1
)

REM Launch without a console window when pythonw is available.
if defined PYTHONW_EXE if exist "%PYTHONW_EXE%" (
    start "" "%PYTHONW_EXE%" "%GUI%"
    goto done
)

start "" "%PYTHON_EXE%" "%GUI%"

:done
endlocal
