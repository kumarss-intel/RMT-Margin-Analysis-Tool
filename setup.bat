@echo off
REM ===================================================================
REM  RMT Margin Analysis Tool - Environment Setup
REM
REM  Creates a local virtual environment (.venv) next to this script and
REM  installs every Python dependency listed in requirements.txt.
REM
REM  JMP Pro is intentionally NOT installed here: it is a licensed
REM  Windows desktop application. This script only detects the existing
REM  installation and reports where it found it.
REM
REM  Usage:
REM     setup.bat              Create/refresh the environment
REM     setup.bat --recreate   Delete .venv and build it from scratch
REM     setup.bat --help       Show this help
REM ===================================================================
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "VENV_DIR=%SCRIPT_DIR%\.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "REQ_FILE=%SCRIPT_DIR%\requirements.txt"
set "RECREATE=0"

:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--recreate" set "RECREATE=1"
if /I "%~1"=="-r"         set "RECREATE=1"
if /I "%~1"=="--help"     goto usage
if /I "%~1"=="-h"         goto usage
if /I "%~1"=="/?"         goto usage
shift
goto parse_args
:args_done

echo.
echo ===================================================================
echo   RMT Margin Analysis Tool - Setup
echo ===================================================================
echo   Tool folder : %SCRIPT_DIR%
echo   Virtual env : %VENV_DIR%
echo.

if not exist "%REQ_FILE%" (
    echo [ERROR] requirements.txt not found next to setup.bat:
    echo         %REQ_FILE%
    goto fail
)

REM ---------------------------------------------------------------
REM  Step 1/5 - Locate a suitable Python interpreter (3.10 or newer)
REM ---------------------------------------------------------------
echo [1/5] Locating a Python 3.10+ interpreter...
set "BASE_PY="

call :try_python "C:\Program Files\Python314\python.exe"
if not defined BASE_PY call :try_python "C:\Program Files\Python313\python.exe"
if not defined BASE_PY call :try_python "C:\Program Files\Python312\python.exe"

if not defined BASE_PY (
    for /f "delims=" %%P in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do (
        call :try_python "%%P"
    )
)

if not defined BASE_PY (
    for /f "delims=" %%P in ('where python 2^>nul') do (
        call :try_python "%%P"
    )
)

if not defined BASE_PY (
    echo.
    echo [ERROR] No Python 3.10 or newer interpreter was found.
    echo         Install Python for Windows ^(3.14 recommended, with the
    echo         "tcl/tk and IDLE" option enabled^) and re-run setup.bat.
    echo         https://www.python.org/downloads/windows/
    goto fail
)

set "BASE_VER=unknown"
set "_VERFILE=%TEMP%\_rmt_setup_pyver.txt"
"%BASE_PY%" -c "import sys; print(sys.version.split()[0])" > "%_VERFILE%" 2>nul
if exist "%_VERFILE%" (
    set /p BASE_VER=<"%_VERFILE%"
    del "%_VERFILE%" >nul 2>&1
)
echo       Using Python !BASE_VER!  ^(%BASE_PY%^)

REM tkinter ships with CPython on Windows but can be skipped at install time.
"%BASE_PY%" -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] This Python has no tkinter module, so the RMT GUI cannot run.
    echo         Re-run the Python installer and enable "tcl/tk and IDLE".
    goto fail
)

REM ---------------------------------------------------------------
REM  Step 2/5 - Create (or reuse) the virtual environment
REM ---------------------------------------------------------------
echo.
echo [2/5] Preparing virtual environment...

if "%RECREATE%"=="1" (
    if exist "%VENV_DIR%" (
        echo       --recreate requested, removing existing .venv...
        rmdir /s /q "%VENV_DIR%"
    )
)

if exist "%VENV_PY%" (
    echo       Existing .venv found - reusing it.
) else (
    echo       Creating .venv ^(this can take a few seconds^)...
    "%BASE_PY%" -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to create the virtual environment.
        goto fail
    )
)

if not exist "%VENV_PY%" (
    echo.
    echo [ERROR] .venv appears to be corrupt ^(python.exe missing^).
    echo         Re-run:  setup.bat --recreate
    goto fail
)

REM ---------------------------------------------------------------
REM  Step 3/5 - Install the pinned dependencies
REM ---------------------------------------------------------------
echo.
echo [3/5] Installing dependencies from requirements.txt...
echo       ^(behind the Intel network a proxy may be required, e.g.
echo        set HTTPS_PROXY=http://proxy-chain.intel.com:912^)
echo.

"%VENV_PY%" -m pip install --upgrade pip --disable-pip-version-check
if errorlevel 1 (
    echo.
    echo [WARN] Could not upgrade pip - continuing with the bundled version.
)

"%VENV_PY%" -m pip install --disable-pip-version-check -r "%REQ_FILE%"
if errorlevel 1 (
    echo.
    echo [ERROR] Dependency installation failed.
    echo         Most common cause: no network / proxy not configured.
    echo         Try:
    echo             set HTTPS_PROXY=http://proxy-chain.intel.com:912
    echo             set HTTP_PROXY=http://proxy-chain.intel.com:912
    echo             setup.bat
    goto fail
)

REM ---------------------------------------------------------------
REM  Step 4/5 - Verify the environment can import everything
REM ---------------------------------------------------------------
echo.
echo [4/5] Verifying installed packages...
"%VENV_PY%" -c "import openpyxl, pptx, tkinter; print('      openpyxl  ' + openpyxl.__version__); print('      python-pptx ready'); print('      tkinter   ready')"
if errorlevel 1 (
    echo.
    echo [ERROR] A required package could not be imported.
    echo         Re-run:  setup.bat --recreate
    goto fail
)

"%VENV_PY%" -c "import PIL; print('      pillow    ' + PIL.__version__)" 2>nul
if errorlevel 1 echo       pillow    NOT installed ^(optional - GUI screenshot capture disabled^)

REM ---------------------------------------------------------------
REM  Step 5/5 - Detect the JMP installation (never installed by us)
REM ---------------------------------------------------------------
echo.
echo [5/5] Detecting JMP installation ^(provided by Windows, not by pip^)...
call :detect_jmp

if defined JMP_EXE (
    echo       Found JMP: !JMP_EXE!
) else (
    echo.
    echo [WARN] No JMP installation was detected on this machine.
    echo        The tool still runs: log parsing, CSV, Excel and the HTML
    echo        report work without JMP, and "--jmp-jsl-only" emits the JSL
    echo        script for a machine that does have JMP.
    echo        JMP chart / PPT stages need JMP Pro installed locally, e.g.
    echo            C:\Program Files\SAS\JMPPRO\17\jmp.exe
    echo        Install it from the Intel software portal, then point the
    echo        GUI field "JMP executable" at it.
)

echo.
echo ===================================================================
echo   SETUP COMPLETE
echo ===================================================================
echo.
echo   Launch the GUI          :  Launch_RMT_GUI.bat
echo   Interactive CLI wrapper :  .venv\Scripts\python.exe rmt_pipeline_runner.py
echo   Raw pipeline CLI        :  .venv\Scripts\python.exe rmt_log_pipeline.py --help
echo.
echo   Activate the venv manually in a shell:
echo       .venv\Scripts\activate.bat
echo.
endlocal
exit /b 0

REM ===================================================================
REM  Subroutines
REM ===================================================================

:try_python
REM %~1 = candidate interpreter. Accepts it only when it exists, is not the
REM Microsoft Store alias stub, and reports version >= 3.10.
if defined BASE_PY goto :eof
if "%~1"=="" goto :eof
echo %~1 | findstr /I "WindowsApps" >nul && goto :eof
if not exist "%~1" goto :eof
"%~1" -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 goto :eof
set "BASE_PY=%~1"
goto :eof

:detect_jmp
REM Registry "App Paths" is the authoritative record written by the JMP
REM installer; the hard-coded folders and PATH lookup are fallbacks.
set "JMP_EXE="
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe" /ve 2^>nul ^| find "REG_SZ"') do (
    if not defined JMP_EXE set "JMP_EXE=%%B"
)
if defined JMP_EXE if not exist "!JMP_EXE!" set "JMP_EXE="

if not defined JMP_EXE (
    for %%V in (19 18 17 16) do (
        if not defined JMP_EXE if exist "C:\Program Files\SAS\JMPPRO\%%V\jmp.exe" set "JMP_EXE=C:\Program Files\SAS\JMPPRO\%%V\jmp.exe"
        if not defined JMP_EXE if exist "C:\Program Files\SAS\JMP\%%V\jmp.exe"    set "JMP_EXE=C:\Program Files\SAS\JMP\%%V\jmp.exe"
    )
)

if not defined JMP_EXE (
    for /f "delims=" %%P in ('where jmp.exe 2^>nul') do (
        if not defined JMP_EXE set "JMP_EXE=%%P"
    )
)
goto :eof

:usage
echo.
echo RMT Margin Analysis Tool - setup.bat
echo.
echo   setup.bat              Create .venv and install requirements.txt
echo   setup.bat --recreate   Delete .venv first, then rebuild it
echo   setup.bat --help       Show this help
echo.
echo JMP Pro is never installed by this script - it must already be present
echo as a Windows application. setup.bat only detects and reports it.
echo.
endlocal
exit /b 0

:fail
echo.
echo ===================================================================
echo   SETUP FAILED - see the error above.
echo ===================================================================
echo.
endlocal
exit /b 1
