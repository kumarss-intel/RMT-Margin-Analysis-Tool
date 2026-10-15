@echo off
rem ---------------------------------------------------------------------------
rem  MarginIQ command-line entry point for cmd.exe, PowerShell and automation.
rem
rem    marginiq.cmd --project WCL --input "C:\logs" --pattern "*.log" ^
rem                 --outdir "C:\out" --no-ask-chart-approval
rem
rem  Runs rmt_log_pipeline.py with this repository's .venv Python, forces UTF-8
rem  console output (safe for redirected / piped logs) and returns the
rem  pipeline's exit code unchanged (0-5, see CLI_AUTOMATION.md).
rem  Works from any working directory; quote paths that contain spaces.
rem
rem    marginiq.cmd --guide [readme|quickstart|cli|install|pipeline|projects]
rem  opens a guide (or the projects / spec-limits page) in the default browser.
rem ---------------------------------------------------------------------------
setlocal
set "MIQ_HOME=%~dp0"
set "MIQ_PY=%MIQ_HOME%.venv\Scripts\python.exe"
if not exist "%MIQ_PY%" (
    echo [MarginIQ] Python environment not found: "%MIQ_PY%" 1>&2
    echo [MarginIQ] Run "%MIQ_HOME%setup.bat --no-prompt" once, then retry. 1>&2
    exit /b 2
)
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
if /i "%~1"=="--guide" goto :guide
"%MIQ_PY%" "%MIQ_HOME%rmt_log_pipeline.py" %*
exit /b %ERRORLEVEL%

:guide
rem Forward the raw text after --guide: %2 %3 would split "--anchor=-id" at "=".
set "MIQ_ARGS=%*"
call set "MIQ_ARGS=%%MIQ_ARGS:*--guide=%%"
"%MIQ_PY%" "%MIQ_HOME%rmt_help.py" %MIQ_ARGS%
exit /b %ERRORLEVEL%
