# Copilot Instructions — RMT Margin Analysis Tool

This repository contains a Windows Python tool that extracts `START_RMT`
(Rank Margin Tool) blocks from Intel MRC debug logs and produces CSV, Excel,
HTML, PowerPoint and optional JMP Pro charts.

## Running the tool from chat

When the user asks to plot, chart, analyze or report RMT margins, follow the
[rmt-margin-plots skill](skills/rmt-margin-plots/SKILL.md). In short:

- Use `<repo>\.venv\Scripts\python.exe` (created by `setup.bat`), never bare `python`.
- Run only `rmt_log_pipeline.py`, always with `--no-ask-chart-approval`.
- Never launch `rmt_gui.py`, `Launch_RMT_GUI.bat` or `rmt_pipeline_runner.py` —
  they are interactive and will hang an agent session.
- Never invent input/output paths; ask the user.
- Resolve every path from `git rev-parse --show-toplevel`; never hardcode
  user-specific paths.

## Editing the code

- Keep the tool self-contained: no imports or file references outside this
  repository other than the pinned packages in `requirements.txt`, the Python
  standard library, and the user's own JMP Pro install.
- Add any new third-party dependency to `requirements.txt` with a pinned version.
- Use `pathlib.Path` for file paths and f-strings for formatting; keep Python
  3.10 compatibility.
- File I/O must handle `FileNotFoundError` / `PermissionError` with clear messages;
  `subprocess` calls must check return codes and must not use `shell=True`.
- Keep the documented exit codes (0-5) stable — the agent skill depends on them.
