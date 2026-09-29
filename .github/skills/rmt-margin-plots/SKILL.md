---
name: rmt-margin-plots
description: Run the RMT Margin Analysis Tool to extract START_RMT blocks from MRC debug logs and produce CSV, Excel, JMP scatter charts, an HTML report, and a PowerPoint deck. Use when asked to plot/chart/analyze RMT margins, RxDqVrefByte/TxVref/RecEnDelay margins, build RMT slides, or convert lab logs into margin charts.
---

# RMT Margin Plotting

Drives the RMT Margin Analysis Tool — the scripts at the **root of this
repository** — non-interactively, so margin plots can be produced directly from
a chat request. The skill is self-contained: it depends only on files inside this
repository plus the machine's own Python / optional JMP Pro install.

## Tool layout

| Path | Role |
|------|------|
| `rmt_log_pipeline.py` | Core CLI engine — **this is what the agent runs** |
| `rmt_pipeline_runner.py` | Interactive menu wrapper — **never run from the agent** (it blocks on `input()`) |
| `rmt_gui.py` / `Launch_RMT_GUI.bat` | Tk GUI — **never run from the agent** |
| `setup.bat` | One-time `.venv` creation + dependency install |
| `jmp_axis_settings.json` | Per-parameter Y-axis min/max/inc/ref-line config |

All paths above are relative to the repository root.

## Rules

1. **Always use the tool's own interpreter**, never bare `python`:
   `<repo>\.venv\Scripts\python.exe`.
   Resolve `<repo>` with `git rev-parse --show-toplevel` (see step 2) — **never
   hardcode a user-specific path**, since clone locations differ per machine.
   If the interpreter is missing, run `setup.bat` first.
2. **Always pass `--no-ask-chart-approval`.** The pipeline otherwise prompts on
   stdin and the run will hang. Get the user's approval with `ask_user` *before*
   launching instead.
3. **Run with a long `initial_wait`** (180+ s). JMP runs take 3-5 minutes
   because JMP launches as a desktop app.
4. **Never invent paths.** Input log folder and output folder come from the user.
   If either is missing, ask with `ask_user` — do not guess a `c:\lab_Logs\...` path.
5. JMP charts require `--jmp-exe`; there is no auto-detection inside the pipeline.
   Resolve it first (see below) and fall back to base mode if JMP is absent.

## What actually gets produced

The tool has only **two** real modes. The "Stage 1/2/3/4" wording in the READMEs
describes the interactive wrapper's menu, not the CLI — the core engine always
writes the full CSV + Excel + HTML + native PPT set.

| Mode | Flags | Output |
|------|-------|--------|
| **Base** (no JMP needed) | *(none extra)* | `RMT_Combined_Similar.csv`, `RMT_Combined_Extended.csv`, `csv_by_file/`, `RMT_Extraction.xlsx`, `RMT_Report.html`, `RMT_Summary.pptx`, `rmt_metadata.json` |
| **JMP** | `--generate-jmp-charts --jmp-exe <path>` | everything above **plus** `rmt_jmp_charts.jsl`, `jmp_charts/*.png`, `RMT_Summary_JMP_Charts.pptx` |

`--ppt-template <pptx>` (optional) styles the JMP deck from an existing template.

> `--chart-fields` only *selects which parameters are charted*. It does **not**
> disable charting — an empty or omitted value means **all** parameters. There is
> no flag to skip the base PPT/HTML.

Chart parameters (`--chart-fields`, comma separated; all of them is the default):
`RecEnDelay, TxDqsDelay, RxDqsDelay, TxDqDelay, RxDqVrefByte, TxVref, ClkGrpPi, CmdVref`

### Thermal (DTR) logs

Add `--dtr` when the logs are Thermal Experiment / Temp-Drift profiles (BCRH / BHRC).
Each log then yields two data sets — a Boot-temperature RMT (under the
`Temp Drift RMT Test` task) and a Run-temperature RMT (under `Rank Margin Tool`) —
with `BootTemp` / `RunTemp` snapped to the nearest setpoint (0 or 90).
If the user mentions temperature drift, hot/cold, BCRH or BHRC, ask whether to use `--dtr`.

## Procedure

### 1. Confirm inputs

Use `ask_user` for anything not supplied: input log folder/file, output folder,
whether JMP charts are wanted, and which parameters to chart. Default the output
folder to the input folder if the user has no preference.

### 2. Resolve the tool folder, interpreter, and JMP

The tool lives at the repository root, so the git root **is** the tool folder.
Run this from anywhere inside the repo:

```powershell
$tool = (Resolve-Path (git rev-parse --show-toplevel)).Path   # git prints forward slashes; normalize
$py   = Join-Path $tool ".venv\Scripts\python.exe"
$pipe = Join-Path $tool "rmt_log_pipeline.py"
$axis = Join-Path $tool "jmp_axis_settings.json"
Test-Path $py, $pipe

$jmp = @(
  (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe" -ErrorAction SilentlyContinue).'(default)',
  "C:\Program Files\SAS\JMPPRO\18\jmp.exe",
  "C:\Program Files\SAS\JMPPRO\17\jmp.exe",
  "C:\Program Files\SAS\JMP\18\jmp.exe",
  "C:\Program Files\SAS\JMP\17\jmp.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
$jmp
```

If `git rev-parse` fails (shell started outside the repo), fall back to the
workspace root reported in the environment context and confirm
`rmt_log_pipeline.py` exists there — still never a hardcoded user-specific path.

If `$py` is missing: run `& (Join-Path $tool "setup.bat")` (add `--recreate` if
the venv is broken).
If `$jmp` is empty: tell the user JMP Pro was not found and run base mode (CSV +
Excel + HTML + native PPT still work without JMP).

> Each `powershell` call starts a fresh process, so `$tool`/`$py`/`$jmp` do **not**
> persist between calls. Re-derive them at the top of every command block, or keep
> the resolution and the run in a single call.

### 3. Sanity-check the logs before a long run

```powershell
Select-String -Path "<input>\*.txt" -Pattern "START_RMT" | Select-Object -First 3
```

No matches means the logs hold no RMT data — report that instead of running the pipeline.

### 4. Run

> **PowerShell gotcha:** `--chart-fields ""` **fails** — PowerShell drops the empty
> argument and argparse errors with `expected one argument`. Use the glued form
> `"--chart-fields="` instead, or just omit the flag entirely (same effect: all params).

Base run (fast, no JMP required) — self-contained, re-derives paths:

```powershell
$tool = (Resolve-Path (git rev-parse --show-toplevel)).Path
& (Join-Path $tool ".venv\Scripts\python.exe") (Join-Path $tool "rmt_log_pipeline.py") `
  --input "<input folder>" `
  --pattern "*.txt" `
  --outdir "<output folder>" `
  --no-ask-chart-approval
```

Full run with JMP charts:

```powershell
$tool = (Resolve-Path (git rev-parse --show-toplevel)).Path
$jmp  = @(
  (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe" -ErrorAction SilentlyContinue).'(default)',
  "C:\Program Files\SAS\JMPPRO\18\jmp.exe",
  "C:\Program Files\SAS\JMPPRO\17\jmp.exe",
  "C:\Program Files\SAS\JMP\18\jmp.exe",
  "C:\Program Files\SAS\JMP\17\jmp.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
& (Join-Path $tool ".venv\Scripts\python.exe") (Join-Path $tool "rmt_log_pipeline.py") `
  --input "<input folder>" `
  --pattern "*.txt" `
  --outdir "<output folder>" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "$jmp" `
  --jmp-axis-config (Join-Path $tool "jmp_axis_settings.json")
```

Add `--ppt-template "<deck.pptx>"` when the user wants corporate slide styling.

If `--ask-chart-approval` is left on (the default) and there is no TTY — which is
always the case for agent-run commands — the tool prints a warning and silently
charts **nothing**. That is why `--no-ask-chart-approval` is mandatory here.

### 5. Re-use existing artifacts instead of re-parsing

| Situation | Flag |
|-----------|------|
| CSV already extracted, only charts wanted | `--jmp-from-csv "<csv>" [more.csv ...]` (multiple CSVs are concatenated and colour-split by `SourceFile`) |
| Excel already built | `--jmp-from-excel "<xlsx>"` (reads the `All_RMT` sheet) |
| PNGs already rendered, only PPT wanted | `--ppt-from-charts "<outdir>\jmp_charts"` |
| Want the JSL without launching JMP | `--jmp-jsl-only` |

All of these still need `--outdir`, and the JMP variants still need `--jmp-exe`.

### 6. Report back

List the generated artifacts with their full paths and flag anything skipped
(e.g. "JMP not installed — Excel + HTML only"). Do not open the PPT/GUI.

## Exit codes

| Code | Meaning | Action |
|------|---------|--------|
| 0 | Success | Report artifacts |
| 1 | No input files matched `--input` / `--pattern` | Re-check the path and pattern with the user |
| 2 | Input CSV / charts dir / Excel not found, or bad CLI arguments | Re-check the path with the user |
| 3 | Chart approval denied (interactive mode only) | Always pass `--no-ask-chart-approval` |
| 4 | JMP chart generation requested without `--jmp-exe` | Resolve JMP or run base mode |
| 5 | JMP failed to generate or run the charts | Report the error; the JSL is kept in `<outdir>` for a manual run |

## Axis tuning

To change Y-axis scale, tick increment, or reference lines, edit
`jmp_axis_settings.json` — each parameter has `plus`/`minus` blocks with
`min`, `max`, `inc`, `ref_line`. No code change is needed; re-run with
`--generate-jmp-charts` (or `--jmp-from-csv`) afterwards.

## Reference docs

- `README.md` — overview
- `RMT_QUICKSTART.md` — copy-paste examples
- `RMT_LOG_PIPELINE_README.md` — full CLI reference
