# MarginIQ — Command-Line & Automation Guide

Use this guide to run MarginIQ from **Command Prompt (cmd.exe)**, PowerShell, a batch file,
a scheduler, CI (Jenkins / Azure DevOps / GitHub Actions runners), or any automation tool
that can start a Windows process and read its exit code.

Everything the GUI does can be done from the command line. The GUI itself only builds the
command lines described here.

---

## Contents

1. [Prerequisites](#1-prerequisites)
2. [Entry points](#2-entry-points)
3. [Recipes (copy & paste)](#3-recipes-copy--paste)
4. [GUI ↔ CLI mapping](#4-gui--cli-mapping)
5. [Flag reference](#5-flag-reference)
6. [What `--project` changes (NVL vs WCL)](#6-what---project-changes-nvl-vs-wcl)
7. [Exit codes](#7-exit-codes)
8. [Output contract](#8-output-contract)
9. [Console messages to scrape](#9-console-messages-to-scrape)
10. [Environment variables](#10-environment-variables)
11. [Integration examples](#11-integration-examples)
12. [Rules for unattended runs](#12-rules-for-unattended-runs)
13. [Interactive command-line menu (people, not automation)](#13-interactive-command-line-menu-people-not-automation)

---

## 1. Prerequisites

| Item | How |
|------|-----|
| Repository clone | `git clone https://github.com/kumarss-intel/RMT-Margin-Analysis-Tool.git` |
| Python environment | Run **once**: `setup.bat --no-prompt`. This creates `.venv` with the pinned packages and asks no questions, so it is safe for automation. Plain `setup.bat` also asks about shortcuts. |
| JMP Pro 17 (optional) | Needed only for JMP PNG charts (`--generate-jmp-charts`, `--jmp-from-csv`, `--jmp-from-excel`). CSV, Excel, HTML and the native PPT need no JMP. |
| Updates | `git pull` in the repository. Re-run `setup.bat --no-prompt` only when `requirements.txt` changed. |

Check the installation:

```bat
C:\tools\MarginIQ\marginiq.cmd --version
```

---

## 2. Entry points

| Entry point | Use it for |
|-------------|-----------|
| **`marginiq.cmd`** (repository root) | **Recommended for automation.** Works from any working directory. Uses the repository's `.venv` Python, forces UTF-8 output (safe when stdout is redirected to a file or pipe), and returns the pipeline exit code unchanged. Takes exactly the same flags as `rmt_log_pipeline.py`. |
| `.venv\Scripts\python.exe rmt_log_pipeline.py …` | The same pipeline called directly. Set `PYTHONUTF8=1` yourself when redirecting output. |
| `rmt_pipeline_runner.py` | Interactive menu for people ([section 13](#13-interactive-command-line-menu-people-not-automation)). **Not** for automation: it waits for keyboard input. |
| **`marginiq.cmd --guide [page]`** | Opens a guide in the default browser with MarginIQ's built-in viewer (same pages as the GUI *Help* menu). Pages: `readme` (User Guide, default), `quickstart`, `cli` (this guide), `install`, `pipeline` (flag manual), `projects` (Supported Projects & Spec Limits, generated from `projects.json`). Add `--anchor=<section-id>` to jump to a section (use the `=` form: ids such as `-whats-new-in-20` start with a dash), or `--no-open` to only build the HTML (printed path; exit 0, 2 for an unknown page). Same as `.venv\Scripts\python.exe rmt_help.py <page>`. |
| `rmt_gui.py`, `Launch_RMT_GUI.bat`, `MarginIQ.lnk` | The GUI. **Never** start these from automation; they wait for a user. |

In the examples below `%MIQ%` is the full path of `marginiq.cmd`, e.g.

```bat
set "MIQ=C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd"
```

A command is **one line**. The examples wrap long commands with the cmd.exe continuation
character `^` (PowerShell uses a back-tick `` ` ``). Always quote paths that contain spaces.

---

## 3. Recipes (copy & paste)

### 3.1 Non-thermal logs (single boot)

```bat
:: NVL (NVL-S / NVL-HX, DDR5) - CSV + Excel + HTML + native PPT, no JMP
call "%MIQ%" --project NVL --input "D:\lab\NVL\run42" --pattern "*.log" ^
  --outdir "D:\results\NVL_run42" --no-ask-chart-approval --no-update-check

:: WCL LPDDR5 RVP (memory type and frequency are read from the log)
call "%MIQ%" --project WCL --input "D:\lab\WCL\teraterm_17092026.log" ^
  --outdir "D:\results\WCL_LP5_7467" --no-ask-chart-approval --no-update-check

:: WCL DDR5 RVP
call "%MIQ%" --project WCL --input "D:\lab\WCL\WCL_DDR5_RVP\5600_g4.log" ^
  --outdir "D:\results\WCL_DDR5_5600" --no-ask-chart-approval --no-update-check
```

### 3.2 Thermal DTR logs (BCRH / BHRC)

```bat
:: NVL DTR: Temp Drift RMT Test + Rank Margin Tool, temps from PHY Temperature1
call "%MIQ%" --project NVL --dtr --input "D:\lab\NVL\DTR" --pattern "*.log" ^
  --outdir "D:\results\NVL_DTR" --no-ask-chart-approval --no-update-check

:: WCL DTR: first / last Rank Margin Tool run, temps from BCRH / BHRC in the file name
call "%MIQ%" --project WCL --dtr --input "D:\lab\WCL\Thermal_DTR_Logs" --pattern "*.log" ^
  --outdir "D:\results\WCL_DTR" --no-ask-chart-approval --no-update-check
```

### 3.3 With JMP charts (GUI Stage 3 / 4)

```bat
call "%MIQ%" --project WCL --input "D:\lab\WCL\run7" --pattern "*.log" ^
  --outdir "D:\results\WCL_run7" --no-ask-chart-approval --no-update-check ^
  --chart-fields "RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" ^
  --generate-jmp-charts --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" ^
  --jmp-axis-config "C:\tools\RMT-Margin-Analysis-Tool\jmp_axis_settings.json" ^
  --ppt-template "D:\templates\RMT.pptx"
```

Omit `--ppt-template` for the built-in layout. Add `--jmp-jsl-only` to write
`rmt_jmp_charts.jsl` without starting JMP (useful on machines without JMP).

### 3.4 Two-step run, like the GUI (extract once, chart later)

```bat
:: Step 1 - extract (CSV, Excel, HTML, native PPT, rmt_metadata.json)
call "%MIQ%" --project WCL --input "D:\lab\WCL\run7" --pattern "*.log" ^
  --outdir "D:\results\WCL_run7" --no-ask-chart-approval --no-update-check --chart-fields ""

:: Step 2 - JMP charts, JMP PPT and HTML with embedded charts from the Step-1 CSV.
:: Use the SAME --outdir so rmt_metadata.json (platform, memory type, training steps)
:: is reused - WCL needs it to pick the spec limits.
call "%MIQ%" --project WCL --jmp-from-csv "D:\results\WCL_run7\RMT_Combined_Extended.csv" ^
  --outdir "D:\results\WCL_run7" --no-update-check ^
  --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" ^
  --jmp-axis-config "C:\tools\RMT-Margin-Analysis-Tool\jmp_axis_settings.json"
```

### 3.5 Reuse existing outputs

```bat
:: Several CSVs merged into one comparison (rows tagged by SourceFile)
call "%MIQ%" --project WCL --jmp-from-csv "D:\r\a\RMT_Combined_Extended.csv" "D:\r\b\RMT_Combined_Extended.csv" ^
  --outdir "D:\results\a_vs_b" --no-update-check --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe"

:: From an Excel workbook (All_RMT sheet)
call "%MIQ%" --project WCL --jmp-from-excel "D:\results\WCL_run7\RMT_Extraction.xlsx" ^
  --outdir "D:\results\WCL_run7" --no-update-check --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe"

:: Rebuild the JMP-chart PowerPoint from PNGs only
call "%MIQ%" --ppt-from-charts "D:\results\WCL_run7\jmp_charts" --outdir "D:\results\WCL_run7"
```

### 3.6 One output folder per log ("Process each file SEPARATELY")

```bat
@echo off
setlocal EnableDelayedExpansion
set "MIQ=C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd"
for %%F in ("D:\lab\WCL\batch\*.log") do (
    call "%MIQ%" --project WCL --input "%%~fF" --outdir "D:\results\batch\%%~nF" ^
        --no-ask-chart-approval --no-update-check > "D:\results\batch\%%~nF.txt" 2>&1
    echo %%~nxF exit=!ERRORLEVEL!
)
```

Run JMP-enabled jobs **one at a time**. See [section 12](#12-rules-for-unattended-runs).

### 3.7 Custom ±Ref instead of the WCL spec limits

```bat
call "%MIQ%" --project WCL --input "D:\lab\WCL\run7" --outdir "D:\results\WCL_customref" ^
  --no-ask-chart-approval --no-update-check --no-rmt-limits ^
  --jmp-axis-config "D:\cfg\my_axis.json"
```

---

## 4. GUI ↔ CLI mapping

| GUI setting | CLI equivalent |
|-------------|----------------|
| Tab 1 *Project* | `--project NVL` / `--project WCL` |
| Tab 1 *Experiment profile* = Thermal (DTR) | `--dtr` |
| Tab 1 *Add files / Add folder* + *File mask* | `--input <files or folders…>` + `--pattern "*.log;*.txt"` |
| Tab 1 *Output folder* | `--outdir` |
| Stage 1 / 2 (*Generate CSV & Continue*) | `--chart-fields ""` (CSV, Excel, HTML, native PPT) |
| Stage 3 / 4 (*Run Pipeline* with JMP) | `--chart-fields "<params>" --generate-jmp-charts --jmp-exe <jmp.exe>` (+ `--ppt-template` for stage 4) |
| Source = *Existing CSV* / *Excel* / *JMP PNG folder* | `--jmp-from-csv` / `--jmp-from-excel` / `--ppt-from-charts` |
| Tab 2 *Apply axis settings* | `--jmp-axis-config <json>` (the GUI writes the grid to `jmp_axis_settings.json`) |
| Tab 2 *Spec RMT limits as ±Ref (per frequency)* unticked | `--no-rmt-limits` |
| Tab 2 *Recalculate from data* | Not needed on the CLI: DTR charts size their axis from the data, and WCL widens the axis to show the limit lines. For non-thermal charts, give an axis JSON that fits the data, or watch for `[warn] JMP axis hides data`. |
| Tab 3 *JSL only* | `--jmp-jsl-only` |
| *Process each file SEPARATELY* | One command per log with its own `--outdir` ([3.6](#36-one-output-folder-per-log-process-each-file-separately)) |
| The GUI always adds | `--no-ask-chart-approval` |

The GUI command preview (Tab 3) shows the exact command it will run. Copy it into your
automation as a starting point.

---

## 5. Flag reference

Run `marginiq.cmd --help` for the built-in help.

**Input / output**

| Flag | Default | Meaning |
|------|---------|---------|
| `--project KEY` | `NVL` (registry default) | Project key, code or name from `projects.json` (`NVL`, `WCL`, `"Wildcat Lake"`). Selects the parsing rules ([section 6](#6-what---project-changes-nvl-vs-wcl)). Unknown → exit 2 |
| `--input PATH [PATH …]` | – | Log files and/or folders (folders are searched recursively). Required unless `--jmp-from-csv`, `--jmp-from-excel` or `--ppt-from-charts` is used |
| `--pattern GLOB` | `*.txt` | File mask for folders. Several masks: `"*.log;*.txt"` |
| `--dtr` / `--no-dtr` | off | Thermal Experiment (BCRH / BHRC): Boot-temperature + Run-temperature RMT per log |
| `--outdir DIR` | **required** | Output folder (created if missing) |
| `--excel-name NAME` | `RMT_Extraction.xlsx` | Excel file name |
| `--ppt-name NAME` | `RMT_Summary.pptx` | Native PPT file name. The JMP deck is `<stem>_JMP_Charts.pptx` |

**Charts**

| Flag | Default | Meaning |
|------|---------|---------|
| `--chart-fields "A,B"` | all 8 | Parameters to chart: `RecEnDelay, TxDqsDelay, RxDqsDelay, TxDqDelay, RxDqVrefByte, TxVref, ClkGrpPi, CmdVref`. `""` = all (no JMP run without `--generate-jmp-charts`) |
| `--no-ask-chart-approval` | asks | **Always pass in automation.** Without it the tool asks for approval on a console (deny → exit 3); with no console it prints a warning and **discards your `--chart-fields` selection** |
| `--generate-jmp-charts` | off | Run JMP to create `jmp_charts\<Param>.png`, the JMP PPT and the HTML JMP tab. Needs `--jmp-exe` |
| `--jmp-exe PATH` | – | `jmp.exe`, e.g. `C:\Program Files\SAS\JMPPRO\17\jmp.exe` |
| `--jmp-jsl-only` | off | Write `rmt_jmp_charts.jsl` only; do not start JMP |
| `--jmp-axis-config JSON` | none | Axis Min / Max / Inc / ±Ref per parameter (see `jmp_axis_settings.json`) |
| `--rmt-limits` / `--no-rmt-limits` | on | Projects with `rmt_limits` (WCL): spec limit for the log's memory type + frequency as ±Ref, and the Spec RMT Limits PASS/FAIL table. No effect for NVL |
| `--ppt-template PPTX` | none | Template for the JMP-chart PPT |

**Reuse / misc**

| Flag | Meaning |
|------|---------|
| `--jmp-from-csv CSV [CSV …]` | Skip log parsing; Excel + HTML + JMP charts + JMP PPT from Extended CSV(s). Needs `--jmp-exe` |
| `--jmp-from-excel XLSX` | Same from the `All_RMT` sheet of a workbook. Needs `--jmp-exe` |
| `--ppt-from-charts DIR` | Only build the JMP-chart PPT from existing PNGs |
| `--no-update-check` | Skip the once-a-day "newer version on the server?" check (recommended in CI) |
| `--version` | Print `MarginIQ <version>` and exit 0 |

---

## 6. What `--project` changes (NVL vs WCL)

All rules live in [`projects.json`](projects.json); no code change is needed for a new platform.

| Rule | NVL (NVL-S / NVL-HX, DDR5) | WCL (Wildcat Lake) |
|------|----------------------------|--------------------|
| RMT block | `START_RMT` per rank (`Mc0.C0.R0`) | `START_RMT` (DDR5 RVP) or `START_BYTE_RMT` per byte-rank (`Mc0.C0.B0.R0`, LPDDR5 RVP); header `RxVref` = `RxDqVrefByte` |
| Which blocks count | Inside the *Rank Margin Tool* task. Logs without task markers (consolidated snippets) are read anywhere | **Only** inside *Rank Margin Tool -- Started … SUCCEEDED*. Per-training-step "Composite UI" blocks are never plotted. A log without that data prints `[warn] … NO RMT DATA …` and is skipped |
| Platform variant | – | `WCL DDR5 RVP` / `WCL LPDDR5 RVP`, decoded from the log (HTML header + Platform tab, PPT) |
| `--dtr` flow | *Temp Drift RMT Test* (boot) + *Rank Margin Tool* (run), temps from `PHY Temperature1` | First / last *Rank Margin Tool* run (`x 0xdd44` re-runs), temps from **BCRH** (0 → 90 °C) / **BHRC** (90 → 0 °C) in the file name |
| ±Ref lines | Axis JSON (default ±10) | **Spec RMT limits** for the log's memory type + frequency (PTL/WCL limits table). `--no-rmt-limits` restores the axis JSON ±Ref |

WCL spec limits, in ticks (L = H). Mail columns: RdT, WrT, RdV, WrV, CmdT, CmdV. The full
table, with ps / mV values, is in `projects.json`:

| Memory | MT/s | RxDqsDelay | TxDqDelay | RxDqVrefByte | TxVref | ClkGrpPi | CmdVref |
|--------|------|-----------:|----------:|-------------:|-------:|---------:|--------:|
| LPDDR5 | 2400 – 8533 | 14 | 14 | 45 | 25 | 40 | 40 |
| LPDDR5 | 9600 | 13 | 13 | 45 | 25 | 40 | 40 |
| DDR5 | 3200 | 14 | 14 | 32 | 20 | 40 | 18 |
| DDR5 | 4800 | 14 | 14 | 32 | 20 | 20 | 18 |
| DDR5 | 5600 | 14 | 14 | 32 | 20 | 20 | 16 |
| DDR5 | 6000 | 14 | 14 | 32 | 20 | 20 | 15 |
| DDR5 | 6400 | 14 | 14 | 32 | 20 | 20 | 14 |
| DDR5 | 6800 | 13 | 13 | 28 | 18 | 20 | 14 |
| DDR5 | 7200 | 13 | 13 | 25 | 16 | 20 | 13 |

A frequency that is not listed uses the nearest listed one and prints a `[warn]`.

---

## 7. Exit codes

The codes are stable; automation can rely on them.

| Code | Meaning | Typical cause / action |
|-----:|---------|------------------------|
| **0** | Success | Outputs written. **A spec FAIL is still exit 0**: read `rmt_metadata.json` → `rmt_limits.overall` ([8.2](#82-rmt_metadatajson)) |
| **1** | No input files | `--input` / `--pattern` matched nothing |
| **2** | Error | Bad / unknown arguments, unknown `--project`, missing CSV / Excel / charts folder, **no RMT data in any input file** (e.g. WCL log whose Rank Margin Tool never ran), `.venv` missing (`marginiq.cmd`) |
| **3** | Chart approval denied | Only with interactive approval; pass `--no-ask-chart-approval` |
| **4** | JMP needed but `--jmp-exe` not given | Add `--jmp-exe` |
| **5** | JMP failed | `jmp.exe` not found at the given path, could not start or crashed; check JMP / licence |

When some files have RMT data and others do not, the run succeeds (exit 0). The skipped
files print `[warn] <file>: NO RMT DATA - …`.

---

## 8. Output contract

### 8.1 Files in `--outdir`

| File | When | Content |
|------|------|---------|
| `RMT_Combined_Extended.csv` | log parsing, `--jmp-from-csv`, `--jmp-from-excel` | One row per rank (or byte-rank): `SourceFile, Frequency, Gear, BlockIndex, Params, BootTemp, RunTemp, <Param>-, <Param>+ …` |
| `RMT_Combined_Similar.csv` | log parsing | Same without the first four columns |
| `csv_by_file\<log>_RMT.csv` | log parsing | One CSV per log |
| `RMT_Extraction.xlsx` | log parsing, `--jmp-from-csv` | `All_RMT`, `File_<log>`, `Freq_<MT/s>` sheets |
| `RMT_Report.html` | log parsing, `--jmp-from-csv`, `--jmp-from-excel` | Interactive report (Overview with the WCL Spec RMT Limits card, Training Steps, Platform, Mode Registers, ODT, JMP Charts, Raw Data) |
| `RMT_Summary.pptx` | log parsing | Native summary deck (incl. Spec RMT Limits slide for WCL) |
| `rmt_metadata.json` | log parsing (updated by `--jmp-from-*`) | Machine-readable run data ([8.2](#82-rmt_metadatajson)) |
| `rmt_jmp_charts.jsl` | JMP stages | JMP script |
| `jmp_charts\<Param>.png`, `RMT_Summary_JMP_Charts.pptx` | JMP ran | Charts and the chart deck |

Values are RMT ticks: timing = UI/64 PI ticks, voltage = Vref steps. `<Param>-` is negative.

### 8.2 `rmt_metadata.json`

```jsonc
{
  "tool": "MarginIQ 2.0.0",
  "project": "WCL",
  "training_steps": [            // MRC task timeline per log
    {"task_type": "GREEN MRC", "name": "Read Voltage Centering", "status": "SUCCEEDED",
     "status_code": "0h", "start_line": 10427, "end_line": 10842, "iteration": 0,
     "source": "teraterm_17092026.log", "post_code": "DD57h"},   // POST code printed just before "Started" (null if none)
    {"task_type": "GREEN MRC", "name": "DDRIO Config", "status": "INCOMPLETE",  // SUCCEEDED|FAILED|RUNNING|INCOMPLETE
     "status_code": null, "start_line": 1648, "end_line": null, "iteration": 0,
     "source": "teraterm_17092026.log", "post_code": "DD12h",
     "note": "end line missing - log gap between lines 1648 and 2857"}
  ],
  "platform_infos": [            // one per log
    {"source_file": "...", "board": "WCL LP5 RVP2", "platform_variant": "WCL LPDDR5 RVP",
     "phy_version": "...", "dimms": [{"slot": "0/0/0", "module_type": "LPDDR5X Memory Down", "...": "..."}],
     "mr_odt": {"...": "..."}}
  ],
  "rmt_limits": {                // only when spec limits applied (WCL)
    "overall": "FAIL",           // PASS | FAIL  <- gate your automation on this
    "combos": [{"family": "LPDDR5", "table_freq": 7467, "log_freqs": [7467], "exact": true,
                "files": ["teraterm_17092026.log"], "ticks": {"...": [14, 14]}, "physical": {"...": [29, 29]}}],
    "chart":  {"TxVref": [25.0, 25.0], "...": "..."},        // ±Ref used in the charts
    "check": [{"param": "CmdVref", "limit": [40.0, 40.0], "worst_minus": -50.0, "worst_plus": 10.0,
               "n": 16, "result": "FAIL", "fails": ["Mc0.C3.B0.R1 (-58/10)", "Mc0.C3.B1.R1 (-58/10)"]}]
  }
}
```

---

## 9. Console messages to scrape

Output is plain text, one message per line. With `marginiq.cmd` it is UTF-8 and can be
redirected safely.

| Pattern | Meaning |
|---------|---------|
| `MarginIQ 2.0.0 \| Project: Wildcat Lake (WCL)` | First line: version and project |
| `Done. Parsed <N> RMT rows from <M> files.` | Log-parsing run finished |
| `Output folder: <dir>` | Where the outputs are |
| `[warn] <file>: NO RMT DATA - …; file skipped.` | That log has no RMT data inside the Rank Margin Tool task (WCL) |
| `No START_RMT blocks found in input files.` | No file had data → exit 2 |
| `[info] RMT limits (<family> <MT/s> MT/s, ticks): …` | Spec limits picked for the run (WCL) |
| `[warn] RMT spec FAIL <Param> (limit +/-<L>): <n>/<total> rank(s) - …` | Spec failure (exit code stays 0) |
| `[warn] RMT limits: … not in the limit table; using the nearest …` | Frequency not listed |
| `[warn] … log capture gap; marked INCOMPLETE.` | Serial capture lost lines (training steps only; margins unaffected) |
| `[warn] JMP axis hides data for …` | Axis JSON range clips points; widen it or recalculate |
| `[info] DTR: parsed <N> rows from <file> (Freq=…, Gear=…, BootTemp=…, temps from …)` | Thermal log parsed |
| `[MarginIQ update] A newer version is available …` | Run `git pull` in the repository |

---

## 10. Environment variables

| Variable | Effect |
|----------|--------|
| `MARGINIQ_NO_UPDATE_CHECK=1` | Never contact the git server (same as `--no-update-check` on every run) |
| `MARGINIQ_UPDATE_CHECK_HOURS=<h>` | Update-check interval (default 24) |
| `PYTHONUTF8=1`, `PYTHONIOENCODING=utf-8` | Set by `marginiq.cmd`. Set them yourself when calling `python.exe` directly with redirected output |
| `RMT_PPT_TEMPLATE=<pptx>` | Default template for the interactive runner (`rmt_pipeline_runner.py`) only |

---

## 11. Integration examples

### 11.1 Batch file with error handling

```bat
@echo off
setlocal
set "MIQ=C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd"
set "OUT=D:\results\WCL_nightly"

call "%MIQ%" --project WCL --input "D:\lab\WCL\today" --pattern "*.log" --outdir "%OUT%" ^
     --no-ask-chart-approval --no-update-check > "%OUT%.log" 2>&1
set RC=%ERRORLEVEL%
if %RC%==0 (echo MarginIQ OK: %OUT%) else (echo MarginIQ failed with exit %RC% - see %OUT%.log & exit /b %RC%)
findstr /C:"RMT spec FAIL" "%OUT%.log" >nul && echo WARNING: spec limit failures - see %OUT%\RMT_Report.html
```

### 11.2 PowerShell

```powershell
$miq = 'C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd'
$out = 'D:\results\WCL_run7'
& $miq --project WCL --input 'D:\lab\WCL\run7' --pattern '*.log' --outdir $out `
       --no-ask-chart-approval --no-update-check 2>&1 | Tee-Object "$out.log"
if ($LASTEXITCODE -ne 0) { throw "MarginIQ exit $LASTEXITCODE" }
$meta = Get-Content "$out\rmt_metadata.json" -Raw | ConvertFrom-Json
if ($meta.rmt_limits -and $meta.rmt_limits.overall -eq 'FAIL') {
    $meta.rmt_limits.check | Where-Object result -eq 'FAIL' |
        ForEach-Object { "{0}: {1}" -f $_.param, ($_.fails -join ', ') }
}
```

### 11.3 Python (any automation framework)

```python
import json, subprocess
from pathlib import Path

MIQ = r"C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd"
out = Path(r"D:\results\WCL_run7")
proc = subprocess.run(
    [MIQ, "--project", "WCL", "--input", r"D:\lab\WCL\run7", "--pattern", "*.log",
     "--outdir", str(out), "--no-ask-chart-approval", "--no-update-check"],
    capture_output=True, text=True, encoding="utf-8", errors="replace")
if proc.returncode != 0:
    raise RuntimeError(f"MarginIQ exit {proc.returncode}:\n{proc.stdout[-2000:]}")
meta = json.loads((out / "rmt_metadata.json").read_text(encoding="utf-8"))
limits = meta.get("rmt_limits")                      # None for NVL / --no-rmt-limits
verdict = limits["overall"] if limits else "N/A"
fails = [c for c in (limits or {}).get("check", []) if c["result"] == "FAIL"]
print(verdict, [(c["param"], c["fails"]) for c in fails])
```

`marginiq.cmd` is a batch file. `subprocess.run([...])` starts it through `cmd.exe`
automatically, and no `shell=True` is needed.

### 11.4 Windows Task Scheduler

*Program/script*: `C:\tools\RMT-Margin-Analysis-Tool\marginiq.cmd`
*Add arguments*: `--project WCL --input "D:\lab\WCL\nightly" --pattern "*.log" --outdir "D:\results\nightly" --no-ask-chart-approval --no-update-check`
*Start in*: any folder. The task result is the MarginIQ exit code.

---

## 12. Rules for unattended runs

1. Always pass `--no-ask-chart-approval`; add `--no-update-check` (or set
   `MARGINIQ_NO_UPDATE_CHECK=1`) on build agents.
2. Give every run its **own `--outdir`**. Files in an existing folder are overwritten.
3. **JMP is single-instance.** Before every chart run MarginIQ **closes any running JMP**,
   including an interactive session on the same machine. Run JMP-enabled jobs one at a time,
   and never on a machine where someone is using JMP. Parallel runs without JMP are fine.
4. Each JMP run waits up to 2 minutes for JMP. Allow about 5 minutes per JMP job in your
   scheduler's timeout.
5. Always pass `--project`. Without it the registry default (`NVL`) is used, and WCL logs
   would then use the NVL rules.
6. For thermal logs pass `--dtr`. Keep the `BCRH` / `BHRC` tag in WCL file names: the
   temperatures come from it.
7. Use one `--outdir` for Stage 1 and `--jmp-from-csv` (two-step runs) so
   `rmt_metadata.json` is found. WCL needs it to pick the spec limits.
8. Gate on the **exit code** first, then on `rmt_metadata.json` → `rmt_limits.overall` for
   spec pass/fail (WCL).
9. Do not start `rmt_gui.py`, `Launch_RMT_GUI.bat`, `MarginIQ.lnk` or
   `rmt_pipeline_runner.py` from automation: they wait for a user.

---

## 13. Interactive command-line menu (people, not automation)

```bat
.venv\Scripts\python.exe rmt_pipeline_runner.py          :: full menu, stages 1-7
.venv\Scripts\python.exe rmt_pipeline_runner.py --quick  :: minimal prompts, stages 1-4
.venv\Scripts\python.exe rmt_pipeline_runner.py --cmds   :: print ready-made cmd.exe commands
.venv\Scripts\python.exe rmt_pipeline_runner.py --guide cli  :: open a guide (menu option 8)
```

The menus ask for the **stage**, then the **Project** (NVL / WCL / …), the
**Experiment profile** (Non-Thermal or Thermal DTR → `--dtr`), whether to use the
**WCL spec limits**, then input and output paths. They run the same pipeline as
`marginiq.cmd`. See [RMT_QUICKSTART.md](RMT_QUICKSTART.md) for a walkthrough.
