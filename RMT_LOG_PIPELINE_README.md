# MarginIQ RMT Log Pipeline - Complete Reference Manual

*Intel CCG CVE DDR5 RMT Margin Analysis Tool*

📖 **Quick Links:** [Main README](README.md) · [Quick Start Guide](RMT_QUICKSTART.md) · [Agentic mode](README.md#-agentic-mode-github-copilot) · [To Top](#marginiq-rmt-log-pipeline---complete-reference-manual)

---

## Overview

`rmt_log_pipeline.py` is the MarginIQ engine used by the GUI, the agent skill and scripts. It
extracts `START_RMT` data from one or more MRC `.txt` / `.log` files of any project in
[`projects.json`](projects.json) and generates:

1. CSV files in a format similar to your reference files.
2. An Excel workbook with individual sheets.
3. A self-contained HTML report, including the **Mode Registers** and **ODT** tabs and the JMP
   chart comparison workspace.
4. A PowerPoint summary with metric and chart slides.
5. Optionally, JMP Graph Builder charts and a PPT built from them.

The same engine runs in three modes: the **GUI** (MarginIQ desktop icon / `Launch_RMT_GUI.bat`),
**agentic** (GitHub Copilot Agent mode, driven by
[`.github/skills/rmt-margin-plots/SKILL.md`](.github/skills/rmt-margin-plots/SKILL.md); see
[README › Agentic mode](README.md#-agentic-mode-github-copilot)), and the **CLI** described here.

Script path:
- `rmt_log_pipeline.py` (repository root)

> All commands below use `.venv/Scripts/python.exe`, the interpreter created by [setup.bat](setup.bat). Run them from the repository root (the folder that contains `setup.bat`).

---

## Input Log Pattern

The parser expects sections similar to:

- `Setting boot frequency to <value>`
- `Setting gear ratio to <value>`
- `START_RMT`
- `Params: RecEnDelay TxDqsDelay ...`
- Rows such as `Mc0.C0.R0: -32.0 32.0 ...` (or per-byte `Mc0.C0.B0.R0: ...`, e.g. Wildcat Lake)

Each row holds one minus/plus pair per column of the `Params:` header (16 values for the 8
standard parameters). Values are mapped by **header name**, so a different column order is
handled. Project aliases from `projects.json` (e.g. WCL `RxVref` → `RxDqVrefByte`) are applied
when `--project` is given.

Additionally parsed for the HTML report and `rmt_metadata.json`:

| Log section | Used for |
|-------------|----------|
| `MCx.Cy.Rz  Data  Delay (nCK)` table (SAGV Finalization) | Final per-rank Mode Register values |
| `InitMrwDdr5:` MR writes (JEDEC RESET) | JEDEC-reset MR values ("changed by training" highlight) |
| `DIMM ODT summary:` + `CPU Summary: ... Read ODT` (MRC Data Summary) | RTT_WR / NomWr / NomRd / Park / ParkDqs / CA / CS / Ron and CPU read ODT |
| `Mc0.Ch0.R0: Ddr5OdtlOnWr:...` (DDR5 ODT Timing Config) | ODT latency offsets |
| `DimmN ODT Values:`, `RcompTarget[RdOdt]`, `DDRIO ODT Mode` | BIOS / CPU ODT inputs |

---

## Dependencies

All Python dependencies are installed into the tool-local `.venv` by [setup.bat](setup.bat); there is no manual pip step.

Required:
- Python 3.10+ (3.14 recommended), installed with tcl/tk so the GUI works
- `openpyxl`
- `python-pptx` (PPT generation)

Optional:
- `pillow` (GUI screenshot capture only)
- SAS JMP Pro — a licensed Windows application, never installed by pip. `setup.bat` only detects it.

Install / repair the environment:

```powershell
.\setup.bat            # create .venv and install requirements.txt
.\setup.bat --recreate # wipe .venv and rebuild from scratch
```

---

## Usage

### 🚀 User-Friendly Mode (Recommended)

For simpler, interactive commands use the wrapper script:

```powershell
python "rmt_pipeline_runner.py"
```

Or in quick mode with minimal prompts:

```powershell
python "rmt_pipeline_runner.py" --quick
```

Or in quick mode with minimal prompts:

```powershell
python "rmt_pipeline_runner.py" --quick
```

The wrapper will guide you through:
1. Selecting a workflow stage (1-4)
2. Entering your input folder path
3. Entering your output folder path
4. Choosing chart fields (for stages 3 and 4)

**What each stage does:**
- **Stage 1** 📄 CSV Only → Fast data extraction only
- **Stage 2** 📊 CSV + Excel → Best for data analysis and pivots
- **Stage 3** 🎨 CSV + Excel + JMP Charts → JMP scatter charts per parameter
- **Stage 4** 📑 CSV + Excel + JMP Charts + PPT → Full report with PowerPoint **← default**

---

### Output Matrix: Which Outputs Are Generated at Each Stage?

| Stage | CSV | Excel | JMP Charts | PPT | Default | Use Case |
|-------|-----|-------|-----------|-----|---------|----------|
| 1 | ✓ | | | | | Quick data extraction |
| 2 | ✓ | ✓ | | | | Data analysis & pivot tables |
| 3 | ✓ | ✓ | ✓ | | | JMP scatter charts per parameter |
| 4 | ✓ | ✓ | ✓ | ✓ | ✅ | Full report with PPT from JMP PNGs |

---

### Quick Reference: One-Shot Complete Pipeline (Stage 4 — default)

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --ppt-template "C:/path/to/template.pptx" `
  --jmp-axis-config "jmp_axis_settings.json"
```

---

## Staged Workflows

### Stage 1: CSV Only

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval "--chart-fields="
```

**Outputs:**
- `RMT_Combined_Similar.csv`
- `RMT_Combined_Extended.csv`
- `csv_by_file/*.csv`

---

### Stage 2: CSV + Excel

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval "--chart-fields="
```

**Outputs:**
- All Stage 1 outputs
- `RMT_Extraction.xlsx` (All_RMT, per-file, per-frequency sheets)

---

### Stage 3: CSV + Excel + JMP Charts

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --jmp-axis-config "jmp_axis_settings.json"
```

**Outputs:**
- All Stage 2 outputs
- `rmt_jmp_charts.jsl`
- `jmp_charts/<Param>.png` (one per parameter, X=Params Y=±values Group=Frequency, or Boot/Run temperature for DTR)

---

### Stage 4: CSV + Excel + JMP Charts + PPT  (default)

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --ppt-template "C:/path/to/template.pptx" `
  --jmp-axis-config "jmp_axis_settings.json"
```

**Outputs:**
- All Stage 3 outputs
- `RMT_Summary_JMP_Charts.pptx` (JMP PNGs assembled using template)

---

## Input Variations

### Parse all txt files from a folder

```powershell
& ".venv/Scripts/python.exe" "rmt_log_pipeline.py" --input "C:/path/to/logs" --pattern "*.txt" --outdir "C:/path/to/output" --no-ask-chart-approval "--chart-fields="
```

### Parse a single txt file

```powershell
& ".venv/Scripts/python.exe" "rmt_log_pipeline.py" --input "C:/path/to/logs/6400_G2.txt" --outdir "C:/path/to/output" --no-ask-chart-approval "--chart-fields="
```

### Parse multiple folders/files

```powershell
& ".venv/Scripts/python.exe" "rmt_log_pipeline.py" --input "C:/path/to/logs" "C:/path/to/other_logs" --pattern "*.txt" --outdir "C:/path/to/output" --no-ask-chart-approval "--chart-fields="
```

---

## CLI Switches Quick Reference

| Switch | Default | Purpose |
|--------|---------|---------|
| `--project` | `NVL` (registry default) | Project from `projects.json` (key, code or name, e.g. `NVL`, `WCL`, `"Wildcat Lake"`); labels reports and selects START_RMT header aliases. Unknown value → exit 2 |
| `--dtr` | False | Parse Thermal Experiment (DTR, BCRH/BHRC) logs: Boot-temp + Run-temp RMT per log |
| `--version` | - | Print the MarginIQ version |
| `--input` | *required* | Path(s) to input folder(s) or file(s) |
| `--pattern` | `*.txt` | File glob pattern (when input is folder) |
| `--outdir` | *required* | Output directory for all generated files |
| `--excel-name` | `RMT_Extraction.xlsx` | Excel workbook filename |
| `--ppt-name` | `RMT_Summary.pptx` | PPT presentation filename |
| `--chart-fields` | All params | Comma-separated fields for PPT/JMP charts |
| `--ask-chart-approval` | True | Prompt user before PPT chart generation |
| `--no-ask-chart-approval` | - | Skip approval prompt (automation mode) |
| `--generate-jmp-charts` | False | Enable JMP chart generation |
| `--jmp-exe` | None | Path to jmp.exe executable |
| `--jmp-jsl-only` | False | Generate JSL script without launching JMP |

---

## Output Files

In `--outdir`, the script creates:

1. `RMT_Combined_Similar.csv`
- Core columns to match your sample CSV style.

2. `RMT_Combined_Extended.csv`
- Core columns + metadata (`SourceFile`, `Frequency`, `Gear`, `BlockIndex`).

3. `csv_by_file/<file_stem>_RMT.csv`
- Per-log CSV outputs in similar format.

4. `RMT_Extraction.xlsx`
- `All_RMT` sheet.
- One sheet per source file.
- One sheet per frequency (`Freq_5600`, `Freq_6400`, ...).

5. `RMT_Report.html`
- Self-contained interactive report: Overview, Frequency, Parameters, RunTemp, Training Steps,
  Platform, **Mode Registers**, **ODT**, **JMP Charts** (dropdown comparison panels) and **Raw Data**
  (frozen-column explorer with PASS/WARN/FAIL status, heatmap, numeric filters `<15` / `10..20`,
  Margins / Window width / Slack-vs-Ref views, Copy / CSV export, and an A → B **Δ Drift
  Comparison**, e.g. Boot → Run RMT, BCRH vs BHRC or 4800 vs 5200).
  The report needs no jQuery / DataTables; only Bootstrap and Chart.js are loaded from the CDN.

6. `rmt_metadata.json`
- Tool version, project, training steps, platform info and MR / ODT data
  (`platform_infos[*].mr_odt`). Reused by later `--jmp-from-csv` runs in the same `--outdir`.

7. `RMT_Summary.pptx`
- Title/coverage slide (MarginIQ branding, project, tool version).
- Average window width text summary.
- Overall chart (approved fields only).
- Frequency comparison chart (approved fields only).
- RunTemp trend charts by rank (approved fields only).
- Per-frequency summary slides.

Before PPT chart generation, the script asks for chart field approval in terminal (default behavior).

8. Optional JMP artifacts
- `rmt_jmp_charts.jsl` (auto-generated JSL script)
- `jmp_charts/<Param>.png` (one PNG per parameter: X=Params, Y=param+/param-, Group X=Frequency)
- `rmt_graph_builder.jsl` (GUI Tab 2 *JMP Graph Builder*: interactive windows, no PNG export / Quit)
- `<ppt-name stem>_JMP_Charts.pptx` (PPT assembling all JMP PNGs using the `--ppt-template` styling)

---

## Column Mapping (Similar CSV)

The similar CSV columns are:

- `Params`
- `BootTemp`, `RunTemp` (if found in logs)
- `RecEnDelay-`, `RecEnDelay+`
- `TxDqsDelay-`, `TxDqsDelay+`
- `RxDqsDelay-`, `RxDqsDelay+`
- `TxDqDelay-`, `TxDqDelay+`
- `RxDqVrefByte-`, `RxDqVrefByte+`
- `TxVref-`, `TxVref+`
- `ClkGrpPi-`, `ClkGrpPi+`
- `CmdVref-`, `CmdVref+`

---

## Notes and Limits

1. If `BootTemp` and `RunTemp` are not present in logs, those fields remain blank.
2. If a row has fewer value pairs than its `Params:` header, it is skipped.
3. If `python-pptx` is missing, CSV and Excel are still generated.
4. By default, chart field approval is interactive. For automation, use `--no-ask-chart-approval` and optionally `--chart-fields`.

---

## Chart Field Approval (New)

By default, script behavior is:

1. Show available chart fields.
2. Ask user to select fields (`all`, `none`, names, or numeric indexes).
3. Ask explicit approval (`y/n`).
4. Generate chart slides only for approved fields.
5. If approval is denied (`n`), script exits with code `3` and does not generate PPT.

Automation examples:

```powershell
# Non-interactive: generate charts for selected fields
& ".venv/Scripts/python.exe" "rmt_log_pipeline.py" --input "C:/path/to/logs" --pattern "*.txt" --outdir "C:/path/to/output" --no-ask-chart-approval --chart-fields "RxDqVrefByte,TxVref,ClkGrpPi"

# Non-interactive: no chart slides
& ".venv/Scripts/python.exe" "rmt_log_pipeline.py" --input "C:/path/to/logs" --pattern "*.txt" --outdir "C:/path/to/output" --no-ask-chart-approval "--chart-fields="
```

---

## JMP Axis Configuration

The file `jmp_axis_settings.json` (located in the same directory as the scripts) controls
the Y-axis scale, tick increment, and reference lines for every JMP chart.

### JSON Structure

```json
{
  "defaults": {
    "inc": 5,
    "minor_ticks": 0,
    "show_major_grid": 1,
    "show_minor_grid": 1,
    "ref_line_plus_color": "Medium Light Red",
    "ref_line_minus_color": "Blue",
    "ref_line_plus": 10,
    "ref_line_minus": -10
  },
  "parameters": {
    "TxVref": {
      "plus":  { "min": 5,  "max": 35, "inc": 2, "ref_line": 10  },
      "minus": { "min": -35, "max": -5, "inc": 2, "ref_line": -10 }
    }
  }
}
```

### Fields

| Field | Description |
|-------|-------------|
| `min` | Y-axis minimum value |
| `max` | Y-axis maximum value |
| `inc` | Major tick increment |
| `minor_ticks` | Number of minor ticks between major ticks (0 = off) |
| `ref_line` | Value at which a reference line is drawn |
| `ref_line_plus_color` | Color for positive-panel reference line (default: `"Medium Light Red"`) |
| `ref_line_minus_color` | Color for negative-panel reference line (default: `"Blue"`) |
| `ref_line_plus` / `ref_line_minus` | **Default ±Ref** (defaults block only, default `10` / `-10`). Used for any parameter without its own `ref_line`: JMP reference lines (including DTR temperature charts), PASS/WARN/FAIL status and the red raw-data cells in the HTML report. GUI: Tab 2 ▸ Defaults ▸ + Ref / − Ref. |
| `show_major_grid` | 1 = show major grid lines, 0 = hide |
| `show_minor_grid` | 1 = show minor grid lines, 0 = hide |

`defaults` values are used for any field not overridden in the `parameters` block.

### Passing the Config File

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" --outdir "C:/path/to/output" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --jmp-axis-config "jmp_axis_settings.json"
```

When using `rmt_pipeline_runner.py`, the config file is detected and passed automatically.

---

## JMP Chart Format

Each generated JMP chart uses Graph Builder with:
- **X axis**: Params (rank labels, nominal/categorical)
- **Y top panel**: `param+` values with configurable scale and reference line
- **Y bottom panel**: `param-` values with configurable scale and reference line
- **Group X**: Frequency (side-by-side facets per frequency)
- **Chart type**: Points (scatter)

Axis scale, tick increment, and reference lines are controlled by `jmp_axis_settings.json`.

---

## JMP CLI Chart Generation

### Run with JMP charts (all defaults)

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RxDqVrefByte,TxVref,ClkGrpPi" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe"
```

### Run with PPT template + axis config (recommended)

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --ppt-template "C:/path/to/template.pptx" `
  --jmp-axis-config "jmp_axis_settings.json"
```

### Generate JSL only (test mode, do not launch JMP)

```powershell
& ".venv/Scripts/python.exe" `
  "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" --outdir "C:/path/to/output" `
  --generate-jmp-charts --jmp-exe "...\jmp.exe" --jmp-jsl-only
```

Notes:
1. `--generate-jmp-charts` requires `--jmp-exe`.
2. JMP output images are saved under `jmp_charts/` inside your output directory.
3. When `--ppt-template` is provided, a second PPT (`*_JMP_Charts.pptx`) is built from the PNGs.
4. `--jmp-axis-config` is optional; omit it to use hardcoded defaults.

---

## Advanced Usage: Direct Script Commands

For power users and automation, use the main script directly with full CLI control. These examples use PowerShell line continuation (backtick) for readability:

**Stage 1: CSV Only**
```powershell
python "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  "--chart-fields="
```

**Stage 4: Full Pipeline with JMP Charts + PPT (default)**
```powershell
python "rmt_log_pipeline.py" `
  --input "C:/path/to/logs" `
  --pattern "*.txt" `
  --outdir "C:/path/to/output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:/Program Files/SAS/JMPPRO/17/jmp.exe" `
  --ppt-template "C:/path/to/template.pptx" `
  --jmp-axis-config "jmp_axis_settings.json"
```

---

## CLI Switches Quick Reference

| Switch | Default | Purpose |
|--------|---------|---------|
| `--project` | `NVL` (registry default) | Project from `projects.json` (key, code or name, e.g. `NVL`, `WCL`, `"Wildcat Lake"`); labels reports and selects START_RMT header aliases. Unknown value → exit 2 |
| `--dtr` | False | Parse Thermal Experiment (DTR, BCRH/BHRC) logs: Boot-temp + Run-temp RMT per log |
| `--version` | - | Print the MarginIQ version |
| `--input` | *required* | Path(s) to input folder(s) or file(s) |
| `--pattern` | `*.txt` | File glob pattern (when input is folder) |
| `--outdir` | *required* | Output directory for all generated files |
| `--excel-name` | `RMT_Extraction.xlsx` | Excel workbook filename |
| `--ppt-name` | `RMT_Summary.pptx` | PPT presentation filename |
| `--chart-fields` | All params | Comma-separated fields for PPT/JMP charts |
| `--ask-chart-approval` | True | Prompt user before PPT chart generation |
| `--no-ask-chart-approval` | - | Skip approval prompt (automation mode) |
| `--generate-jmp-charts` | False | Enable JMP chart generation |
| `--jmp-exe` | None | Path to jmp.exe executable |
| `--jmp-jsl-only` | False | Generate JSL script without launching JMP |
| `--ppt-template` | None | Path to .pptx template used for the JMP-charts PPT |
| `--jmp-axis-config` | None | Path to `jmp_axis_settings.json` for chart axis customisation. `defaults.ref_line_plus` / `ref_line_minus` (default ±10) apply to any parameter without its own `ref_line` |

---

## Exit Codes & Status

The script returns exit codes to indicate success or failure:

| Code | Status | Meaning | Action |
|------|--------|---------|--------|
| 0 | ✅ Success | Extraction completed | Check output folder |
| 1 | ⚠️ Warning | No input files matched `--input` / `--pattern` | Check the path and pattern |
| 2 | ❌ Error | Invalid arguments, missing files, no START_RMT blocks, or unknown `--project` | Review --help, paths and `projects.json` |
| 3 | ❌ Denied | Chart approval was rejected | Re-run with --no-ask-chart-approval or approve charts |
| 4 | ❌ Config | JMP executable not found | Verify --jmp-exe path |
| 5 | ❌ Runtime | JMP execution failed | Check JSL script or JMP installation |

---

## Troubleshooting

### No rows extracted

Check that logs contain `START_RMT` and lines beginning with rank IDs (`Mc0.C0.R0:`).

```powershell
# Verify logs contain expected patterns
Select-String -Path "c:\logs\*.txt" -Pattern "START_RMT" | Select-Object -First 5
```

### PPT not generated

Rebuild the environment so `python-pptx` is present:

```powershell
.\setup.bat --recreate
```

### Permission errors

Run terminal with access rights to input and output directories.

```powershell
Start-Process PowerShell -Verb RunAs
```

### Charts not appearing in PPT

Various reasons charts may not appear:

1. **No data in logs** - Verify log contains START_RMT and rank data
2. **Approval denied** - Script exits before generating PPT if you deny approval (exit code 3)
3. **Missing python-pptx** - Re-run `.\setup.bat --recreate` from the tool folder
4. **Empty chart fields** - Use `--chart-fields "RxDqVrefByte,TxVref"` or similar

### JMP script generated but no charts

Manually run the JSL script to see detailed errors:

```powershell
& "C:\Program Files\SAS\JMPPRO\17\jmp.exe" "C:\path\to\output\rmt_jmp_charts.jsl"
```

Or inspect the JSL file for syntax:

```powershell
Get-Content "C:\path\to\output\rmt_jmp_charts.jsl" | Select-Object -First 20
```

### Different output on different runs

Common causes:

1. **Input files changed** - Same data = same output
2. **Field selection changed** - Use `--chart-fields` explicitly
3. **Pattern doesn't match files** - Check `--pattern` glob (default: `*.txt`)
4. **File encoding issues** - Script tries UTF-8, UTF-16, Latin-1

Troubleshoot with:

```powershell
# Check which files match
Get-ChildItem "c:\logs" -Filter "*.txt"

# Count extracted rows
python -c "import csv; print(sum(1 for _ in csv.DictReader(open('RMT_Combined_Similar.csv'))))"
```

---

## Performance & Optimization

### Extraction Speed

Typical performance on modern hardware:

| Stage | Input Size | Time | Notes |
|-------|-----------|------|-------|
| 1 (CSV) | 100 MB | 30 sec | File I/O bottleneck |
| 2 (Excel) | 100 MB | 1-2 min | Workbook generation |
| 3 (PPT) | 100 MB | 2-3 min | Chart rendering |
| 4 (JMP) | 100 MB | 3-5 min | JMP GUI overhead |

### Speed Tips

1. **Use Stage 1 first** if you just need CSV data
2. **Batch multiple files** in one run rather than separately
3. **Run JSL-only mode** (`--jmp-jsl-only`) to skip JMP GUI
4. **Pre-filter logs** - Only process relevant files with `--pattern`

### Memory Usage

- **Large logs (1+ GB):** Script loads entire file into memory
- **Workaround:** Split logs into smaller chunks and process separately

---

## Recommended Workflow

1. Keep this script as the primary extraction pipeline (reliable and repeatable).
2. Use AI/agent tools only for post-processing insights and anomaly commentary.
3. Version-control this script and output schema to keep reports stable across runs.
