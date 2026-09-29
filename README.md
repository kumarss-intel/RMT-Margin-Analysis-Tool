# RMT Margin Analysis Tool

This repository is the self-contained home of the RMT Margin Analysis Tool: everything needed to extract and analyze **START_RMT data blocks** from MRC debug logs.

> **First time here?** Run [`setup.bat`](setup.bat) once. It creates a local `.venv` and installs every Python dependency. JMP Pro is *not* installed by setup — it must already be present as a Windows application.

## 📋 Available Tools

### 1. **RMT Pipeline Wrapper** (Easiest)
**File:** `rmt_pipeline_runner.py`

User-friendly interactive interface for running the RMT extraction pipeline.

**Best for:** First-time users, interactive workflows, ad-hoc analyses

**Quick start:**
```powershell
python "rmt_pipeline_runner.py"
```

**Features:**
- 📚 Interactive menu with per-stage descriptions
- 🎯 Step-by-step path input prompts
- 🖼️ JMP chart generation and PPT assembly
- ❌ Clear error messages with solutions
- `--help` shows full stage/field/axis reference
- `--cmds` prints copy-paste CLI commands for all stages

**Learn more:** See [RMT_QUICKSTART.md](RMT_QUICKSTART.md)

---

### 2. **RMT Pipeline Core Script**
**File:** `rmt_log_pipeline.py`

Main extraction and analysis engine. Powers both the wrapper and direct CLI usage.

**Best for:** Automation, batch processing, CI/CD integration, advanced users

**Quick start (Stage 2: CSV + Excel):**
```powershell
python "rmt_log_pipeline.py" `
  --input "C:\path\to\logs" `
  --pattern "*.txt" `
  --outdir "C:\path\to\output"
```

**Features:**
- 📄 CSV extraction in standard + extended formats
- 📊 Excel workbooks with multi-sheet pivot-ready output
- 📈 PowerPoint presentations with charts
- 🎨 JMP chart generation with CLI automation
- ✅ Approval workflow for chart generation
- 🛑 Hard-stop on denial with exit codes

**Learn more:** See [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md)

---

## 🚀 Quick Navigation

| Goal | Start Here | Command |
|------|-----------|--------|
| First time using? | [RMT_QUICKSTART.md](RMT_QUICKSTART.md) | `python rmt_pipeline_runner.py` |
| See all options? | Run the script | `python rmt_pipeline_runner.py --help` |
| Copy-paste CLI commands? | Run the script | `python rmt_pipeline_runner.py --cmds` |
| Need full CLI reference? | [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md) | `python rmt_log_pipeline.py --help` |

---

## 📦 Workflow Stages

The pipeline supports 4 stages of increasing output. **Stage 4 is the default.**

| Stage | Output | Default? | Approx. time |
|-------|--------|---------|--------------|
| **1** | CSV | | 🚀 ~30 sec |
| **2** | CSV + Excel | | ⚡ 1-2 min |
| **3** | CSV + Excel + JMP Charts | | 🎨 3-5 min |
| **4** | CSV + Excel + JMP Charts + PPT | ✅ yes | 📑 3-5 min |

---

## 🔧 Common Tasks

### Extract data to CSV only
```powershell
python "rmt_pipeline_runner.py"
# Select Stage 1
```

### Generate Excel with pivot tables
```powershell
python "rmt_pipeline_runner.py"
# Select Stage 2
```

### Generate JMP scatter charts only (no PPT)
```powershell
python "rmt_pipeline_runner.py"
# Select Stage 3
```

### Full analysis — JMP charts + PPT (default)
```powershell
python "rmt_pipeline_runner.py"
# Press Enter at the stage prompt (default is 4)
```

### Customise JMP chart axis scales
Edit `jmp_axis_settings.json` in this folder, then re-run Stage 3 or 4.
Each parameter has `plus`/`minus` blocks with `min`, `max`, `inc`, and `ref_line` values.

---

## 📊 Output Structure

After running the pipeline, your output folder will contain:

```
output_dir/
├── RMT_Combined_Similar.csv          # Standard format (like your reference)
├── RMT_Combined_Extended.csv         # Standard + metadata columns
├── csv_by_file/
│   ├── file1_RMT.csv
│   ├── file2_RMT.csv
│   └── ...
├── RMT_Extraction.xlsx               # Multi-sheet Excel workbook
│   ├── All_RMT (combined)
│   ├── Freq_5600, Freq_6400, ... (by frequency)
│   └── File_log1, File_log2, ... (by source file)
├── RMT_Summary_JMP_Charts.pptx       # PPT built from JMP PNG exports (Stage 4)
└── jmp_charts/
    ├── RecEnDelay.png                 # Per-parameter: X=Params, Y=±values
    ├── TxDqsDelay.png
    ├── RxDqsDelay.png
    ├── TxDqDelay.png
    ├── RxDqVrefByte.png
    ├── TxVref.png
    ├── ClkGrpPi.png
    ├── CmdVref.png
    └── RMT_Dashboard.png             # All parameters combined
```

---

## 🔧 Installation & Setup

### Prerequisites

- Windows
- Python 3.10+ (3.14 recommended) installed with the **tcl/tk and IDLE** option, so `tkinter` is available for the GUI
- Optional: SAS JMP Pro (licensed Windows install) for the JMP chart and PPT stages

### One-time setup

From this folder, double-click `setup.bat` or run:

```powershell
.\setup.bat
```

`setup.bat` will:

1. Locate a Python 3.10+ interpreter (prefers `C:\Program Files\Python314`, then the `py` launcher, then `PATH`)
2. Create the local virtual environment `.venv`
3. Install the pinned dependencies from [requirements.txt](requirements.txt) (`openpyxl`, `python-pptx`, `pillow`)
4. Verify every package imports correctly
5. Detect the existing JMP installation and report its path (warns, but does not fail, when JMP is absent)

To rebuild the environment from scratch:

```powershell
.\setup.bat --recreate
```

### Running the tool

| How | Command |
|-----|---------|
| GUI | Double-click `Launch_RMT_GUI.bat` (uses `.venv` automatically) |
| Interactive CLI | `.venv\Scripts\python.exe rmt_pipeline_runner.py` |
| Raw pipeline CLI | `.venv\Scripts\python.exe rmt_log_pipeline.py --help` |

Prefer working in an activated shell? Run `.venv\Scripts\activate.bat`, then plain `python` commands resolve to the virtual environment.

### About JMP (excluded from the virtual environment)

JMP Pro is licensed desktop software and can never be installed by pip. Setup only detects it, via the
`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe` registry entry, the standard
`C:\Program Files\SAS\JMPPRO\<version>` folders, and `PATH`.

```powershell
# Verify JMP installation manually
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"
```

Without JMP, log parsing, CSV, Excel and the HTML report all still work, and `--jmp-jsl-only` produces the JSL script to run later on a JMP-equipped machine.

---

## 📖 Documentation Files

| File | Purpose | Audience |
|------|---------|----------|
| [README.md](README.md) | This index (you are here) | Everyone |
| [setup.bat](setup.bat) | One-time environment setup (`.venv` + dependencies) | Everyone |
| [requirements.txt](requirements.txt) | Pinned Python dependencies | Everyone |
| [RMT_QUICKSTART.md](RMT_QUICKSTART.md) | 5-minute beginner guide | New users |
| [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md) | Complete reference manual | Advanced users |
| [jmp_axis_settings.json](jmp_axis_settings.json) | JMP chart axis customisation | Stage 4 users |
| [.github/skills/rmt-margin-plots/SKILL.md](.github/skills/rmt-margin-plots/SKILL.md) | Copilot agent skill (agentic use) | Copilot users |

---

## ❓ Frequently Asked Questions

### Q: Which tool should I use?
**A:** Use `rmt_pipeline_runner.py` for your first run. It's interactive and guides you through selections. For scripting/automation, use `rmt_log_pipeline.py` directly.

### Q: How do I extract specific fields for charts?
**A:** When prompted, enter field names or numbers:
- By number: `1,2,3` → RecEnDelay, TxDqsDelay, RxDqsDelay
- By name: `RxDqVrefByte,TxVref` → Those fields only
- All: type `all`
- None: type `none`

### Q: Can I generate charts without PowerPoint?
**A:** Yes! Use Stage 4 (JMP) which generates PNG images directly from the data.

### Q: What if I deny chart approval?
**A:** The script exits cleanly with exit code 3. CSV and Excel still generate fine. No PPT will be created.

### Q: How do I run this in batch/automated mode?
**A:** Use `rmt_log_pipeline.py` with `--no-ask-chart-approval` and `--chart-fields` specified. See [Advanced Usage](RMT_LOG_PIPELINE_README.md#advanced-usage-direct-script-commands).

### Q: Where are my output files?
**A:** In the folder you specified with `--outdir` or the output path you entered in the prompts. See [Output Structure](#output-structure) above.

---

## 🐛 Troubleshooting

### Script won't start
```powershell
# Check Python is installed
python --version

# Verify script exists
Test-Path "rmt_pipeline_runner.py"

# Check file permissions
Get-Acl "rmt_pipeline_runner.py"
```

### No data extracted
- ✅ Verify input files contain `START_RMT` text
- ✅ Check for lines like `Mc0.C0.R0: -32.0 32.0 ...`
- ✅ Ensure file extension is `.txt` (or adjust `--pattern`)

### Missing CSV/Excel/PPT
```powershell
# Rebuild the environment with all required packages
.\setup.bat --recreate

# Verify installation
.venv\Scripts\python.exe -c "import openpyxl, pptx; print('OK')"
```

### Permission denied
- Run PowerShell as Administrator
- Check folder write permissions
- Ensure input files are readable

### JMP charts not generating
```powershell
# Check JMP is installed
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"

# Use JSL-only mode to test script generation
python "rmt_log_pipeline.py" ... --jmp-jsl-only
```

### Exit code 3 (chart approval denied)
- This is intentional behavior when you deny chart approval
- CSV and Excel still generate successfully
- Re-run and approve if you need PPT charts

---

## 🎓 Example Workflows

### Workflow 1: Quick Analysis (2 minutes)
```powershell
# Interactive mode
python "rmt_pipeline_runner.py"
# → Select Stage 2
# → Enter paths
# → Get CSV + Excel output
```

### Workflow 2: Automated Daily Extraction
```powershell
# Batch script - runs daily at 6 AM
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
python "rmt_log_pipeline.py" `
  --input "E:\logs\daily_run" `
  --pattern "*.txt" `
  --outdir "E:\rmt_results\$timestamp" `
  --no-ask-chart-approval `
  "--chart-fields="
```

### Workflow 3: Full Presentation Report
```powershell
# Interactive: select Stage 3
python "rmt_pipeline_runner.py"
# → Choose Stage 3
# → Approve specific fields when prompted
# → Get PowerPoint with your selected charts
```

### Workflow 4: Complete Analysis with JMP
```powershell
# Direct script with all outputs
python "rmt_log_pipeline.py" `
  --input "C:\path\to\logs" `
  --pattern "*.txt" `
  --outdir "c:\rmt_analysis" `
  --no-ask-chart-approval `
  --chart-fields "RxDqVrefByte,TxVref,ClkGrpPi" `
  --generate-jmp-charts `
  --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe"
```

---

## 📝 Column Reference

### CSV Output Columns (Standard Format)
- `Params` - Rank identifier (e.g., Mc0.C0.R0)
- `BootTemp`, `RunTemp` - Temperature readings
- `RecEnDelay-`, `RecEnDelay+` - Timing window
- `TxDqsDelay-`, `TxDqsDelay+` - Timing window
- `RxDqsDelay-`, `RxDqsDelay+` - Timing window
- `TxDqDelay-`, `TxDqDelay+` - Timing window
- `RxDqVrefByte-`, `RxDqVrefByte+` - Voltage window
- `TxVref-`, `TxVref+` - Voltage window
- `ClkGrpPi-`, `ClkGrpPi+` - Clock timing
- `CmdVref-`, `CmdVref+` - Command voltage

### CSV Output Columns (Extended Format - adds)
- `SourceFile` - Input filename
- `Frequency` - Memory frequency (e.g., 5600, 6400)
- `Gear` - Gear ratio (e.g., G2, G1)
- `BlockIndex` - START_RMT block number in file

---

## 🤖 Agentic Use (GitHub Copilot)

The repository ships its own Copilot agent skill in
[.github/skills/rmt-margin-plots/SKILL.md](.github/skills/rmt-margin-plots/SKILL.md).
It is self-contained: every path it uses is resolved relative to this repository,
so it works from any clone location.

1. Clone the repository and run `setup.bat` once.
2. Open the repository folder in VS Code with GitHub Copilot Chat in **Agent** mode.
3. Ask, for example:
   - *"Plot RMT margins for the logs in C:\path\to\logs"*
   - *"Generate RMT JMP charts for RxDqVrefByte and TxVref from C:\path\to\logs into C:\path\to\output"*
   - *"Build the RMT report for these thermal (BCRH) logs"*

The agent runs `rmt_log_pipeline.py` non-interactively with the `.venv` interpreter,
detects JMP Pro when it is installed (falling back to CSV + Excel + HTML + PPT when it
is not), and reports the generated artifacts. It never launches the GUI or the
interactive runner.

---

## 📞 Support

- **Script Issues:** Check [RMT_QUICKSTART.md](RMT_QUICKSTART.md#troubleshooting) Troubleshooting section
- **Usage Questions:** See [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md#troubleshooting)
- **Advanced Options:** Run `python rmt_log_pipeline.py --help`

---

## 📜 License & Notes

This tool is part of the Intel Memory Reference Code (MRC) development environment.

**Version:** 1.0  
**Last Updated:** 2026-06-23  
**Tested on:** Python 3.14.4, Windows PowerShell 5.1+
