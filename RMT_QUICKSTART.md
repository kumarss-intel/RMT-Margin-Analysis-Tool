# RMT Pipeline - Quick Start Guide

📖 **See Also:** [Main README](README.md) · [Full Documentation](RMT_LOG_PIPELINE_README.md)

## For Beginners: Easiest Way to Get Started

### Step 0: One-time environment setup

Open the cloned repository folder and double-click **`setup.bat`**, or run it from PowerShell:

```powershell
.\setup.bat
```

This creates the local `.venv` and installs all Python dependencies. You only do this once per machine.
JMP Pro is not installed by setup — it must already exist as a Windows application; setup just reports where it found it.

> Prefer a GUI? Double-click **`Launch_RMT_GUI.bat`** after setup and skip the rest of this guide.

### Step 1: Open PowerShell

Open Windows PowerShell or PowerShell Terminal and change into the cloned repository folder:

```powershell
cd C:\path\to\RMT-Margin-Analysis-Tool
```

### Step 2: Run the User-Friendly Wrapper

Copy and paste this command:

```powershell
& ".\.venv\Scripts\python.exe" ".\rmt_pipeline_runner.py"
```

Then press **Enter**.

### Step 3: Follow the Interactive Menu

The script will ask you questions:

```
Select Workflow Stage:

  1) 📄 CSV Only
     Extracts all START_RMT blocks to CSV files.
     Output: RMT_Combined_Similar.csv, RMT_Combined_Extended.csv, csv_by_file/

  2) 📊 CSV + Excel
     Adds a multi-sheet Excel workbook (one sheet per frequency + combined).
     Output: above + RMT_Extraction.xlsx

  3) 🎨 CSV + Excel + JMP Charts
     Generates per-parameter scatter charts in JMP.
     Output: above + jmp_charts/*.png  (one PNG per parameter + dashboard)

  4) 📑 CSV + Excel + JMP Charts + PPT  ← default
     Everything in stage 3, plus assembles the JMP PNGs into a PowerPoint.
     Output: above + RMT_Summary_JMP_Charts.pptx

Enter stage number (1-4) [default: 4]:
```

**For most users:** just press **Enter** to accept the default (Stage 4 — full report).

### Step 4: Provide Paths

The script will ask for:

```
Enter input folder or file path:
```

Example:
```
C:\path\to\logs
```

```
Enter output directory path [default: C:\path\to\logs]:
```

Press **Enter** to save outputs alongside your input logs, or type a different path.

### Done!

The script will start processing your files and show progress. When finished, check the output folder for:
- ✅ CSV files (data tables)
- ✅ Excel workbook (pivot-ready sheets)
- ✅ JMP chart PNGs — one per parameter + dashboard (`jmp_charts/`) *if stage 3+*
- ✅ JMP script (`rmt_jmp_charts.jsl`) *if stage 3+*
- ✅ PowerPoint from JMP charts (`RMT_Summary_JMP_Charts.pptx`) *if stage 4*

---

## For Automation: Quick Copy-Paste Commands

Use these if you want to run the same extraction repeatedly without prompts.

### Simplest (Just CSV + Excel):

```powershell
python ".\rmt_log_pipeline.py" --input "C:\path\to\logs" --pattern "*.txt" --outdir "C:\path\to\output" --no-ask-chart-approval "--chart-fields="
```

Replace `C:\path\to\logs` with your input folder.
Replace `C:\path\to\output` with your output folder.

---

## Real-World Examples

### Example 1: Daily Log Analysis
Extract new logs every morning and save with timestamp:

```powershell
$date = Get-Date -Format "yyyyMMdd"
python "rmt_pipeline_runner.py"
# Then enter: c:\lab_Logs\daily_run
# Then enter: c:\results\rmt_$date
```

### Example 2: Batch Processing Multiple Folders
Analyze all runs from a week:

```powershell
python "rmt_log_pipeline.py" `
  --input "c:\lab_Logs\Week47_Run1" "c:\lab_Logs\Week47_Run2" `
  --outdir "c:\results\week47_summary" `
  --no-ask-chart-approval `
  --chart-fields "RxDqVrefByte,TxVref"
```

### Example 3: Automated with JMP for Reporting
Scheduled daily extraction with full charts:

```powershell
$outdir = "c:\reports\$(Get-Date -Format 'yyyy-MM-dd')"
New-Item -ItemType Directory -Path $outdir -Force | Out-Null

python "rmt_log_pipeline.py" `
  --input "E:\validation_logs" `
  --pattern "rmt_*.txt" `
  --outdir $outdir `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" `
  --ppt-template "C:\path\to\template.pptx" `
  --jmp-axis-config ".\jmp_axis_settings.json"

Write-Host "Results saved to: $outdir"
```

---

### With PowerPoint (Stage 4 — JMP Charts + PPT):

```powershell
python ".\rmt_log_pipeline.py" `
  --input "C:\path\to\logs" `
  --pattern "*.txt" `
  --outdir "C:\path\to\output" `
  --no-ask-chart-approval `
  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" `
  --generate-jmp-charts `
  --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" `
  --ppt-template "C:\path\to\template.pptx" `
  --jmp-axis-config ".\jmp_axis_settings.json"
```

> **Tip:** To change Y-axis scale, tick increments, or reference lines, edit `jmp_axis_settings.json` in the repository root — no code changes needed.

---

## Troubleshooting

### "Command not found" error

Use the full path to the script:

```powershell
python ".\rmt_pipeline_runner.py"
```

### "No rows extracted"

Check that your input files contain `START_RMT` and lines like `Mc0.C0.R0: -32.0 32.0 ...`

```powershell
Select-String -Path "c:\logs\*.txt" -Pattern "START_RMT" | Select-Object -First 5
```

### Excel or PowerPoint not generated

```powershell
.\setup.bat --recreate
```

### JMP executable not found

```powershell
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"
```

If not installed, use Stage 2 (CSV + Excel) instead.
```powershell
Start-Process PowerShell -Verb RunAs
```

Or specify an output folder you have write access to:
```powershell
python rmt_pipeline_runner.py
# Enter: c:\users\YourName\rmt_output
```

### Charts not appearing in PowerPoint

Various reasons charts may not appear:

1. **No data in logs** - Verify log contains START_RMT and rank data
2. **Approval denied** - Script exits before generating PPT if you deny approval
3. **Missing python-pptx** - Re-run `.\setup.bat --recreate` from the tool folder
4. **Empty chart fields** - Use `--chart-fields "RxDqVrefByte,TxVref"` or similar

---

### Need help?

**Quick Reference:** [README.md](README.md#-frequently-asked-questions)

**Full Documentation:** [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md#troubleshooting)

**See Examples:** Scroll up to [Real-World Examples](#real-world-examples) section

---

## Command Cheat Sheet

| Task | Command |
|------|---------|
| Interactive mode | `python rmt_pipeline_runner.py` |
| Quick mode | `python rmt_pipeline_runner.py --quick` |
| CSV only | `"--chart-fields="` |
| CSV + Excel | `"--chart-fields="` |
| CSV + Excel + PPT | `--chart-fields "RxDqVrefByte,TxVref,ClkGrpPi"` |
| Add JMP charts | `--generate-jmp-charts --jmp-exe "C:\...\jmp.exe"` |
