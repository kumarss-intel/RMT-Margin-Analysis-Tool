# MarginIQ

### Intel CCG CVE DDR5 RMT Margin Analysis Tool

<img src="assets/marginiq.png" alt="MarginIQ icon" width="72" align="right">

**MarginIQ** (v2.0.0) turns Intel MRC debug logs into margin analysis. It extracts every
`START_RMT` (Rank Margin Tool) block and produces CSV, Excel, an interactive HTML report
(including DDR5 Mode Register and ODT tabs), PowerPoint decks and optional JMP Pro charts.
It is project-agnostic: Nova Lake, Wildcat Lake and future platforms are listed in
[`projects.json`](projects.json).

You can drive it in three ways:

| Mode | Best for | Entry point |
|------|----------|-------------|
| 🖥️ **GUI** | Day-to-day lab analysis, axis tuning, JMP Graph Builder | **MarginIQ** desktop icon / `Launch_RMT_GUI.bat` |
| 🤖 **Agentic** (GitHub Copilot) | "Plot these logs for me" from chat, no clicks | Copilot Chat in **Agent** mode ([details](#-agentic-mode-github-copilot)) |
| ⌨️ **CLI** | Scripts, batch jobs, CI | `.venv\Scripts\python.exe rmt_log_pipeline.py` |

---

## 📑 Contents

1. [What's new in 2.0](#-whats-new-in-20)
2. [Quick start](#-quick-start)
3. [Installation & setup](#-installation--setup) — incl. [desktop / taskbar shortcuts](#desktop-taskbar-and-start-menu-shortcuts)
4. [Using the GUI](#-using-the-gui)
5. [Agentic mode (GitHub Copilot)](#-agentic-mode-github-copilot)
6. [Command line](#-command-line)
7. [HTML report](#-html-report)
8. [Configuration](#-configuration) — projects, axis settings, default ±Ref
9. [Output structure](#-output-structure) and [column reference](#-column-reference)
10. [Documentation files](#-documentation-files)
11. [FAQ](#-frequently-asked-questions)
12. [Troubleshooting](#-troubleshooting)

---

## ✨ What's new in 2.0

| Area | Change |
|------|--------|
| Branding | Product name **MarginIQ** with the subtitle *Intel CCG CVE DDR5 RMT Margin Analysis Tool* in the GUI, HTML report and PPT. Version via `--version` or Help ▸ About. |
| Icon & shortcuts | Unique MarginIQ icon ([`assets/marginiq.ico`](assets/marginiq.ico)) on the GUI window, taskbar button and header. `setup.bat` creates a **Desktop shortcut** and asks whether to **pin to the taskbar** and add a **Start menu / Quick Launch** entry. `Launch_RMT_GUI.bat` gets an icon'd twin, `MarginIQ.lnk`. |
| Project-agnostic | **Project** dropdown on Tab 1 and `--project NVL\|WCL` on the CLI, driven by `projects.json` (name, code, START_RMT header aliases, axis file). Add a project without code changes. WCL-style `RxVref` headers and per-byte `Mc0.C0.B0.R0` rows are parsed. |
| Workflow | **↺ Analyze Another Log** replaces the disabled *Next* button on Tab 3 after a successful run; also *File ▸ New Analysis* (`Ctrl+N`). It resets Tab 1 to the launch state (*Boot / MRC logs* source, empty file list, default stage/profile, locked tabs, default axis values). Project, JMP path and PPT template are kept. |
| Tab 2 layout | Grouped toolbar: **Axis Values** (apply / recalculate), **Axis Presets** (Save JSON, Load JSON, Load .jrp) and **JMP Graph Builder** (opens a live, editable Graph Builder with the current axis values). |
| Default ±Ref | Configurable **+ Ref / − Ref** in the Tab 2 *Defaults*, saved as `defaults.ref_line_plus/minus`. Used whenever a parameter's own Ref is empty: JMP ref lines (now also on DTR charts), PASS/WARN/FAIL status and red raw-data cells. *Apply ±Ref to all* copies it to every charted parameter. |
| HTML report | New **Mode Registers** and **ODT** tabs; the **JMP Charts** tab is now a dropdown-driven side-by-side comparison workspace. See [HTML report](#-html-report). |
| Tab 3 | *Open HTML Report* button. |

---

## 🚀 Quick start

```powershell
git clone https://github.com/kumarss-intel/RMT-Margin-Analysis-Tool.git
cd RMT-Margin-Analysis-Tool
.\setup.bat          # once: .venv + dependencies + MarginIQ desktop icon
```

Then do **one** of the following:

- **GUI:** double-click the **MarginIQ** desktop icon. On Tab 1 pick the *Project*, add your logs and click *Generate CSV & Continue*. Tune the axes on Tab 2, click *Run Pipeline* on Tab 3, then *↺ Analyze Another Log* for the next set.
- **Agentic:** open the folder in VS Code, switch Copilot Chat to **Agent** and ask *"Plot the RMT margins for the NVL logs in C:\path\to\logs"*.
- **CLI:** run the pipeline directly:
  ```powershell
  .venv\Scripts\python.exe rmt_log_pipeline.py --project NVL --input "C:\path\to\logs" --pattern "*.log" --outdir "C:\path\to\out" --no-ask-chart-approval
  ```

---

## 🔧 Installation & setup

### Prerequisites

- Windows 10 or 11
- Python 3.10+ (3.14 recommended), installed with the **tcl/tk and IDLE** option so `tkinter` is available for the GUI
- Optional: SAS JMP Pro (licensed Windows install) for JMP charts, the JMP PPT and Graph Builder
- Optional (agentic mode): VS Code with GitHub Copilot Chat
- Access to this private repository (ask the owner to add you as a collaborator)

### Get the code

```powershell
git clone https://github.com/kumarss-intel/RMT-Margin-Analysis-Tool.git
cd RMT-Margin-Analysis-Tool
```

Use `git clone` rather than **Download ZIP**, so you can pick up fixes later with `git pull`.
If the clone fails with `Unsupported proxy syntax in 'proxy-server:port'`, see
[Cloning behind a broken proxy](#cloning-behind-a-broken-proxy).

### One-time setup

From this folder, double-click `setup.bat` or run `.\setup.bat`. It performs six steps:

1. Locates a Python 3.10+ interpreter (prefers `C:\Program Files\Python314`, then the `py` launcher, then `PATH`).
2. Creates the local virtual environment `.venv`.
3. Checks the proxy settings pip will use (stopping with a clear message if one is malformed), then installs the pinned [requirements.txt](requirements.txt) (`openpyxl`, `python-pptx`, `pillow`).
4. Verifies every package imports correctly.
5. Detects the existing JMP installation and reports its path. It warns, but does not fail, when JMP is absent.
6. **Creates the MarginIQ shortcuts** (see below).

| Option | Effect |
|--------|--------|
| `setup.bat` | Normal install. Creates the Desktop shortcut and **asks** about the taskbar pin and the Start menu / Quick Launch entry. |
| `setup.bat --recreate` | Deletes `.venv` and rebuilds it from scratch. |
| `setup.bat --proxy http://proxy-chain.intel.com:912` | Uses this proxy for pip, for this run only. |
| `setup.bat --all-shortcuts` | Creates the Desktop, Start menu and Quick Launch shortcuts plus the taskbar pin, without asking. |
| `setup.bat --no-prompt` | Desktop shortcut only, never asks (unattended or agent runs). |
| `setup.bat --no-shortcuts` | Creates no shortcuts. |
| `setup.bat --help` | Shows all options. |

### Desktop, taskbar and Start menu shortcuts

Every shortcut shows the **MarginIQ icon** and starts the GUI exactly like `Launch_RMT_GUI.bat`:
the `.venv` `pythonw.exe` runs `rmt_gui.py` with no console window.

| Shortcut | Created | Location |
|----------|---------|----------|
| **Desktop** | Always (unless `--no-shortcuts`) | `<Desktop>\MarginIQ.lnk` |
| **Launcher twin** | Always, and on the first run of `Launch_RMT_GUI.bat` | `<tool folder>\MarginIQ.lnk`, next to `Launch_RMT_GUI.bat` |
| **Taskbar** | Asked: *"Pin MarginIQ to the taskbar [Y,N]?"* | Pinned taskbar icon |
| **Start menu + Quick Launch** | Asked: *"Add MarginIQ to the Start menu and Quick Launch toolbar [Y,N]?"* | *Start ▸ All apps ▸ MarginIQ* and `%APPDATA%\Microsoft\Internet Explorer\Quick Launch` |

Each prompt defaults to **No** after 30 seconds, so setup never hangs.

> **Why the taskbar needs one click from you.** Since Windows 10 (1809), Windows blocks
> programs from pinning themselves to the taskbar. When you answer **Y**, setup creates the
> Start-menu shortcut, tries to pin, and if Windows refuses it prints the steps (no window is
> opened). Open *Start ▸ All apps*, right-click **MarginIQ** ▸ *Pin to taskbar*. Alternatively,
> start MarginIQ and right-click its taskbar button ▸ *Pin to taskbar*.
>
> The GUI and its shortcuts share the AppUserModelID `Intel.CCG.CVE.MarginIQ`, so the running
> window always shows the MarginIQ icon and groups under the pinned button.

> **Why `Launch_RMT_GUI.bat` itself shows the generic batch icon.** Windows cannot attach an
> icon to a `.bat` file. The icon'd `MarginIQ.lnk` next to it is the same launcher; copy or pin
> that shortcut instead.

> **Quick Launch toolbar.** Windows 10 shows it only when you enable it (*taskbar ▸ Toolbars ▸
> New toolbar ▸* `%APPDATA%\Microsoft\Internet Explorer\Quick Launch`). Windows 11 has no Quick
> Launch toolbar, so use the Start menu entry or the taskbar pin there.

Manage the shortcuts at any time without re-running setup:

```powershell
# create / refresh (any combination of -Desktop -Local -StartMenu -QuickLaunch -Taskbar)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\create_shortcuts.ps1 -Desktop -StartMenu -Taskbar

# remove every MarginIQ shortcut
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\create_shortcuts.ps1 -Remove
```

The icon is generated by [`scripts/make_icon.py`](scripts/make_icon.py): a DDR eye diagram
with a cyan margin arrow on an Intel-blue tile, at 16–256 px. The generated
`assets\marginiq.ico` / `marginiq.png` are committed; re-run the script only if you change the
design.

### Running the tool

| How | Command |
|-----|---------|
| GUI | **MarginIQ** desktop / Start / taskbar icon, `MarginIQ.lnk`, or `Launch_RMT_GUI.bat` |
| Agentic | VS Code ▸ Copilot Chat ▸ **Agent** mode — see [Agentic mode](#-agentic-mode-github-copilot) |
| Pipeline CLI | `.venv\Scripts\python.exe rmt_log_pipeline.py --help` |
| Interactive CLI menu | `.venv\Scripts\python.exe rmt_pipeline_runner.py` |

Prefer an activated shell? Run `.venv\Scripts\activate.bat`; plain `python` then resolves to the virtual environment.

### About JMP (excluded from the virtual environment)

JMP Pro is licensed desktop software and can never be installed by pip. Setup only detects it, via the
`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe` registry entry, the standard
`C:\Program Files\SAS\JMPPRO\<version>` folders, and `PATH`.

Without JMP, log parsing, CSV, Excel, the HTML report and the native PPT all still work, and
`--jmp-jsl-only` produces the JSL script to run later on a JMP-equipped machine.

---

## 🖥️ Using the GUI

The window has a **File** menu (*New Analysis* `Ctrl+N`, *Open Output Folder*, *Exit*), a **Help**
menu (*User Guide*, *Open Debug Log Folder*, *About MarginIQ*), and three tabs driven by the
*Previous / Next* bar.

### Tab 1 · Input & Workflow

| Section | What to do |
|---------|------------|
| **Project & Experiment Profile** | **Project** (Nova Lake, Wildcat Lake, …) labels the reports and selects that platform's START_RMT header aliases. **Profile**: *Non-Thermal* or *Thermal Experiment (DTR)* for BCRH / BHRC logs (adds `--dtr`). |
| **Data Source** | *Boot / MRC logs* (default), *Existing CSV*, *Existing Excel* or *PNG charts → PPT only*. |
| **Files / Folder** | *Add Files…* / *Add Folder…* (filtered by *File Mask*, e.g. `*.log;*.txt`). |
| **Multiple-File Handling** | *Concatenate* into one report (overlay by source) or process each file *Separately* (one subfolder each). |
| **Output**, **PPT Template**, **Workflow Stage**, **Tools & Options** | Output folder (auto-filled), file names, optional `.pptx` template, stage 1–4, JMP executable, JSL-only. |

**Generate CSV & Continue →** runs Stage 1 and unlocks Tabs 2–3, which then reuse that CSV.

### Tab 2 · Parameters & Axis

| Group | Controls |
|-------|----------|
| **Axis Values** | *Apply axis settings*; **↻ Recalculate from data** (Min / Max / Inc from the logs; ±Ref preserved). |
| **Axis Presets** | **Save JSON…**, **Load JSON…**, **Load .jrp…** (clone scales from a JMP report). |
| **JMP Graph Builder (interactive)** | Choose *All charted parameters* or one parameter, then **▶ Open**. JMP opens live Graph Builder windows using the current axis values. Fine-tune them, save as `.jrp` in JMP, and bring the scales back with *Load .jrp…*. |
| **Defaults** | Inc, minor ticks, grids, **+ Ref / − Ref** (default ±10), *Apply ±Ref to all*, ref-line colours. |
| **Per-Parameter Y-Axis** | *Chart* checkbox, +/− Min / Max / Inc / Ref per parameter, and a live eye-preview. An empty Ref uses the default ±Ref. |

### Tab 3 · Preview & Run

An editable command preview (what you see is what runs), **▶ Run Pipeline**, **■ Stop**,
*Open Output Folder*, *Open HTML Report* and a live output log. After a successful run the
bottom-right button becomes **↺ Analyze Another Log**. It clears the session back to the
launch state, so the next log set starts on Tab 1 with *Boot / MRC logs* selected.

Every GUI action is logged in detail to `logs\rmt_gui_debug_<timestamp>.log`
(*Help ▸ Open Debug Log Folder*). Attach it when reporting a problem.

---

## 🤖 Agentic mode (GitHub Copilot)

MarginIQ ships a Copilot **agent skill**,
[`.github/skills/rmt-margin-plots/SKILL.md`](.github/skills/rmt-margin-plots/SKILL.md), plus
repository instructions,
[`.github/copilot-instructions.md`](.github/copilot-instructions.md). Together they let
Copilot run the full pipeline from a chat request. The skill is self-contained: every path is
resolved from `git rev-parse --show-toplevel`, so it works from any clone location.

### Set up

1. Clone the repository and run `setup.bat` once. The agent needs `.venv`; if it is missing,
   the agent runs `setup.bat --no-prompt` itself.
2. Open the **repository folder** in VS Code (the skill is discovered from `.github/skills`).
3. In Copilot Chat, select **Agent** mode.

### Ask in plain language

| You ask | The agent runs |
|---------|----------------|
| *"Plot RMT margins for the NVL logs in C:\lab\NVL\run1"* | Base pipeline with `--project NVL`: CSV, Excel, HTML (MR / ODT tabs) and PPT |
| *"Generate JMP charts for RxDqVrefByte and TxVref from C:\lab\logs into C:\lab\out"* | `--generate-jmp-charts --jmp-exe <detected> --chart-fields RxDqVrefByte,TxVref` |
| *"Build the thermal (BCRH/BHRC) report for C:\lab\NVL\DTR\...\test"* | `--dtr`; the agent confirms the profile first |
| *"Analyze these Wildcat Lake logs"* | `--project WCL` (handles `RxVref` and per-byte rows) |
| *"What RTT_WR / RTT_PARK did training pick for these logs?"* | Base run, then reads the **ODT** / **Mode Registers** tabs or `rmt_metadata.json` |
| *"Re-chart the existing CSV with ±12 reference lines"* | Copies `jmp_axis_settings.json` into the output folder, sets `ref_line_plus/minus` to ±12, then runs `--jmp-from-csv ... --jmp-axis-config <copy>` |
| *"Rebuild the PPT from the charts in C:\lab\out\jmp_charts"* | `--ppt-from-charts` |

### How the agent works

1. **Confirms inputs.** It asks for anything missing (input folder, output folder, project,
   JMP yes/no, parameters) and never invents paths. It infers the project only when the path
   or log is unambiguous (e.g. `\NVL\`, `Detected board: WCL …`).
2. **Resolves the tool.** Repository root, `.venv\Scripts\python.exe`, and JMP from the
   registry or standard folders. Without JMP it falls back to base mode (CSV, Excel, HTML, PPT).
3. **Sanity-checks the logs.** It looks for `START_RMT` before starting a long run.
4. **Runs `rmt_log_pipeline.py` non-interactively.** Always with `--no-ask-chart-approval` and
   a long wait, because JMP runs take 3–5 minutes.
5. **Reports.** It lists every artifact with its full path and flags anything skipped. It reads
   the exit code: `0` success, `1` no files, `2` bad path / unknown project / no RMT blocks,
   `3` approval denied, `4` JMP requested without `--jmp-exe`, `5` JMP failed.

### Guardrails built into the skill

- Never starts `rmt_gui.py`, `Launch_RMT_GUI.bat` or `rmt_pipeline_runner.py` (interactive, would hang).
- Never uses bare `python` or a hard-coded, user-specific path.
- Never runs a JMP job without your approval of the chart parameters (asked in chat).
- Runs `setup.bat --no-prompt` (no shortcut questions) if the environment is missing.

### Agentic tips

- Mention **DTR / thermal / BCRH / BHRC** for temperature-drift logs, and the **project** name.
- Ask follow-ups on the results, for example *"which rank has the worst TxVref margin?"* or
  *"compare MR34 between 4800 and 5200"*. The agent can read the generated CSV and
  `rmt_metadata.json`.
- For a one-off ±Ref or axis change, ask for it explicitly. The agent edits a **copy** in the
  output folder, never the shipped `jmp_axis_settings.json`.
- To teach the agent about a new platform, add it to `projects.json`. No skill change is needed.

---

## ⌨️ Command line

```powershell
.venv\Scripts\python.exe rmt_log_pipeline.py `
  --project NVL `
  --input "C:\path\to\logs" `
  --pattern "*.log" `
  --outdir "C:\path\to\out" `
  --no-ask-chart-approval `
  --generate-jmp-charts `
  --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" `
  --jmp-axis-config ".\jmp_axis_settings.json"
```

| Key flags | Purpose |
|-----------|---------|
| `--project NVL\|WCL\|"Wildcat Lake"` | Platform from `projects.json` (default: registry default, `NVL`) |
| `--dtr` | Thermal Experiment (BCRH / BHRC) parsing: Boot-temp and Run-temp RMT per log |
| `--chart-fields a,b,c` | Parameters to chart (omit for all) |
| `--generate-jmp-charts --jmp-exe <exe>` | JMP PNG charts and the JMP PPT |
| `--jmp-from-csv <csv…>` / `--jmp-from-excel <xlsx>` / `--ppt-from-charts <dir>` | Reuse existing artifacts |
| `--jmp-jsl-only` | Write the JSL without launching JMP |
| `--version` | Print the MarginIQ version |

Full reference: [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md). Beginner walkthrough:
[RMT_QUICKSTART.md](RMT_QUICKSTART.md).

### Workflow stages

| Stage | Output | Approx. time |
|-------|--------|--------------|
| **1** | CSV | ~30 s |
| **2** | CSV + Excel | 1–2 min |
| **3** | CSV + Excel + JMP charts | 3–5 min |
| **4** *(default)* | CSV + Excel + JMP charts + PPT | 3–5 min |

The HTML report and the native PPT are written in every stage that parses logs.

---

## 📊 HTML report

`RMT_Report.html` is a single self-contained file. Share it by e-mail or on a Teams / SharePoint site.

| Tab | Content |
|-----|---------|
| Overview | KPI cards, margin degradation below ±Ref, window-width chart, statistics (Cpk, guardband), interactive multi-source dashboard |
| Frequency / Parameters / RunTemp | Per-frequency, per-rank and temperature-drift charts |
| Training Steps / Platform | MRC task timeline; board, CPU and DIMM (SPD) details |
| **Mode Registers** | Final per-rank DDR5 MR values (from *SAGV Finalization*), JEDEC decodes (CL, Ron, VrefDQ/CA/CS, RTT_*, ODTL), amber = changed by training vs the JEDEC-reset value, bold = differs between ranks, filters, and a **cross-log MR diff** |
| **ODT** | DIMM ODT summary (RonUp/Dn, **RTT_WR**, NomWr, NomRd, Park, ParkDqs, CA/CS groups), CPU read ODT, per-rank ODT decoded from MR32–36, ODT latency offsets, BIOS ODT inputs, and a **cross-log ODT matrix** |
| **JMP Charts** | Comparison workspace. Every panel has a chart dropdown; set 1–3 panels per row, add or remove panels, reset. Pre-filled with every chart, so it is never blank. |
| Raw Data | Filterable table; cells weaker than ±Ref are highlighted red |

---

## ⚙️ Configuration

### Projects — [`projects.json`](projects.json)

```json
"WCL": { "name": "Wildcat Lake", "code": "WCL", "memory": "DDR5",
         "param_aliases": { "RxVref": "RxDqVrefByte" },
         "axis_config": "jmp_axis_settings.json" }
```

Add an entry to support a new platform. It appears in the GUI dropdown and is accepted by
`--project`. `param_aliases` maps a log's START_RMT header name to MarginIQ's canonical
parameter name.

### Axis settings and default ±Ref — [`jmp_axis_settings.json`](jmp_axis_settings.json)

`defaults` holds `inc`, `minor_ticks`, grids, ref-line colours and **`ref_line_plus` /
`ref_line_minus`** (the default ±Ref, ±10). Each parameter has `plus` / `minus` blocks with
`min`, `max`, `inc` and `ref_line`. A missing `ref_line` falls back to the default ±Ref for JMP
ref lines, PASS/WARN/FAIL and raw-data highlighting. Edit the file in Tab 2 or by hand; no
code change is needed.

### Branding / version — [`rmt_project.py`](rmt_project.py)

`TOOL_NAME`, `TOOL_SUBTITLE` and `TOOL_VERSION` are defined once and used by the GUI, CLI,
HTML and PPT.

---

## 📁 Output structure

```
output_dir/
├── RMT_Combined_Similar.csv        # standard columns
├── RMT_Combined_Extended.csv       # + SourceFile / Frequency / Gear / BlockIndex
├── csv_by_file/<log>_RMT.csv       # one CSV per source log
├── RMT_Extraction.xlsx             # All_RMT + per-file + per-frequency sheets
├── RMT_Report.html                 # interactive report (MR / ODT / JMP comparison tabs)
├── RMT_Summary.pptx                # native PowerPoint summary
├── rmt_metadata.json               # training steps, platform, MR / ODT data (reused by --jmp-from-csv)
├── rmt_jmp_charts.jsl              # JMP script (JMP stages)
├── rmt_graph_builder.jsl           # interactive Graph Builder script (GUI Tab 2)
├── RMT_Summary_JMP_Charts.pptx     # deck built from the JMP PNGs
└── jmp_charts/<Parameter>.png      # one Graph Builder chart per parameter
```

---

## 📝 Column reference

| Column | Meaning |
|--------|---------|
| `SourceFile`, `Frequency`, `Gear`, `BlockIndex` | Extended format only: source log, MT/s, gear, START_RMT block number |
| `Params` | Rank (`Mc0.C0.R0`) or byte-rank (`Mc0.C0.B0.R0`, WCL) |
| `BootTemp`, `RunTemp` | DTR temperatures snapped to 0 / 90 °C (empty for non-thermal) |
| `<Param>-`, `<Param>+` | Negative / positive margin for RecEnDelay, TxDqsDelay, RxDqsDelay, TxDqDelay, RxDqVrefByte, TxVref, ClkGrpPi, CmdVref |

---

## 📖 Documentation files

| File | Purpose |
|------|---------|
| [README.md](README.md) | This user guide |
| [RMT_QUICKSTART.md](RMT_QUICKSTART.md) | 5-minute beginner guide (GUI, agent, CLI) |
| [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md) | Complete CLI reference manual |
| [.github/skills/rmt-margin-plots/SKILL.md](.github/skills/rmt-margin-plots/SKILL.md) | Copilot agent skill (agentic mode) |
| [.github/copilot-instructions.md](.github/copilot-instructions.md) | Repository rules for Copilot |
| [setup.bat](setup.bat) | Environment setup and shortcuts |
| [scripts/create_shortcuts.ps1](scripts/create_shortcuts.ps1) | Create / remove the MarginIQ Desktop / Start / Quick Launch / taskbar shortcuts |
| [scripts/make_icon.py](scripts/make_icon.py) | Regenerates `assets/marginiq.ico` / `.png` |
| [scripts/check_proxy.py](scripts/check_proxy.py) | Proxy pre-flight check run by `setup.bat` |
| [projects.json](projects.json) | Project registry |
| [jmp_axis_settings.json](jmp_axis_settings.json) | Axis scales and default ±Ref |
| [requirements.txt](requirements.txt) | Pinned Python dependencies |

---

## ❓ Frequently asked questions

**Which mode should I use?** The GUI for everyday analysis. Agentic mode when you just want
the outputs from a chat request. The CLI for automation.

**How do I analyze a second batch of logs in the GUI?** After *Run Pipeline* finishes, click
**↺ Analyze Another Log** (or press `Ctrl+N`). Tab 1 returns to *Boot / MRC logs* with an
empty file list.

**My platform is not in the Project list.** Add it to `projects.json` and restart the GUI.

**Where are the Mode Register and RTT_WR values?** In the **Mode Registers** and **ODT** tabs
of `RMT_Report.html`, and in `rmt_metadata.json` under `platform_infos[*].mr_odt`.

**The JMP Charts tab is missing.** It appears only when JMP produced PNGs (Stage 3/4 or
`--jmp-from-csv` without `--jmp-jsl-only`).

**Can I generate charts without PowerPoint?** Yes. Stage 3 produces the PNGs; the HTML
report embeds them.

**What if I deny chart approval (interactive CLI)?** The script exits with code 3. The CSV
and Excel are still written.

---

## 🐛 Troubleshooting

### Shortcuts / icon

| Symptom | Fix |
|---------|-----|
| No desktop icon after setup | Run `powershell -ExecutionPolicy Bypass -File scripts\create_shortcuts.ps1 -Desktop`. The Desktop may be redirected to OneDrive; the script uses the real location. |
| Taskbar pin was not created | Expected on Windows 10 1809+ / 11. Right-click *Start ▸ All apps ▸ MarginIQ* ▸ *Pin to taskbar*. |
| Taskbar shows a Python icon | Run setup again (or `create_shortcuts.ps1 -StartMenu`) so the shortcut carries the MarginIQ AppUserModelID, then re-pin. |
| Shortcut does nothing | `.venv` is missing or broken. Run `setup.bat --recreate`. |
| Old / blank icon | Windows icon cache. Sign out and back in, or run `ie4uinit.exe -show`. |

### Cloning behind a broken proxy

Symptom:
`fatal: unable to access 'https://github.com/...': Unsupported proxy syntax in 'proxy-server:port': Port number was not a decimal number between 0 and 65535`

A proxy on your machine is still set to the template placeholder `proxy-server:port`,
so git (and pip) cannot connect. Get the tool first, then fix the setting at its source.

**1. Clone and set up, passing a valid proxy each time:**

```powershell
git -c http.proxy=http://proxy-chain.intel.com:912 clone https://github.com/kumarss-intel/RMT-Margin-Analysis-Tool.git
cd RMT-Margin-Analysis-Tool
.\setup.bat --proxy http://proxy-chain.intel.com:912
```

**2. Save the proxy for this repository** so that later `git pull` commands work.
The `-c` option above applied only to the clone; the saved value takes precedence
over both the broken git setting and the environment variables:

```powershell
git config http.proxy http://proxy-chain.intel.com:912
```

**3. Recommended: find the placeholder and fix it at its source**, because other tools on the
machine will keep failing until you do:

```powershell
git config --show-origin --get-all http.proxy
git config --show-origin --get-all https.proxy
foreach($s in 'User','Machine'){ foreach($k in 'HTTP_PROXY','HTTPS_PROXY','ALL_PROXY','PIP_PROXY'){
  $v=[Environment]::GetEnvironmentVariable($k,$s); if($v){ "$s  $k = $v" } } }
```

| Placeholder shown in | Fix |
|---|---|
| git config, `global` file | `git config --global http.proxy http://proxy-chain.intel.com:912` (or `--unset http.proxy`) |
| git config, `system` file | Same with `--system`, from an administrator PowerShell |
| `User` environment variable | `[Environment]::SetEnvironmentVariable('HTTPS_PROXY','http://proxy-chain.intel.com:912','User')` (repeat for `HTTP_PROXY`) |
| `Machine` environment variable | Same with `'Machine'` from an administrator PowerShell, or ask IT if it was pushed by policy |

Open a new terminal after changing an environment variable. Use your site's proxy
(for example `proxy-iind.intel.com:912`) if `proxy-chain` does not work for you.

### Setup fails with a proxy error

Symptom: `ERROR: Could not install packages due to an OSError: Failed to parse: http://proxy-server:port`

A proxy on your machine is set to a template placeholder (the port must be a number).
It usually comes from a `proxy =` line in a pip config file such as `%APPDATA%\pip\pip.ini`,
or from the `PIP_PROXY` / `HTTP_PROXY` / `HTTPS_PROXY` environment variables. `setup.bat`
now prints the exact source. Fix it in either of these ways:

```powershell
# Quick: use a valid proxy for this setup run only
.\setup.bat --proxy http://proxy-chain.intel.com:912

# Permanent: find and correct the pip config file
.venv\Scripts\python.exe -m pip config debug
.venv\Scripts\python.exe -m pip config set global.proxy http://proxy-chain.intel.com:912
```

For an environment variable, fix or delete it under Windows Settings >
"Edit environment variables for your account", then open a new terminal.

### No data extracted

- Verify the input files contain `START_RMT` text and rows like `Mc0.C0.R0: -32.0 32.0 ...`.
- Check `--pattern` / *File Mask* matches your extension (`*.log` vs `*.txt`).
- For another platform, pick the right **Project**, or add its header aliases to `projects.json`.

### Missing CSV / Excel / PPT

```powershell
.\setup.bat --recreate
.venv\Scripts\python.exe -c "import openpyxl, pptx; print('OK')"
```

### JMP charts not generating

```powershell
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"
.venv\Scripts\python.exe rmt_log_pipeline.py ... --jmp-jsl-only   # inspect the JSL
```

MarginIQ closes any running JMP instance before a chart run (JMP is single-instance).
Save your JMP work first.

### Reporting a problem

Attach the GUI debug log (*Help ▸ Open Debug Log Folder*), the command preview from Tab 3,
and the output folder's `rmt_metadata.json`.

---

## 📜 Version

**MarginIQ 2.0.0** · Intel CCG CVE · Tested with Python 3.14, Windows 11 25H2, PowerShell 5.1,
JMP Pro 17. Part of the Intel Memory Reference Code (MRC) validation environment.
