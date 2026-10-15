# MarginIQ

### Intel CCG CVE Memory RMT Margin Analysis Tool

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
| ⌨️ **CLI** | Scripts, batch jobs, CI, other automation tools | `marginiq.cmd` (or `.venv\Scripts\python.exe rmt_log_pipeline.py`) — see [CLI_AUTOMATION.md](CLI_AUTOMATION.md) |

---

## 📑 Contents

1. [What's new in 2.0](#-whats-new-in-20) — incl. [Wildcat Lake support & automation](#wildcat-lake-wcl-support--automation)
2. [Quick start](#-quick-start)
3. [Installation & setup](#-installation--setup) — incl. [desktop / taskbar shortcuts](#desktop-taskbar-and-start-menu-shortcuts), [staying up to date](#staying-up-to-date) and the step-by-step [INSTALL.md](INSTALL.md)
4. [Using the GUI](#-using-the-gui)
5. [Agentic mode (GitHub Copilot)](#-agentic-mode-github-copilot)
6. [Command line](#-command-line) — full automation guide: [CLI_AUTOMATION.md](CLI_AUTOMATION.md)
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
| Branding | Product name **MarginIQ** with the subtitle *Intel CCG CVE Memory RMT Margin Analysis Tool* in the GUI, HTML report and PPT. Version via `--version` or Help ▸ About. |
| Icon & shortcuts | Unique MarginIQ icon ([`assets/marginiq.ico`](assets/marginiq.ico)) on the GUI window, taskbar button and header. `setup.bat` creates a **Desktop shortcut** and asks whether to **pin to the taskbar** and add a **Start menu / Quick Launch** entry. `Launch_RMT_GUI.bat` gets an icon'd twin, `MarginIQ.lnk`. |
| Project-agnostic | **Project** dropdown on Tab 1 and `--project NVL\|WCL` on the CLI, driven by `projects.json` (name, code, START_RMT header aliases, axis file). Add a project without code changes. WCL-style `RxVref` headers and per-byte `Mc0.C0.B0.R0` rows are parsed; WCL DDR5 / LPDDR5 RVP variants are detected from the log and WCL DTR logs use their own `dtr` rules. |
| Workflow | **↺ Analyze Another Log** replaces the disabled *Next* button on Tab 3 after a successful run; also *File ▸ New Analysis* (`Ctrl+N`). It resets Tab 1 to the launch state (*Boot / MRC logs* source, empty file list, default stage/profile, locked tabs, default axis values). Project, JMP path and PPT template are kept. Tab 1 is greyed out from *Run Pipeline* until then (unlocked again if the run fails or is stopped). |
| Tab 2 layout | Grouped toolbar: **Axis Values** (apply / recalculate), **Axis Presets** (Save JSON, Load JSON, Load .jrp) and **JMP Graph Builder** (opens a live, editable Graph Builder with the current axis values). |
| Default ±Ref | Configurable **+ Ref / − Ref** in the Tab 2 *Defaults*, saved as `defaults.ref_line_plus/minus`. Used whenever a parameter's own Ref is empty: JMP ref lines (now also on DTR charts), PASS/WARN/FAIL status and red raw-data cells. *Apply ±Ref to all* copies it to every charted parameter. |
| HTML report | New **Mode Registers** and **ODT** tabs; the **JMP Charts** tab is now a dropdown-driven side-by-side comparison workspace; **Raw Data** is now an explorer (status, heatmap, numeric filters, views, export) with a **Δ Drift Comparison** (e.g. Boot → Run RMT). See [HTML report](#-html-report). |
| Tab 3 | *Open HTML Report* button. |
| Update notice | The GUI banner and CLI line say when a newer version is on the server and ask you to `git pull`. The check runs at most once a day. See [Staying up to date](#staying-up-to-date). |

### Wildcat Lake (WCL) support & automation

| Area | Change |
|------|--------|
| WCL log formats | **LPDDR5 RVP** (`START_BYTE_RMT`, per byte-rank `Mc0.C0.B0.R0`, header `RxVref`) and **DDR5 RVP** (`START_RMT` per rank). Validated with LP5 (7467 MT/s), DDR5 (5600 MT/s) and DDR5 thermal DTR logs. |
| Only real RMT data | For WCL, data is taken **only** from inside *Rank Margin Tool -- Started … SUCCEEDED*. The "Composite UI" margin checks that MRC prints after every training step are never plotted. A log whose Rank Margin Tool never ran (or printed no block) gets `[warn] … NO RMT DATA … file skipped`; if no file has data the run exits with code 2. |
| Platform variant | **WCL DDR5 RVP** / **WCL LPDDR5 RVP**, decoded from the log (`DDR5 CSODIMM` / `LPDDR5X Memory Down`, `Detected board:`), shown in the HTML header, the Platform tab and the PPT. LPDDR5X memory-down parts are decoded (type, ranks, width, size). |
| WCL thermal (DTR) | `--dtr` uses the first and last *Rank Margin Tool* run (`x 0xdd44` re-runs) as Boot / Run RMT. Temperatures come from the **BCRH / BHRC** tag anywhere in the file name, because WCL logs print no usable PHY temperature. |
| **Spec RMT limits** | PTL/WCL limits table (DDR5 3200–7200, LPDDR5 2400–9600 MT/s) in `projects.json`. The limit for the memory type + frequency **read from each log** becomes the ±Ref line in JMP / HTML / PPT, and a **Spec RMT Limits** PASS/FAIL table is added (HTML Overview card, PPT slide, `rmt_metadata.json` → `rmt_limits`). `--no-rmt-limits` or the Tab 2 checkbox switches it off. See [WCL spec RMT limits](#wcl-spec-rmt-limits-rmt_limits). |
| Axis auto-fit | The GUI recalculates Min / Max / Inc from **each new log's data** as soon as Stage 1 finishes (no stale ranges from the previous log). WCL fills ±Ref with the spec limits. The CLI prints `[warn] JMP axis hides data …` when an axis JSON would clip points. |
| Log capture gaps | A GREEN MRC task whose end line was lost by the serial capture is shown as **INCOMPLETE** (not RUNNING), with a log-gap alert in the Training Steps tab and PPT, and a `[warn]` line. |
| Command line / automation | **`marginiq.cmd`** runs the pipeline from any folder with the repository's `.venv`, UTF-8 output and the exit code passed through. **[CLI_AUTOMATION.md](CLI_AUTOMATION.md)** has the complete command-prompt guide: recipes, GUI↔CLI mapping, exit codes, output / JSON contract, console messages and integration examples. The interactive runner (`rmt_pipeline_runner.py`) now asks for the Project, Thermal profile and spec limits; `--cmds` prints cmd.exe commands. |

NVL (NVL-S / NVL-HX, DDR5) parsing and outputs are unchanged by all of the above.

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
- **CLI:** run the pipeline directly from Command Prompt or PowerShell:
  ```bat
  marginiq.cmd --project WCL --input "C:\path\to\logs" --pattern "*.log" --outdir "C:\path\to\out" --no-ask-chart-approval
  ```
  Everything for scripting and automation tools is in [CLI_AUTOMATION.md](CLI_AUTOMATION.md).

---

## 🔧 Installation & setup

### Prerequisites

> **New user?** Follow the step-by-step **[Installation Guide (INSTALL.md)](INSTALL.md)**. It covers
> checking Python and adding it to `PATH`, repairing a broken Python install, installing JMP Pro from
> Company Portal, and fixing the JMP Pro licence prompt.

- Windows 10 or 11
- Python 3.10+ (3.14 recommended), installed with the **tcl/tk and IDLE** option so `tkinter` is available for the GUI. Ticking **Add python.exe to PATH** is recommended.
- Optional: SAS JMP Pro (licensed Windows install) for JMP charts, the JMP PPT and Graph Builder.
  Install it from **Company Portal** (search *JMP*, pick **JMP Pro**). JMP Pro is **not available
  on lab machines**, so install MarginIQ on your laptop for the JMP stages.
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

1. Locates a Python 3.10+ interpreter, even one that is not on `PATH`. It searches `C:\Program Files\Python3xx`, the per-user `%LOCALAPPDATA%\Programs\Python\Python3xx`, Python Install Manager runtimes, then the `py` launcher, then `PATH`.
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
| Pipeline CLI | `marginiq.cmd --help` (any folder) or `.venv\Scripts\python.exe rmt_log_pipeline.py --help` — [CLI_AUTOMATION.md](CLI_AUTOMATION.md) |
| Interactive CLI menu | `.venv\Scripts\python.exe rmt_pipeline_runner.py` (`--quick`, `--cmds`, `--guide`) |
| Guides in the browser | GUI *Help* menu / `F1`, or `marginiq.cmd --guide [readme\|quickstart\|cli\|install\|pipeline\|projects]` — see [Help menu](#help-menu) |

Prefer an activated shell? Run `.venv\Scripts\activate.bat`; plain `python` then resolves to the virtual environment.

### Staying up to date

MarginIQ tells you when the repository has a newer version on the server. The check is in
[`rmt_update.py`](rmt_update.py):

| Where | What you see |
|-------|--------------|
| GUI | An amber banner under the header (*"A newer MarginIQ version is available on origin/main…"*) on every launch until you update, plus a one-time pop-up for each new server version. *Help ▸ Check for Updates…* checks on demand. |
| CLI / agent | One line after the version banner: `[MarginIQ update] A newer version is available …`. The exit code is never affected. |

To update, close MarginIQ and run `git pull` in the tool folder. If `requirements.txt` changed,
then run `setup.bat --no-prompt`.

**Light on the server.** The check contacts the git server **at most once every 24 hours** per
clone; every other launch reuses the cached answer in `.git\marginiq-update-check.json`. A
failed check (offline, proxy, no access) is cached too, so it is not retried on each launch. The
network call is one `git ls-remote` of your upstream branch. It downloads nothing and never
changes your files or branches. It never prompts for credentials, times out after 10 s, and runs
in the background in the GUI, so it cannot slow the GUI or hang a run.

| Setting | Effect |
|---------|--------|
| `MARGINIQ_UPDATE_CHECK_HOURS=<h>` | Change the interval (default `24`) |
| `MARGINIQ_NO_UPDATE_CHECK=1` | Turn off automatic checks (*Help ▸ Check for Updates…* still works) |
| `--no-update-check` | Skip the check for one CLI run |
| `.venv\Scripts\python.exe rmt_update.py --force` | Check now from a terminal |

### About JMP (excluded from the virtual environment)

JMP Pro is licensed desktop software and can never be installed by pip. Setup only detects it, via the
`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\jmp.exe` registry entry, the standard
`C:\Program Files\SAS\JMPPRO\<version>` folders, and `PATH`.

Without JMP, log parsing, CSV, Excel, the HTML report and the native PPT all still work, and
`--jmp-jsl-only` produces the JSL script to run later on a JMP-equipped machine.

**Getting JMP Pro:** lab machines do not have it, so use your laptop. Install **JMP Pro** from
**Company Portal** (search *JMP*) and start it once by hand. If JMP Pro 17 then reports a licence
problem, get the licence from the intranet page **"JMP Statistical Software – JMP Top Level"**, or
raise an IT ticket. See [INSTALL.md › JMP Pro](INSTALL.md#2-jmp-pro).

---

## 🖥️ Using the GUI

The window has a **File** menu (*New Analysis* `Ctrl+N`, *Open Output Folder*, *Exit*), a **Help**
menu, and three tabs driven by the *Previous / Next* bar.

### Help menu

The guides open in your default browser through MarginIQ's built-in viewer
([`rmt_help.py`](rmt_help.py)). It renders the local Markdown files as styled pages with a guide
sidebar, an *On this page* contents panel, in-page search (`Ctrl+F` or `/`, `Enter` for the
next match), *Copy* buttons on every command block, and working links between the guides. It
uses the standard library only and works offline.

| Item | Opens / does |
|------|--------------|
| **User Guide** (`F1`) · **Quick Start** · **Command Line & Automation** · **Installation Guide** · **Pipeline Reference** | `README.md`, `RMT_QUICKSTART.md`, `CLI_AUTOMATION.md`, `INSTALL.md`, `RMT_LOG_PIPELINE_README.md` |
| **What's New** · **Troubleshooting & FAQ** | The matching section of the User Guide |
| **Supported Projects & Spec Limits** | A page generated from `projects.json`: each project's parsing rules and the WCL DDR5 / LPDDR5 spec RMT limits table (ticks and ps / mV) |
| **Open Debug Log Folder** · **Open Tool Folder** | Explorer at `logs\` / the repository folder |
| **Copy Diagnostics for Support** | Copies version, git branch / commit, Python, Windows, project / profile / stage, JMP path, output folder, debug-log path and the current command preview to the clipboard |
| **Check for Updates…** · **About MarginIQ** | Update check / version information |

The same viewer works outside the GUI:

| Mode | How |
|------|-----|
| Command Prompt / PowerShell | `marginiq.cmd --guide [readme\|quickstart\|cli\|install\|pipeline\|projects]` (or `.venv\Scripts\python.exe rmt_help.py <page>`) |
| Interactive CLI menu | `rmt_pipeline_runner.py`, option **8**, or `rmt_pipeline_runner.py --guide <page>` |
| Agentic (Copilot) | Ask *"open the CLI guide"* or *"show me the WCL spec limits table"*: the agent opens the page with the same viewer. For how-to questions it answers from the guides directly. |

### Tab 1 · Input & Workflow

| Section | What to do |
|---------|------------|
| **Project & Experiment Profile** | **Project** (Nova Lake, Wildcat Lake, …) labels the reports and selects that platform's parsing rules (header aliases, RMT block markers, DTR flow, platform variants, spec limits — see [Projects](#projects--projectsjson)). **Profile**: *Non-Thermal* or *Thermal Experiment (DTR)* for BCRH / BHRC logs (adds `--dtr`). |
| **Data Source** | *Boot / MRC logs* (default), *Existing CSV*, *Existing Excel* or *PNG charts → PPT only*. |
| **Files / Folder** | *Add Files…* / *Add Folder…* (filtered by *File Mask*, e.g. `*.log;*.txt`). |
| **Multiple-File Handling** | *Concatenate* into one report (overlay by source) or process each file *Separately* (one subfolder each). |
| **Output**, **PPT Template**, **Workflow Stage**, **Tools & Options** | Output folder (auto-filled), file names, optional `.pptx` template, stage 1–4, JMP executable, JSL-only. |

**Generate CSV & Continue →** runs Stage 1 and unlocks Tabs 2–3, which then reuse that CSV.

### Tab 2 · Parameters & Axis

| Group | Controls |
|-------|----------|
| **Axis Values** | *Apply axis settings*; **↻ Recalculate from data** (Min / Max / Inc from the logs; ±Ref preserved). Runs automatically, without popups, as soon as Stage 1 produces the CSV (or CSVs are loaded), so every new log is charted on its own data range; manual edits made afterwards are kept. **Spec RMT limits as ±Ref (per frequency)** (default on): for projects with `rmt_limits` (WCL), Recalculate fills ±Ref with the spec limit for the memory type and frequency read from the log; untick to use your own ±Ref (`--no-rmt-limits`). These spec values are **display-only**: they are never saved to `jmp_axis_settings.json` or axis presets, and your own ±Ref returns when you switch project (e.g. to NVL, which always uses the ±Ref from the axis settings, default ±10), untick the box or reset. |
| **Axis Presets** | **Save JSON…**, **Load JSON…**, **Load .jrp…** (clone scales from a JMP report). |
| **JMP Graph Builder (interactive)** | Choose *All charted parameters* or one parameter, then **▶ Open**. JMP opens live Graph Builder windows using the current axis values. Fine-tune them, save as `.jrp` in JMP, and bring the scales back with *Load .jrp…*. |
| **Defaults** | Inc, minor ticks, grids, **+ Ref / − Ref** (default ±10), *Apply ±Ref to all*, ref-line colours. |
| **Per-Parameter Y-Axis** | *Chart* checkbox, +/− Min / Max / Inc / Ref per parameter, and a live eye-preview. An empty Ref uses the default ±Ref. |

### Tab 3 · Preview & Run

An editable command preview (what you see is what runs), **▶ Run Pipeline**, **■ Stop**,
*Open Output Folder*, *Open HTML Report* and a live output log. As soon as *Run Pipeline* starts,
**Tab 1 · Input & Workflow is greyed out** (and *← Previous* on Tab 2 is disabled), so the
finished analysis cannot be re-parsed with half-changed inputs. Tabs 2 and 3 stay available.
After a successful run the bottom-right button becomes **↺ Analyze Another Log**: the only way
back to Tab 1. It clears the session back to the launch state, so the next log set starts on
Tab 1 with *Boot / MRC logs* selected. If the run fails or is stopped, Tab 1 is unlocked again
so you can correct the inputs (e.g. the JMP path).

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

From Command Prompt, PowerShell or any automation tool, use **`marginiq.cmd`** in the repository
root. It runs from any folder with the repository's `.venv`, writes UTF-8 output and returns the
pipeline's exit code. The flags are the same as `rmt_log_pipeline.py`.

```bat
marginiq.cmd --project WCL --input "C:\path\to\logs" --pattern "*.log" ^
  --outdir "C:\path\to\out" --no-ask-chart-approval --no-update-check ^
  --generate-jmp-charts --jmp-exe "C:\Program Files\SAS\JMPPRO\17\jmp.exe" ^
  --jmp-axis-config ".\jmp_axis_settings.json"
```

| Key flags | Purpose |
|-----------|---------|
| `--project NVL\|WCL\|"Wildcat Lake"` | Platform from `projects.json` (default: registry default, `NVL`). Selects that platform's parsing rules |
| `--dtr` | Thermal Experiment (BCRH / BHRC): Boot-temp and Run-temp RMT per log (NVL and WCL flows) |
| `--chart-fields a,b,c` | Parameters to chart (omit for all) |
| `--generate-jmp-charts --jmp-exe <exe>` | JMP PNG charts and the JMP PPT |
| `--no-rmt-limits` | WCL: use the axis-config ±Ref instead of the spec RMT limits |
| `--jmp-from-csv <csv…>` / `--jmp-from-excel <xlsx>` / `--ppt-from-charts <dir>` | Reuse existing artifacts |
| `--jmp-jsl-only` | Write the JSL without launching JMP |
| `--no-ask-chart-approval` | Always pass in scripts (no prompt) |
| `--version` | Print the MarginIQ version |
| `--no-update-check` | Skip the (once-a-day) check for a newer version on the server |

**Exit codes:** 0 success (a WCL spec FAIL is still 0; read `rmt_metadata.json` → `rmt_limits.overall`)
· 1 no input files · 2 bad arguments / unknown project / no RMT data · 3 chart approval denied
· 4 `--jmp-exe` missing · 5 JMP failed.

**Automation and scripting:** [CLI_AUTOMATION.md](CLI_AUTOMATION.md) covers recipes for NVL /
WCL / DTR, the GUI ↔ CLI mapping, every flag, the output and JSON contract, console messages to
scrape, environment variables, and batch / PowerShell / Python / Task Scheduler examples.
Full flag manual: [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md). Beginner walkthrough:
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
| Overview | KPI cards, **Spec RMT Limits** card (WCL: limit per memory type + frequency, worst −/+, failing rows, PASS/FAIL), margin degradation below ±Ref, window-width chart, statistics (Cpk, guardband), interactive multi-source dashboard |
| Frequency / Parameters / RunTemp | Per-frequency, per-rank and temperature-drift charts |
| Training Steps / Platform | MRC task timeline with each task's **Post Code** (the `GREEN Post Code:` / `Post Code:` printed just before the task's *Started* line, e.g. *Read Voltage Centering* → `DD57h`; also shown for failed tasks in the PPT) and **Status** (SUCCEEDED / FAILED / RUNNING; a non-zero MRC status code is shown in the badge, e.g. `FAILED · 23h`; **INCOMPLETE** = the task's end line is missing because the serial capture dropped lines, flagged with a log-gap alert and a `[warn]` line); board, CPU and DIMM (SPD) details |
| **Mode Registers** | Final per-rank DDR5 MR values (from *SAGV Finalization*), JEDEC decodes (CL, Ron, VrefDQ/CA/CS, RTT_*, ODTL), amber = changed by training vs the JEDEC-reset value, bold = differs between ranks, filters, and a **cross-log MR diff** |
| **ODT** | DIMM ODT summary (RonUp/Dn, **RTT_WR**, NomWr, NomRd, Park, ParkDqs, CA/CS groups), CPU read ODT, per-rank ODT decoded from MR32–36, ODT latency offsets, BIOS ODT inputs, and a **cross-log ODT matrix** |
| **JMP Charts** | Comparison workspace. Every panel has a chart dropdown; set 1–3 panels per row, add or remove panels, reset. Pre-filled with every chart, so it is never blank. |
| **Raw Data** | Explorer for every parsed row:<br>• **Frozen** File / Phase / Rank columns; short file labels (`BCRH 3200 G2`, full name on hover); a **Phase** column (Boot RMT / Run RMT for DTR).<br>• Parameter-grouped **− / +** headers.<br>• **Views:** Margins / Window width / Slack vs Ref.<br>• Per-row **Status** (PASS / WARN / FAIL), **Worst** parameter and **Min margin vs Ref** = smallest (\|margin\| − \|Ref\|) in ticks over the visible parameters; negative = below Ref (WCL: below the spec limit).<br>• A **How to read** panel above the table explains the reference in use, the FAIL / WARN / PASS rule (FAIL: \|margin\| < Ref; WARN: Ref ≤ \|margin\| < band × Ref; PASS: ≥ band × Ref), the **Warn band** (2× default, 1.5×, 1.25×; it only moves the WARN/PASS boundary), a worked example from the weakest row of the report, and how to act on FAIL / WARN / PASS.<br>• **Heatmap** plus **red** (below ±Ref) and **amber** (inside a selectable warn band, default 2× Ref) cells.<br>• **Numeric filters** (`<15`, `>=20`, `10..20`, compared on \|margin\|), *Show: All / WARN+FAIL / FAIL*, search, sort, 25 / 50 / 100 / All rows.<br>• **Columns** menu (constant parameters such as RecEnDelay ±32 are hidden by default).<br>• **Copy for Excel** and **CSV** export of the filtered view. |
| **Δ Drift Comparison** (Raw Data tab) | Compare any **A → B**: Phase (Boot → Run RMT, the DTR default), file group (BCRH vs BHRC), frequency, gear or file. Gives a per-rank **Δ = \|B\| − \|A\|** table (per side or window width, red = margin lost), per-parameter mean and worst Δ, a **Top-10 margin losses** list, and **Δ CSV** export. |

---

## ⚙️ Configuration

### Projects — [`projects.json`](projects.json)

```json
"WCL": { "name": "Wildcat Lake", "code": "WCL", "memory": "DDR5 / LPDDR5",
         "param_aliases": { "RxVref": "RxDqVrefByte" },
         "rmt_fallback_markers": ["START_BYTE_RMT"],
         "require_rmt_task": true,
         "platform_variants": { "DDR5": "WCL DDR5 RVP", "LPDDR5": "WCL LPDDR5 RVP" },
         "dtr": { "layout": "rmt_task_runs",
                  "rmt_markers": ["START_RMT", "START_BYTE_RMT"],
                  "temps_from_filename": true },
         "rmt_limits": { "LPDDR5": { "frequencies": { "7467": { "ticks": { "TxVref": [25, 25], … },
                                                               "physical": { "TxVref": [63, 63], … } }, … } },
                         "DDR5":   { … } },
         "axis_config": "jmp_axis_settings.json" }
```

Add an entry to support a new platform. It appears in the GUI dropdown and is accepted by
`--project`. `param_aliases` maps a log's START_RMT header name to MarginIQ's canonical
parameter name. The other keys are optional; projects without them (e.g. NVL) keep the
built-in behaviour.

- **`rmt_fallback_markers`** lists extra block markers (WCL LPDDR5 prints `START_BYTE_RMT`)
  that are parsed **only when a log has no `START_RMT` data**, and only inside
  "Rank Margin Tool" task runs.
- **`require_rmt_task`** (default `false`): RMT data is taken **only** from inside the
  "Rank Margin Tool" MRC task (`-- Started` … `-- SUCCEEDED`). The per-training-step
  "Rank Margin Testing for Composite UI errors" blocks that MRC prints after each GREEN MRC
  task are never plotted. A log whose Rank Margin Tool task never ran, or ran without
  printing a block, gets a warning and is skipped:
  `[warn] <file>: NO RMT DATA - the 'Rank Margin Tool' MRC task never ran in this log; file skipped. 17 per-training-step Composite-UI block(s) elsewhere were ignored.`
  If no input file has RMT data, the run exits with code 2. When this is `false` (NVL), logs
  without any task markers (e.g. consolidated RMT snippets) still read `START_RMT` anywhere.
- **`platform_variants`** maps the memory family decoded from the log (`DDR5 CSODIMM` →
  DDR5, `LPDDR5X Memory Down` → LPDDR5; the `Detected board:` line is the fallback) to a
  variant label. The label appears in the HTML header and Platform tab and on the PPT title
  and Test Configuration slides.
- **`dtr`** overrides `--dtr` parsing. `layout: "rmt_task_runs"` means Boot RMT is the first
  "Rank Margin Tool" run and Run RMT is the last one (WCL re-runs it from the MRC interpreter
  with `x 0xdd44`; there is no NVL-style "Temp Drift RMT Test" / "Hit Enter").
  `rmt_markers` lists the accepted block markers. `temps_from_filename` takes the Boot / Run
  temperatures from the BCRH (0 → 90 °C) or BHRC (90 → 0 °C) tag anywhere in the file name,
  because these logs print no usable PHY temperature.

#### WCL log support status

| WCL log type | Status |
| --- | --- |
| LPDDR5 RVP, non-thermal single boot | Supported: the per-byte `START_BYTE_RMT` block inside the "Rank Margin Tool" task. Validated with `teraterm_17092026.log` (board `WCL LP5 RVP2`, 7467 MT/s). Logs without that task (e.g. `13082026_lp5_siv_minibios.log`) are flagged `NO RMT DATA` and skipped. |
| DDR5 RVP, thermal DTR (`--dtr`, BCRH / BHRC) | Supported (`dtr` rules above) |
| DDR5 RVP, non-thermal single boot | Supported. Same layout as NVL: per-rank `START_RMT` blocks, where only the block inside the "Rank Margin Tool" task is used and the per-training-step Composite-UI blocks are ignored. Validated with `WCL_DDR5_RVP\5600_g4.log` (board `WCL R DDR5 RVP3`). |
| LPDDR5 RVP, thermal DTR | Not yet seen; the `dtr` rules apply if the flow matches the DDR5 one. |

NVL (NVL-S and NVL-HX, DDR5 only) uses the built-in `START_RMT` and DTR flows unchanged.

#### WCL spec RMT limits (`rmt_limits`)

`projects.json` holds the PTL/WCL RMT limits table (A. Magrlashvili, 24 Feb 2026) for
**LPDDR5** (2400 – 9600 MT/s) and **DDR5** (3200 – 7200 MT/s), stored in RMT ticks with the
physical values (ps / mV). Mail columns map to MarginIQ parameters as RdT → `RxDqsDelay`,
WrT → `TxDqDelay`, RdV → `RxDqVrefByte`, WrV → `TxVref`, CmdT → `ClkGrpPi` and
CmdV → `CmdVref`. For each WCL log:

1. The memory family is decoded from the log (`DDR5 CSODIMM` → DDR5, `LPDDR5X Memory Down` →
   LPDDR5) and the frequency is read from the log (`Frequency=7467`). The matching row is
   used. If the frequency is not in the table, the nearest listed frequency is used and a
   `[warn]` is printed.
2. Those limits become the per-parameter **±Ref lines** in the JMP charts (non-thermal and
   DTR), the HTML report (raw-data highlighting, statistics thresholds) and the PPT. The axis
   is widened when needed so the limit line stays visible.
3. A **Spec RMT Limits** card (top of the HTML Overview tab) and PPT slide list each
   parameter's limit in ticks and physical units, the worst − / + margin, and the failing
   rows. A rank fails when either side is weaker than its limit. Failures are also printed
   as `[warn] RMT spec FAIL …`.

If one run mixes several memory/frequency combinations, the chart ±Ref lines use the
strictest limit, while the PASS/FAIL table checks every row against its own limit.
`--no-rmt-limits` (or unticking the GUI checkbox) falls back to the axis-config ±Ref.
Projects without `rmt_limits` (NVL) are not affected.

### Axis settings and default ±Ref — [`jmp_axis_settings.json`](jmp_axis_settings.json)

`defaults` holds `inc`, `minor_ticks`, grids, ref-line colours and **`ref_line_plus` /
`ref_line_minus`** (the default ±Ref, ±10). Each parameter has `plus` / `minus` blocks with
`min`, `max`, `inc` and `ref_line`. A missing `ref_line` falls back to the default ±Ref for JMP
ref lines, PASS/WARN/FAIL and raw-data highlighting. Edit the file in Tab 2 or by hand; no
code change is needed.

### Branding / version — [`rmt_project.py`](rmt_project.py)

`TOOL_NAME`, `TOOL_SUBTITLE` and `TOOL_VERSION` are defined once and used by the GUI, CLI,
HTML and PPT.

### DRAM die / process-node data — [`dram_nodes.json`](dram_nodes.json)

The HTML **Platform** tab and the `RMT_Summary.pptx` deck show a **DRAM Die & Process Node**
card for each unique DIMM part number. It is built by [`rmt_dram.py`](rmt_dram.py) from two
sources, and every value carries a basis label:

- **Log (SPD)**: decoded from the run's own SPD dump (JESD400-5). This covers DRAM stepping
  (byte 554), die density and package, x-width, VDD/VDDQ/VPP, fastest JEDEC bin and
  CL-tRCD-tRP, supported CLs, tRFC1/2/sb, and module organisation.
- **`dram_nodes.json`**: vendor part-number nomenclature, die → node mapping and node metrics,
  each tagged *Published*, *Industry consensus*, *JEDEC*, *Engineering Estimate* or *Not
  publicly disclosed*, with source URLs.

The source URLs appear in the HTML as `[n]` links and in the PPT slide notes. To cover a new
part, add a decoder regex, a `dies` entry and a `node_families` entry. Never add an
unlabelled value, and never cite NDA datasheets.

---

## 📁 Output structure

```
output_dir/
├── RMT_Combined_Similar.csv        # standard columns
├── RMT_Combined_Extended.csv       # + SourceFile / Frequency / Gear / BlockIndex
├── csv_by_file/<log>_RMT.csv       # one CSV per source log
├── RMT_Extraction.xlsx             # All_RMT + per-file + per-frequency sheets
├── RMT_Report.html                 # interactive report (MR / ODT / JMP comparison tabs)
├── RMT_Summary.pptx                # native PowerPoint summary (WCL: + Spec RMT Limits slide)
├── rmt_metadata.json               # training steps, platform / variant, MR / ODT data, WCL rmt_limits PASS/FAIL (reused by --jmp-from-csv)
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
| [INSTALL.md](INSTALL.md) | Step-by-step installation: Python / PATH, JMP Pro from Company Portal, licence, setup errors |
| [RMT_QUICKSTART.md](RMT_QUICKSTART.md) | 5-minute beginner guide (GUI, agent, CLI) |
| [CLI_AUTOMATION.md](CLI_AUTOMATION.md) | Command-prompt & automation guide: recipes, GUI↔CLI mapping, flags, exit codes, output / JSON contract, integration examples |
| [RMT_LOG_PIPELINE_README.md](RMT_LOG_PIPELINE_README.md) | Complete CLI reference manual |
| [marginiq.cmd](marginiq.cmd) | Command-prompt entry point (repo `.venv`, UTF-8 output, exit code passed through) |
| [rmt_pipeline_runner.py](rmt_pipeline_runner.py) | Interactive command-line menu (`--quick`, `--cmds`) for people |
| [rmt_help.py](rmt_help.py) | Built-in Help viewer: renders the guides and the projects / spec-limits page as HTML (GUI *Help* menu, `F1`) |
| [.github/skills/rmt-margin-plots/SKILL.md](.github/skills/rmt-margin-plots/SKILL.md) | Copilot agent skill (agentic mode) |
| [.github/copilot-instructions.md](.github/copilot-instructions.md) | Repository rules for Copilot |
| [setup.bat](setup.bat) | Environment setup and shortcuts |
| [scripts/create_shortcuts.ps1](scripts/create_shortcuts.ps1) | Create / remove the MarginIQ Desktop / Start / Quick Launch / taskbar shortcuts |
| [scripts/make_icon.py](scripts/make_icon.py) | Regenerates `assets/marginiq.ico` / `.png` |
| [scripts/check_proxy.py](scripts/check_proxy.py) | Proxy pre-flight check run by `setup.bat` |
| [rmt_update.py](rmt_update.py) | Rate-limited "newer version on the server?" check used by the GUI and CLI |
| [projects.json](projects.json) | Project registry |
| [dram_nodes.json](dram_nodes.json) | DRAM part-number / die / process-node knowledge base (Platform tab + PPT) |
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

**WCL: why does a log give "NO RMT DATA"?** MarginIQ takes WCL margins only from inside the
*Rank Margin Tool* MRC task. If that task never ran (e.g. MRC stopped in *EV Debug functions*
or an endless run) or printed no `START_RMT` / `START_BYTE_RMT` block, the log is skipped.
Plotting the per-training-step "Composite UI" checks as RMT would be misleading. Re-run the
board with RMT enabled.

**WCL: which spec limit is used?** The one for the memory type decoded from the log (DDR5 or
LPDDR5) and the frequency read from the log (`Frequency=…`). A frequency that is not in the
table uses the nearest listed one, with a `[warn]`. See
[WCL spec RMT limits](#wcl-spec-rmt-limits-rmt_limits).

**Why is a training step INCOMPLETE?** Its SUCCEEDED / FAILED line is missing from the log
because the serial capture (e.g. TeraTerm) dropped lines. The RMT margins are not affected.
Recapture the log if you need the full timeline.

**How do I drive MarginIQ from another tool?** Call `marginiq.cmd` with the same flags, check
the exit code, then read `rmt_metadata.json`. See [CLI_AUTOMATION.md](CLI_AUTOMATION.md).

---

## 🐛 Troubleshooting

### Installation (Python / JMP Pro)

| Symptom | Fix |
|---------|-----|
| `python` works in one terminal but setup or other terminals cannot find it | Python is not on `PATH`, or only the Microsoft Store alias is. See [INSTALL.md › Python](INSTALL.md#1-python). |
| Python sits in an unexpected folder, `venv` creation fails, or `encodings` / `tkinter` errors appear | Reinstall Python, then run `setup.bat --recreate`. See [INSTALL.md › Repair](INSTALL.md#1d-repair-a-broken-python-install). |
| `No JMP installation was detected` on a lab machine | Expected: lab machines have no JMP Pro. Run stages 1–2 there, or use your laptop. |
| JMP Pro 17 licence error after the Company Portal install | Use the intranet page *"JMP Statistical Software – JMP Top Level"*, or raise an IT ticket. See [INSTALL.md › Licence](INSTALL.md#2b-jmp-pro-says-the-licence-is-missing-or-invalid). |

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
- **WCL:** look for `[warn] <file>: NO RMT DATA - …` in the run log. The Rank Margin Tool task
  never ran, or printed no block between *Started* and *SUCCEEDED* (see the FAQ).

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

Use **Help ▸ Copy Diagnostics for Support** and paste the result into your message. Attach
the GUI debug log (*Help ▸ Open Debug Log Folder*) and the output folder's `rmt_metadata.json`.

---

## 📜 Version

**MarginIQ 2.0.0** · Intel CCG CVE · Tested with Python 3.14, Windows 11 25H2, PowerShell 5.1,
JMP Pro 17. Part of the Intel Memory Reference Code (MRC) validation environment.
