# MarginIQ — Installation Guide

*Intel CCG CVE Memory RMT Margin Analysis Tool*

📖 **See Also:** [Main README](README.md) · [Quick Start](RMT_QUICKSTART.md) · [Full CLI reference](RMT_LOG_PIPELINE_README.md)

This guide takes a new user from a fresh machine to a working MarginIQ. It covers the problems
users have hit when installing: Python not on `PATH`, a broken Python install, JMP Pro not being
available on lab machines, and the JMP Pro licence prompt.

---

## Installation checklist

| # | Step | Required? | Section |
|---|------|-----------|---------|
| 0 | Choose the machine (laptop vs lab machine) | Yes | [Choose the machine](#0-choose-the-machine) |
| 1 | Python 3.10+ with `tkinter` | Yes | [Python](#1-python) |
| 2 | JMP Pro, installed **and licensed** | Only for JMP charts, the JMP PPT and Graph Builder | [JMP Pro](#2-jmp-pro) |
| 3 | Clone the repository | Yes | [Get the code](#3-get-the-code) |
| 4 | Run `setup.bat` | Yes | [Run setup](#4-run-setupbat) |
| 5 | Verify the installation | Recommended | [Verify](#5-verify-the-installation) |

---

## 0. Choose the machine

JMP Pro is **not available on lab machines**, so plan where you will run each stage:

| Machine | What works |
|---------|------------|
| **Your laptop** (with JMP Pro) | Everything: CSV, Excel, HTML report, native PPT, JMP charts, JMP PPT, live Graph Builder |
| **Lab machine** (no JMP Pro) | CSV, Excel, HTML report and native PPT (stages 1–2). `--jmp-jsl-only` writes the JSL script so you can render the charts later on a JMP machine. |

The usual setup: install MarginIQ **and** JMP Pro on your laptop, copy the logs from the lab
machine, and run the full pipeline on the laptop.

---

## 1. Python

MarginIQ needs **Python 3.10 or newer** (3.14 recommended; 3.11–3.13 also work) with the
**tcl/tk** component, which provides `tkinter` for the GUI.

### 1a. Check what you already have — before installing or upgrading

Open a **new** Command Prompt or PowerShell and run:

```powershell
py -0p                     # every Python the py launcher knows about, with its path
where.exe python           # every python.exe on PATH, in search order
python -c "import sys, tkinter; print(sys.version); print(sys.executable)"
```

| Result | Meaning | Action |
|--------|---------|--------|
| A 3.10+ version is printed, no error | Python is fine | Go to [JMP Pro](#2-jmp-pro) |
| `where.exe python` lists only `...\AppData\Local\Microsoft\WindowsApps\python.exe` | That is the Microsoft Store **alias stub**, not a real Python | [Install Python](#1b-install-python) |
| `'python' is not recognized` but `py -0p` lists a version | Python is installed but **not on PATH** | [Add Python to PATH](#1c-python-works-in-one-place-but-is-not-on-path) (or just run `setup.bat`, which also finds it) |
| `ModuleNotFoundError: No module named 'tkinter'` | Installed without tcl/tk | Re-run the installer ▸ **Modify** ▸ tick **tcl/tk and IDLE** |
| `No module named 'encodings'`, a crash, or the path shown is not where you installed Python | The install is **broken** | [Repair a broken install](#1d-repair-a-broken-python-install) |
| Version is 3.9 or older | Too old | [Install Python](#1b-install-python); older versions can stay installed side by side |

> **Check before you upgrade.** Several Python versions can be installed side by side, and
> `setup.bat` picks the newest supported one. You do not have to uninstall an older Python. Run
> the checks above first, so you know which Python and which install folder you are working
> with.

### 1b. Install Python

1. Download the Windows installer from <https://www.python.org/downloads/windows/>
   (3.14 recommended).
2. On the first installer page, tick **Add python.exe to PATH**.
3. Choose **Customize installation** and make sure **tcl/tk and IDLE** and **py launcher** are
   ticked.
4. Optional: tick **Install Python for all users** to install to `C:\Program Files\Python3xx`.
   Without it, Python goes to `%LOCALAPPDATA%\Programs\Python\Python3xx`. `setup.bat` finds
   either location.
5. Finish the install, **close every open terminal**, open a new one and repeat the
   [checks](#1a-check-what-you-already-have--before-installing-or-upgrading).

### 1c. Python works in one place but is not on PATH

Sometimes `python` works in one terminal (or through `py`) but not in others, or the
*System variables ▸ Path* entry is missing. To fix it:

- **Simplest:** re-run the Python installer ▸ **Modify** ▸ *Next* ▸ tick
  **Add Python to environment variables** ▸ *Install*.
- **Manually:** Start ▸ *"Edit environment variables for your account"* ▸ **Path** ▸ *Edit* ▸
  add **both** folders, for example:
  ```
  C:\Program Files\Python314
  C:\Program Files\Python314\Scripts
  ```
  Use your real install folder (`py -0p` shows it). If Python was installed for all users, an
  administrator can add the same folders under **System variables ▸ Path** instead.
- Open a **new** terminal afterwards. Running terminals keep the old `PATH`.

`setup.bat` does not need Python on `PATH`. It also searches `C:\Program Files\Python3xx`,
`%LOCALAPPDATA%\Programs\Python\Python3xx`, `%LOCALAPPDATA%\Python\pythoncore-3.x-64`
(Python Install Manager) and the `py` launcher. Fix `PATH` anyway, so that `python` and `pip`
work in your own terminals.

### 1d. Repair a broken Python install

If the PATH looks right but Python still misbehaves, or it sits in an unexpected folder, a
clean reinstall is usually the fastest fix (this resolved the issue for earlier users):

1. *Settings ▸ Apps ▸ Installed apps* ▸ uninstall the broken **Python 3.x** (and its
   *Python Launcher* if it is also broken).
2. Delete the leftover install folder if it still exists (for example `C:\Program Files\Python314`).
3. Remove stale entries for that folder from **Path** (user and system variables).
4. [Install Python](#1b-install-python) again.
5. Rebuild the MarginIQ environment, because `.venv` is tied to the Python that created it:
   ```powershell
   .\setup.bat --recreate
   ```

> Also run `setup.bat --recreate` after you **upgrade or change** the Python version.

---

## 2. JMP Pro

JMP Pro is licensed desktop software. `setup.bat` never installs it; it only detects it. You need
it for **stage 3/4** (JMP chart PNGs, the JMP PowerPoint, the *JMP Charts* tab of the HTML
report) and for the live **Graph Builder** on GUI Tab 2. Everything else works without it.

### 2a. Install from Company Portal

1. On your **laptop** (not a lab machine), open **Company Portal** from the Start menu.
2. Search for **`JMP`**.
3. Install the **JMP Pro** (licensed) entry. MarginIQ is tested with **JMP Pro 17**.
4. When the install completes, the executable is normally at:
   ```
   C:\Program Files\SAS\JMPPRO\17\jmp.exe
   ```
5. **Start JMP Pro once by hand** and make sure it opens without a licence prompt. Close it
   before you run MarginIQ; MarginIQ closes any running JMP before a chart run, so save your
   work first.

### 2b. JMP Pro says the licence is missing or invalid

If JMP Pro 17 reports a licence problem straight after the Company Portal install:

1. Search the Intel intranet for the SharePoint page
   **"JMP Statistical Software – JMP Top Level"** (`JMP Top Level.aspx`). It publishes the
   current JMP licence and the steps to apply it.
2. Follow those steps, then restart JMP Pro and confirm it opens cleanly.
3. If the licence still fails, raise an **IT ticket** for *JMP Pro licence activation* and
   include the exact error text.

### 2c. Confirm MarginIQ can find JMP

```powershell
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"
```

Re-run `setup.bat` (step 5/6 prints `Found JMP: ...`). If JMP is installed somewhere else, set
the **JMP executable** field on GUI Tab 1 or pass `--jmp-exe "<path>"` on the CLI.

---

## 3. Get the code

```powershell
git clone https://github.com/kumarss-intel/RMT-Margin-Analysis-Tool.git
cd RMT-Margin-Analysis-Tool
```

You need collaborator access to this private repository (ask the owner). If the clone fails with
`Unsupported proxy syntax`, see
[Cloning behind a broken proxy](README.md#cloning-behind-a-broken-proxy).

---

## 4. Run `setup.bat`

Double-click `setup.bat`, or from the repository folder run:

```powershell
.\setup.bat
```

It locates Python, creates `.venv`, installs the pinned packages, detects JMP and creates the
**MarginIQ** desktop icon. All options and the shortcut questions are explained in
[README › One-time setup](README.md#one-time-setup).

| Setup message | Fix |
|---------------|-----|
| `No working Python 3.10 or newer interpreter was found` | [Python](#1-python): install it, or repair a broken install |
| `This Python has no tkinter module` | Python installer ▸ **Modify** ▸ tick **tcl/tk and IDLE** |
| `Failed to create the virtual environment` | [Repair the Python install](#1d-repair-a-broken-python-install), then `setup.bat --recreate` |
| `INVALID` proxy / `Failed to parse: http://proxy-server:port` | `.\setup.bat --proxy http://proxy-chain.intel.com:912`; see [README › Setup fails with a proxy error](README.md#setup-fails-with-a-proxy-error) |
| `[WARN] No JMP installation was detected` | Expected on lab machines. On your laptop, see [JMP Pro](#2-jmp-pro) |

---

## 5. Verify the installation

```powershell
.venv\Scripts\python.exe -c "import openpyxl, pptx, tkinter; print('Python packages OK')"
.venv\Scripts\python.exe rmt_log_pipeline.py --version
.\marginiq.cmd --version                               # command-prompt entry point (any folder)
Test-Path "C:\Program Files\SAS\JMPPRO\17\jmp.exe"     # True on a JMP Pro machine
```

Then launch the GUI from the **MarginIQ** desktop icon and continue with the
[Quick Start Guide](RMT_QUICKSTART.md). To drive MarginIQ from scripts or other automation
tools, see [CLI_AUTOMATION.md](CLI_AUTOMATION.md).

---

## Still stuck?

Collect the following and contact the tool owner:

- the full `setup.bat` output (copy the console text),
- the output of `py -0p` and `where.exe python`,
- for GUI problems, the debug log (*Help ▸ Open Debug Log Folder*).
