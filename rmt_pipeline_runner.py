#!/usr/bin/env python3
"""
MarginIQ RMT Log Pipeline Runner - User-Friendly Wrapper

Simplifies running the MarginIQ extraction pipeline (rmt_log_pipeline.py)
with preset workflows and interactive menus.
"""

import os
import sys
import re
import csv as _csv
import subprocess
from datetime import datetime
from pathlib import Path

from rmt_project import TOOL_NAME, TOOL_SUBTITLE, TOOL_VERSION

# Configuration
HERE = Path(__file__).resolve().parent


def _resolve_python() -> str:
    """Prefer the tool-local .venv interpreter created by setup.bat."""
    venv_py = HERE / ".venv" / "Scripts" / "python.exe"
    if venv_py.exists():
        return str(venv_py)
    return sys.executable or r"C:/Program Files/Python314/python.exe"


PYTHON_EXE = _resolve_python()
PIPELINE_SCRIPT = str(HERE / "rmt_log_pipeline.py")
JMP_EXE = r"C:/Program Files/SAS/JMPPRO/17/jmp.exe"
# Optional PPT template; set RMT_PPT_TEMPLATE to apply one automatically.
PPT_TEMPLATE = os.environ.get("RMT_PPT_TEMPLATE", "")
AXIS_CONFIG  = str(HERE / "jmp_axis_settings.json")


def _has_template() -> bool:
    return bool(PPT_TEMPLATE) and Path(PPT_TEMPLATE).is_file()

# Canonical set of known parameters (same order as the pipeline PARAMS constant)
KNOWN_PARAMS = [
    "RecEnDelay", "TxDqsDelay", "RxDqsDelay", "TxDqDelay",
    "RxDqVrefByte", "TxVref", "ClkGrpPi", "CmdVref",
]

# Parameters whose margin values are typically constant (flat) across the
# sweep. They are excluded from JMP chart generation by default to reduce
# noise; the user can opt back in when prompted.
CONSTANT_PARAMS = ["RecEnDelay", "TxDqsDelay"]


# ---------------------------------------------------------------------------
# Field-detection helpers
# ---------------------------------------------------------------------------

def parse_selection(raw: str, count: int) -> list[int]:
    """Parse a user selection string into a list of 0-based indices.

    Supports single numbers (``2``), comma lists (``1,3``), ranges
    (``1-3``), mixtures (``1,3-4``) and ``all``. Out-of-range or invalid
    tokens are ignored. Returns an empty list when nothing valid is parsed.
    """
    raw = (raw or "").strip().lower()
    if not raw:
        return []
    if raw in ("all", "*", "a"):
        return list(range(count))
    picked: list[int] = []
    for token in re.split(r"[,\s]+", raw):
        if not token:
            continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", token)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            if lo > hi:
                lo, hi = hi, lo
            for n in range(lo, hi + 1):
                if 1 <= n <= count and (n - 1) not in picked:
                    picked.append(n - 1)
        elif token.isdigit():
            n = int(token)
            if 1 <= n <= count and (n - 1) not in picked:
                picked.append(n - 1)
    return picked


def detect_params_from_logs(input_path: str, pattern: str = "*.txt") -> tuple[list, list]:
    """Scan up to 10 log files for 'Params:' header lines.

    Returns (known_found, extra_found) where:
      known_found – KNOWN_PARAMS that appear in at least one log
      extra_found – additional parameter names found beyond KNOWN_PARAMS
    Both lists preserve order of first appearance.
    Falls back to (KNOWN_PARAMS copy, []) if nothing can be detected.
    """
    input_obj = Path(input_path)
    files: list[Path] = []
    if input_obj.is_file():
        files = [input_obj]
    elif input_obj.is_dir():
        files = sorted(input_obj.glob(pattern))[:10]

    found: list[str] = []
    seen: set[str] = set()
    for f in files:
        try:
            for line in f.read_text(encoding="utf-8", errors="ignore").splitlines():
                m = re.match(r"\s*Params:\s+(.+)", line)
                if m:
                    for p in m.group(1).split():
                        if p not in seen:
                            seen.add(p)
                            found.append(p)
        except Exception:
            pass

    if not found:
        return list(KNOWN_PARAMS), []

    known = [p for p in KNOWN_PARAMS if p in seen]
    extra = [p for p in found if p not in set(KNOWN_PARAMS)]
    return known, extra


def detect_params_from_excel(excel_path: str) -> tuple[list, list]:
    """Read header row of the All_RMT sheet and return (known_found, extra_found).

    Falls back to (KNOWN_PARAMS copy, []) on any error.
    """
    try:
        from openpyxl import load_workbook as _lw
        _wb = _lw(excel_path, read_only=True, data_only=True)
        _ws = _wb["All_RMT"] if "All_RMT" in _wb.sheetnames else _wb.active
        headers = [str(c.value) for c in next(_ws.iter_rows(min_row=1, max_row=1))]
        _wb.close()
    except Exception:
        return list(KNOWN_PARAMS), []

    plus_bases  = {h[:-1] for h in headers if h.endswith("+")}
    minus_bases = {h[:-1] for h in headers if h.endswith("-")}
    pairs = plus_bases & minus_bases

    known = [p for p in KNOWN_PARAMS if p in pairs]
    extra = sorted(p for p in pairs if p not in set(KNOWN_PARAMS))
    return known, extra


def detect_params_from_csv(csv_path: str) -> tuple[list, list]:
    """Scan CSV header for matching +/- column pairs.

    Returns (known_found, extra_found).
    Falls back to (KNOWN_PARAMS copy, []) on any error.
    """
    try:
        with open(csv_path, newline="", encoding="utf-8") as f:
            headers = next(_csv.reader(f))
    except Exception:
        return list(KNOWN_PARAMS), []

    plus_bases  = {h[:-1] for h in headers if h.endswith("+")}
    minus_bases = {h[:-1] for h in headers if h.endswith("-")}
    pairs = plus_bases & minus_bases          # must have both + and - columns

    known = [p for p in KNOWN_PARAMS if p in pairs]
    extra = sorted(p for p in pairs if p not in set(KNOWN_PARAMS))
    return known, extra


def _prompt_chart_fields(known: list[str], extra: list[str]) -> str:
    """Interactive field picker that shows detected fields and prompts about extras.

    Returns a comma-separated string of selected parameter names.
    """
    print("Available chart fields (enter numbers or names, comma-separated):")
    print("  Press Enter to use all detected fields.")
    print()

    # Constant-value fields (e.g. RecEnDelay, TxDqsDelay) are excluded by
    # default because their margins are flat across the sweep. List the
    # variable fields first; offer the constant ones as an opt-in.
    const_present = [f for f in known if f in CONSTANT_PARAMS]
    base_known    = [f for f in known if f not in CONSTANT_PARAMS]

    for i, f in enumerate(base_known, 1):
        print(f"  {i:2d}) {f}")

    active = list(base_known)

    if const_present:
        print()
        print(f"  ℹ️  Constant-value fields (typically flat across the sweep): "
              f"{', '.join(const_present)}")
        ans = input("  Include these constant fields in charts? (y/n) [default: n]: ").strip().lower()
        if ans == "y":
            start = len(active) + 1
            for i, f in enumerate(const_present, start):
                print(f"  {i:2d}) {f}  ← constant")
            active.extend(const_present)

    if extra:
        print()
        print(f"  ⚡ Extra fields detected in data: {', '.join(extra)}")
        ans = input("  Include extra fields in charts? (y/n) [default: n]: ").strip().lower()
        if ans == "y":
            start = len(active) + 1
            for i, f in enumerate(extra, start):
                print(f"  {i:2d}) {f}  ← extra")
            active.extend(extra)
            print()

    print()
    raw = input("Enter field numbers or names (comma-separated) [default: all]: ").strip()
    return _parse_fields_input(raw, active)


def run_pipeline(input_path, output_dir, stage=4, chart_fields=None, pattern="*.txt"):
    """
    Run the RMT extraction pipeline at the requested stage.

    Stages
    ------
    1  CSV Only       - Extracts RMT blocks to CSV files only.
    2  CSV + Excel    - Adds a multi-sheet Excel workbook (pivot-ready).
    3  CSV + Excel + JMP Charts
                     - Generates per-parameter JMP Graph Builder charts
                       (X=Params, Y=param±, Group X=Frequency) as PNG images.
                       Axis scale / reference lines are read from jmp_axis_settings.json.
    4  CSV + Excel + JMP Charts + PPT  (default)
                     - Everything in stage 3, plus assembles the JMP PNGs into a
                       PowerPoint presentation using the configured PPT template.

    Parameters
    ----------
    input_path  : str   Path to input folder or file.
    output_dir  : str   Path to output directory (created if absent).
    stage       : int   Workflow stage 1-4 (default 4).
    chart_fields: str   Comma-separated parameter names for JMP charts.
                        Pass None to use all 8 parameters.
    pattern     : str   File glob pattern when input_path is a folder.
    """

    ALL_FIELDS = ",".join(KNOWN_PARAMS)

    cmd = [
        PYTHON_EXE,
        PIPELINE_SCRIPT,
        "--input", input_path,
        "--pattern", pattern,
        "--outdir", output_dir,
        "--no-ask-chart-approval",
    ]

    if stage == 1:
        cmd.extend(["--chart-fields", ""])
        print("📄 Stage 1 — CSV Only")

    elif stage == 2:
        cmd.extend(["--chart-fields", ""])
        print("📊 Stage 2 — CSV + Excel")

    elif stage == 3:
        fields = chart_fields or ALL_FIELDS
        cmd += ["--chart-fields", fields,
                "--generate-jmp-charts",
                "--jmp-exe", JMP_EXE]
        if Path(AXIS_CONFIG).exists():
            cmd += ["--jmp-axis-config", AXIS_CONFIG]
        print(f"🎨 Stage 3 — CSV + Excel + JMP Charts  (fields: {fields})")

    elif stage == 4:
        fields = chart_fields or ALL_FIELDS
        cmd += ["--chart-fields", fields,
                "--generate-jmp-charts",
                "--jmp-exe", JMP_EXE]
        if _has_template():
            cmd += ["--ppt-template", PPT_TEMPLATE]
        if Path(AXIS_CONFIG).exists():
            cmd += ["--jmp-axis-config", AXIS_CONFIG]
        print(f"📑 Stage 4 — CSV + Excel + JMP Charts + PPT  (fields: {fields})")

    else:
        print(f"❌ Unknown stage {stage}.")
        return 2
    
    print(f"📁 Input:  {input_path}")
    print(f"📁 Output: {output_dir}")
    print()
    
    try:
        result = subprocess.run(cmd, check=True)
        print(f"\n✅ Pipeline completed successfully (Exit code: {result.returncode})")
        return result.returncode
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Pipeline failed (Exit code: {e.returncode})")
        if e.returncode == 3:
            print("   → Chart approval was denied")
        elif e.returncode == 4:
            print("   → JMP executable not found")
        elif e.returncode == 5:
            print("   → JMP execution failed")
        return e.returncode


def interactive_menu():
    """Show interactive menu for user to select stage and input."""
    
    print("\n" + "="*60)
    print(f"  {TOOL_NAME} {TOOL_VERSION} — {TOOL_SUBTITLE}")
    print("  RMT Log Pipeline Runner - User Friendly Interface")
    print("="*60 + "\n")
    
    # Stage selection
    print("Select Workflow Stage:")
    print()
    print("  1) 📄 CSV Only")
    print("     Extracts all START_RMT blocks to CSV files.")
    print("     Output: RMT_Combined_Similar.csv, RMT_Combined_Extended.csv, csv_by_file/")
    print()
    print("  2) 📊 CSV + Excel")
    print("     Adds a multi-sheet Excel workbook (one sheet per frequency + combined).")
    print("     Output: above + RMT_Extraction.xlsx")
    print()
    print("  3) 🎨 CSV + Excel + JMP Charts")
    print("     Generates per-parameter scatter charts in JMP (X=Params, Y=±values,")
    print("     Group X=Frequency). Axis scales come from jmp_axis_settings.json.")
    print("     Output: above + jmp_charts/*.png  (one PNG per parameter + dashboard)")
    print()
    print("  4) 📑 CSV + Excel + JMP Charts + PPT  ← default")
    print("     Everything in stage 3, plus assembles the JMP PNGs into a PowerPoint")
    print("     presentation using the configured PPT template.")
    print("     Output: above + RMT_Summary_JMP_Charts.pptx")
    print()
    print("  ─" * 30)
    print("  5) 🔄 Excel + JMP Charts + PPT + HTML from existing CSV")
    print("     Already have a CSV file or a folder of CSV files?")
    print("     Skip log extraction — load the CSV directly and produce all outputs.")
    print("     Output: RMT_Extraction.xlsx + jmp_charts/*.png")
    print("             RMT_Summary_JMP_Charts.pptx + RMT_Report.html (charts embedded)")
    print()
    print("  6) � JMP Charts + PPT + HTML from RMT_Extraction.xlsx")
    print("     Already have an Excel file (RMT_Extraction.xlsx or any .xlsx)?")
    print("     Read the All_RMT sheet and produce JMP charts, PPT, and HTML report.")
    print("     Output: jmp_charts/*.png + RMT_Summary_JMP_Charts.pptx + RMT_Report.html")
    print()
    print("  7) �📎 PPT from existing JMP charts")
    print("     Already have JMP PNG chart images in a folder?")
    print("     Skip everything — assemble those PNGs into a PowerPoint file.")
    print("     Output: RMT_Summary_JMP_Charts.pptx")
    print()

    stage_input = input("Enter stage number (1-7) [default: 4]: ").strip() or "4"

    try:
        stage = int(stage_input)
        if stage not in [1, 2, 3, 4, 5, 6, 7]:
            print("❌ Invalid stage. Using stage 4.")
            stage = 4
    except ValueError:
        print("❌ Invalid input. Using stage 4.")
        stage = 4

    # --------------------------------------------------------- #
    # Stage 5: JMP charts from an existing CSV (no log parsing)  #
    # --------------------------------------------------------- #
    if stage == 5:
        print()
        csv_raw = input("Enter path to CSV file or folder containing CSV files: ").strip()
        if not csv_raw:
            print("❌ CSV path required. Exiting.")
            return None
        csv_obj = Path(csv_raw)
        selected_files: list[Path] = []
        # ── If a directory is given, scan for CSV files ────────────────
        if csv_obj.is_dir():
            csv_files = sorted(csv_obj.glob("*.csv"))
            if not csv_files:
                print(f"❌ No CSV files found in: {csv_raw}")
                return None
            # Prefer RMT_Combined_Extended.csv (richest data set)
            extended = [f for f in csv_files if f.name == "RMT_Combined_Extended.csv"]
            if extended:
                selected_files = [extended[0]]
                print(f"  ✓ Auto-selected: {extended[0].name}")
            elif len(csv_files) == 1:
                selected_files = [csv_files[0]]
                print(f"  ✓ Using: {csv_files[0].name}")
            else:
                print(f"\n  Found {len(csv_files)} CSV files:")
                for _ci, _cf in enumerate(csv_files, 1):
                    print(f"    {_ci}) {_cf.name}")
                _raw_sel = input(
                    f"  Select CSV(s) [e.g. 1, 1-{len(csv_files)}, 1,3, or 'all'; "
                    f"default: 1]: "
                ).strip()
                _idxs = parse_selection(_raw_sel, len(csv_files))
                if not _idxs:
                    _idxs = [0]
                selected_files = [csv_files[i] for i in _idxs]
                print("  ✓ Selected: " + ", ".join(f.name for f in selected_files))
        elif csv_obj.is_file():
            selected_files = [csv_obj]
        else:
            print(f"❌ Path not found: {csv_raw}")
            return None

        # ── Decide processing mode when several files are selected ─────
        process_mode = "single"  # single | concat | separate
        if len(selected_files) > 1:
            print(f"\n  {len(selected_files)} files selected. How should they be processed?")
            print("    1) Concatenate into ONE combined report")
            print("       (merged data, differentiated by source via colour overlay)")
            print("    2) Process INDIVIDUALLY into separate output folders")
            _mode_raw = input("  Choose [1-2, default: 1]: ").strip()
            process_mode = "separate" if _mode_raw == "2" else "concat"

        _primary = selected_files[0]
        _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        if process_mode == "separate":
            _default_outdir = str(_primary.resolve().parent / f"rmt_output_{_ts}")
            _outdir_raw = input(
                f"Enter base output directory (each file gets its own subfolder) "
                f"[default: {_default_outdir}]: "
            ).strip()
        else:
            _default_outdir = str(_primary.resolve().parent / f"rmt_output_{_ts}")
            _outdir_raw = input(f"Enter output directory path [default: {_default_outdir}]: ").strip()
        output_dir = _outdir_raw or _default_outdir
        print()
        # Detect chart fields from the first selected file
        known, extra = detect_params_from_csv(str(_primary))
        chart_fields = _prompt_chart_fields(known, extra)
        print()
        print("-" * 60)

        if process_mode == "separate":
            return ("__jmp_from_csv_separate__", selected_files, output_dir, chart_fields)
        if process_mode == "concat":
            return ("__jmp_from_csv_multi__", selected_files, output_dir, chart_fields)
        return ("__jmp_from_csv__", str(_primary), output_dir, chart_fields)

    # --------------------------------------------------------- #
    # Stage 6: JMP Charts + PPT + HTML from Excel (.xlsx)        #
    # --------------------------------------------------------- #
    if stage == 6:
        print()
        xl_raw = input("Enter path to Excel file (.xlsx) or folder containing .xlsx files: ").strip()
        if not xl_raw:
            print("❌ Excel path required. Exiting.")
            return None
        xl_obj = Path(xl_raw)
        # ── If a directory is given, scan for xlsx files ───────────────
        if xl_obj.is_dir():
            xl_files = sorted(xl_obj.glob("*.xlsx"))
            if not xl_files:
                print(f"❌ No .xlsx files found in: {xl_raw}")
                return None
            # Prefer RMT_Extraction.xlsx
            preferred = [f for f in xl_files if f.name == "RMT_Extraction.xlsx"]
            if preferred:
                excel_path = str(preferred[0])
                print(f"  ✓ Auto-selected: {preferred[0].name}")
            elif len(xl_files) == 1:
                excel_path = str(xl_files[0])
                print(f"  ✓ Using: {xl_files[0].name}")
            else:
                print(f"\n  Found {len(xl_files)} Excel files:")
                for _xi, _xf in enumerate(xl_files, 1):
                    print(f"    {_xi}) {_xf.name}")
                _raw_sel = input(f"  Select Excel [1-{len(xl_files)}, default: 1]: ").strip()
                try:
                    _chosen = xl_files[int(_raw_sel) - 1]
                except (ValueError, IndexError):
                    _chosen = xl_files[0]
                excel_path = str(_chosen)
                print(f"  ✓ Using: {_chosen.name}")
        elif xl_obj.is_file():
            excel_path = xl_raw
        else:
            print(f"❌ Path not found: {xl_raw}")
            return None
        _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        _default_outdir = str(Path(excel_path).resolve().parent / f"rmt_output_{_ts}")
        _outdir_raw = input(f"Enter output directory path [default: {_default_outdir}]: ").strip()
        output_dir = _outdir_raw or _default_outdir
        print()
        known, extra = detect_params_from_excel(excel_path)
        chart_fields = _prompt_chart_fields(known, extra)
        print()
        print("-" * 60)
        return ("__jmp_from_excel__", excel_path, output_dir, chart_fields)

    # --------------------------------------------------------- #
    # Stage 7: PPT from an existing jmp_charts/ folder           #
    # --------------------------------------------------------- #
    if stage == 7:
        print()
        charts_dir = input("Enter path to folder containing JMP chart PNGs: ").strip()
        if not charts_dir:
            print("❌ Charts folder path required. Exiting.")
            return None
        _ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        _default_outdir = str(Path(charts_dir).resolve().parent / f"rmt_output_{_ts}")
        _outdir_raw = input(f"Enter output directory path [default: {_default_outdir}]: ").strip()
        output_dir = _outdir_raw or _default_outdir
        print()
        print("-" * 60)
        return ("__ppt_from_charts__", charts_dir, output_dir)
    
    # Input path selection
    print()
    input_path = input("Enter input folder or file path: ").strip()

    if not input_path:
        print("❌ Input path required. Exiting.")
        return None

    # Auto-detect file pattern and enumerate matched files
    input_obj = Path(input_path)
    pattern = "*.txt"
    matched_files: list = []

    if input_obj.is_dir():
        log_files = sorted(input_obj.glob("*.log"))
        txt_files = sorted(input_obj.glob("*.txt"))

        if log_files and not txt_files:
            print(f"\n📝 Found {len(log_files)} .log files (no .txt files)")
            pattern = "*.log"
        elif log_files and txt_files:
            print(f"\n📝 Found {len(log_files)} .log files and {len(txt_files)} .txt files")
            print("  Use default pattern (*.txt) or change to *.log?")
            pattern_input = input("Enter pattern [default: *.txt]: ").strip() or "*.txt"
            pattern = pattern_input
        elif txt_files:
            print(f"\n📝 Found {len(txt_files)} .txt files")
            pattern = "*.txt"
        else:
            print(f"\n⚠️  No common log files found in folder")
            pattern_input = input("Enter file pattern (e.g., *.log, *.txt) [default: *.txt]: ").strip() or "*.txt"
            pattern = pattern_input

        matched_files = sorted(input_obj.glob(pattern))
    elif input_obj.is_file():
        matched_files = [input_obj]

    # Multi-file output mode prompt
    separate_outputs = False
    if len(matched_files) > 1:
        print(f"\n📂 Found {len(matched_files)} log files:")
        for f in matched_files[:12]:
            print(f"     {f.name}")
        if len(matched_files) > 12:
            print(f"     … and {len(matched_files) - 12} more")
        print()
        print("  Output mode:")
        print("  1) 📦 Combine all files — one output folder  ← default")
        print("  2) 📂 Separate subfolders — one subfolder per log file")
        print()
        mode_in = input("  Enter choice (1/2) [default: 1]: ").strip() or "1"
        separate_outputs = (mode_in == "2")

    # Output directory selection
    _default_outdir = _default_output_dir(input_path)
    _outdir_prompt  = f"Enter output directory path [default: {_default_outdir}]: "
    output_dir = input(_outdir_prompt).strip() or _default_outdir

    # Chart fields (needed for stages 3/4 in both combined and separate mode)
    chart_fields = None
    if stage >= 3:
        print()
        known, extra = detect_params_from_logs(input_path, pattern)
        chart_fields = _prompt_chart_fields(known, extra)

    print()
    print("-" * 60)

    if separate_outputs:
        return ("__multi_separate__", matched_files, output_dir, stage, chart_fields)
    return input_path, output_dir, stage, chart_fields, pattern


def _default_output_dir(input_path: str) -> str:
    """Return the default output directory for a given input path.

    - folder  -> <folder>/rmt_output_<timestamp>
    - file    -> <parent>/rmt_output_<timestamp>
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"rmt_output_{ts}"
    p = Path(input_path).resolve()
    if p.is_file():
        return str(p.parent / folder_name)
    return str(p / folder_name)


def _parse_fields_input(fields_input: str, fields: list) -> str:
    """Convert user field input (numbers or names) to comma-separated param string."""
    if not fields_input:
        return ",".join(fields)
    selected = []
    for item in fields_input.split(","):
        item = item.strip()
        try:
            idx = int(item) - 1
            if 0 <= idx < len(fields):
                selected.append(fields[idx])
        except ValueError:
            if item in fields:
                selected.append(item)
    return ",".join(selected) if selected else ",".join(fields)


def simple_preset():
    """Simple preset commands without menu."""
    
    print("\n" + "="*60)
    print(f"  {TOOL_NAME} {TOOL_VERSION} — RMT Log Pipeline Quick Start")
    print("="*60 + "\n")
    
    input_path = input("📁 Input folder/file path: ").strip()
    if not input_path:
        print("❌ Input path required.")
        return None
    _default_outdir = _default_output_dir(input_path)
    _outdir_raw = input(f"📁 Output directory path [default: {_default_outdir}]: ").strip()
    output_dir = _outdir_raw or _default_outdir

    if not output_dir:
        print("❌ Output path required.")
    
    # Auto-detect file pattern
    from pathlib import Path
    input_obj = Path(input_path)
    
    pattern = "*.txt"  # default
    
    if input_obj.is_dir():
        log_files = list(input_obj.glob("*.log"))
        txt_files = list(input_obj.glob("*.txt"))
        
        if log_files and not txt_files:
            pattern = "*.log"
        elif log_files and txt_files:
            pattern_input = input("Found both .log and .txt files. Use which? [*.txt]: ").strip() or "*.txt"
            pattern = pattern_input
        elif txt_files:
            pattern = "*.txt"
        else:
            pattern_input = input("File pattern (*.log, *.txt, etc.) [*.txt]: ").strip() or "*.txt"
            pattern = pattern_input
    
    print("\nWorkflow stages:")
    print("  1) 📄 CSV Only                      — Fast data extraction")
    print("  2) 📊 CSV + Excel                   — Add pivot-ready workbook")
    print("  3) 🎨 CSV + Excel + JMP Charts      — Add per-param scatter charts")
    print("  4) 📑 CSV + Excel + JMP Charts + PPT — Add PowerPoint from charts (default)")
    print()

    choice = input("Select stage (1-4) [default: 4]: ").strip() or "4"

    try:
        stage = int(choice)
        if stage not in [1, 2, 3, 4]:
            stage = 4
    except ValueError:
        stage = 4
    
    return input_path, output_dir, stage, None, pattern


def cli_quick_commands():
    """Display quick copy-paste command examples."""
    
    print("\n" + "="*60)
    print(f"  {TOOL_NAME} {TOOL_VERSION} — Quick Copy-Paste Commands")
    print("="*60 + "\n")
    
    print("Stage 1: CSV Only")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --input "YOUR_INPUT_FOLDER" \\')
    print('  --pattern "*.txt" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --no-ask-chart-approval "--chart-fields="')
    print()

    print("Stage 2: CSV + Excel")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --input "YOUR_INPUT_FOLDER" \\')
    print('  --pattern "*.txt" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --no-ask-chart-approval "--chart-fields="')
    print()

    print("Stage 3: CSV + Excel + JMP Charts")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --input "YOUR_INPUT_FOLDER" \\')
    print('  --pattern "*.txt" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --no-ask-chart-approval \\')
    print('  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" \\')
    print('  --generate-jmp-charts \\')
    print(f'  --jmp-exe "{JMP_EXE}" \\')
    print(f'  --jmp-axis-config "{AXIS_CONFIG}"')
    print()

    print("Stage 4: CSV + Excel + JMP Charts + PPT  (default)")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --input "YOUR_INPUT_FOLDER" \\')
    print('  --pattern "*.txt" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --no-ask-chart-approval \\')
    print('  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" \\')
    print('  --generate-jmp-charts \\')
    print(f'  --jmp-exe "{JMP_EXE}" \\')
    print(f'  --ppt-template "{PPT_TEMPLATE or "YOUR_TEMPLATE.pptx"}" \\')
    print(f'  --jmp-axis-config "{AXIS_CONFIG}"')
    print()

    print("Stage 5: Excel + JMP Charts + PPT + HTML from existing CSV")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --jmp-from-csv "YOUR_CSV_FILE.csv" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" \\')
    print(f'  --jmp-exe "{JMP_EXE}" \\')
    print(f'  --ppt-template "{PPT_TEMPLATE or "YOUR_TEMPLATE.pptx"}" \\')
    print(f'  --jmp-axis-config "{AXIS_CONFIG}"')
    print()

    print("Stage 6: JMP Charts + PPT + HTML from RMT_Extraction.xlsx")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --jmp-from-excel "YOUR_EXCEL_FILE.xlsx" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print('  --chart-fields "RecEnDelay,TxDqsDelay,RxDqsDelay,TxDqDelay,RxDqVrefByte,TxVref,ClkGrpPi,CmdVref" \\')
    print(f'  --jmp-exe "{JMP_EXE}" \\')
    print(f'  --ppt-template "{PPT_TEMPLATE or "YOUR_TEMPLATE.pptx"}" \\')
    print(f'  --jmp-axis-config "{AXIS_CONFIG}"')
    print()

    print("Stage 7: PPT from existing JMP charts folder")
    print("─" * 60)
    print(f'python "{PIPELINE_SCRIPT}" \\')
    print('  --ppt-from-charts "YOUR_JMP_CHARTS_FOLDER" \\')
    print('  --outdir "YOUR_OUTPUT_FOLDER" \\')
    print(f'  --ppt-template "{PPT_TEMPLATE or "YOUR_TEMPLATE.pptx"}"')
    print()


def run_jmp_from_csv(csv_path, output_dir: str, chart_fields: str) -> int:
    """Run JMP chart generation from one or more existing CSV files.

    *csv_path* may be a single path (str/Path) or a list of paths. When several
    paths are given their rows are concatenated by the pipeline and tagged with
    a SourceFile column so the merged data sets can be differentiated.
    """
    if isinstance(csv_path, (list, tuple)):
        csv_list = [str(p) for p in csv_path]
    else:
        csv_list = [str(csv_path)]
    cmd = [
        PYTHON_EXE, PIPELINE_SCRIPT,
        "--jmp-from-csv", *csv_list,
        "--outdir", output_dir,
        "--chart-fields", chart_fields,
        "--generate-jmp-charts",
        "--jmp-exe", JMP_EXE,
    ]
    if _has_template():
        cmd += ["--ppt-template", PPT_TEMPLATE]
    if Path(AXIS_CONFIG).exists():
        cmd += ["--jmp-axis-config", AXIS_CONFIG]

    if len(csv_list) > 1:
        print(f"🔄 JMP Charts from {len(csv_list)} concatenated CSVs")
        for _c in csv_list:
            print(f"     • {Path(_c).name}")
    else:
        print(f"🔄 JMP Charts from CSV: {csv_list[0]}")
    print(f"📁 Output: {output_dir}")
    print()
    try:
        result = subprocess.run(cmd, check=True)
        print(f"\n✅ Completed (Exit code: {result.returncode})")
        return result.returncode
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Failed (Exit code: {e.returncode})")
        return e.returncode


def run_jmp_from_excel(excel_path: str, output_dir: str, chart_fields: str) -> int:
    """Run JMP chart + PPT + HTML generation from an existing Excel (.xlsx) file."""
    cmd = [
        PYTHON_EXE, PIPELINE_SCRIPT,
        "--jmp-from-excel", excel_path,
        "--outdir", output_dir,
        "--chart-fields", chart_fields,
        "--jmp-exe", JMP_EXE,
    ]
    if _has_template():
        cmd += ["--ppt-template", PPT_TEMPLATE]
    if Path(AXIS_CONFIG).exists():
        cmd += ["--jmp-axis-config", AXIS_CONFIG]

    print(f"📊 JMP Charts from Excel: {excel_path}")
    print(f"📁 Output: {output_dir}")
    print()
    try:
        result = subprocess.run(cmd, check=True)
        print(f"\n✅ Completed (Exit code: {result.returncode})")
        return result.returncode
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Failed (Exit code: {e.returncode})")
        return e.returncode


def run_ppt_from_charts(charts_dir: str, output_dir: str) -> int:
    """Assemble a PPT from an existing jmp_charts folder."""
    cmd = [
        PYTHON_EXE, PIPELINE_SCRIPT,
        "--ppt-from-charts", charts_dir,
        "--outdir", output_dir,
    ]
    if _has_template():
        cmd += ["--ppt-template", PPT_TEMPLATE]

    print(f"📎 PPT from charts folder: {charts_dir}")
    print(f"📁 Output: {output_dir}")
    print()
    try:
        result = subprocess.run(cmd, check=True)
        print(f"\n✅ Completed (Exit code: {result.returncode})")
        return result.returncode
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Failed (Exit code: {e.returncode})")
        return e.returncode


def main():
    """Main entry point."""
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--help":
            print(f"{TOOL_NAME} {TOOL_VERSION} — {TOOL_SUBTITLE}")
            print("RMT Log Pipeline Runner")
            print("=" * 60)
            print()
            print("Usage:")
            print("  python rmt_pipeline_runner.py           Interactive menu")
            print("  python rmt_pipeline_runner.py --quick   Minimal-prompt quick-start")
            print("  python rmt_pipeline_runner.py --cmds    Print copy-paste CLI commands")
            print("  python rmt_pipeline_runner.py --help    Show this help")
            print()
            print("Workflow Stages")
            print("-" * 60)
            print("  1  CSV Only")
            print("       Extracts all START_RMT blocks to CSV files.")
            print("       Outputs: RMT_Combined_Similar.csv")
            print("                RMT_Combined_Extended.csv")
            print("                csv_by_file/<stem>_RMT.csv")
            print()
            print("  2  CSV + Excel")
            print("       Adds a multi-sheet Excel workbook.")
            print("       Outputs: above + RMT_Extraction.xlsx")
            print("                  sheets: All_RMT, Freq_XXXX, File_<name>")
            print()
            print("  3  CSV + Excel + JMP Charts")
            print("       Opens JMP and generates Graph Builder scatter charts:")
            print("         X = Params (rank labels)")
            print("         Y = param+ (top panel) / param- (bottom panel)")
            print("         Group X = Frequency (side-by-side facets)")
            print("       Axis scale / tick increment / reference lines are read")
            print("       from jmp_axis_settings.json — edit that file to customise.")
            print("       Outputs: above + jmp_charts/<Param>.png  (one per parameter)")
            print("                        rmt_jmp_charts.jsl  (reusable JMP script)")
            print()
            print("  4  CSV + Excel + JMP Charts + PPT  ← DEFAULT")
            print("       Everything in stage 3, plus assembles the exported JMP PNGs")
            print("       into a PowerPoint file using the configured template.")
            print("       Outputs: above + RMT_Summary_JMP_Charts.pptx")
            print()
            print("  5  Excel + JMP Charts + PPT + HTML from existing CSV")
            print("       Have a CSV file or a folder of CSV files?")
            print("       Skip log extraction — load CSV directly and produce all outputs.")
            print("       Input : path to an Extended-format CSV file or containing folder")
            print("       Outputs: RMT_Extraction.xlsx + jmp_charts/<Param>.png")
            print("                RMT_Summary_JMP_Charts.pptx + RMT_Report.html")
            print()
            print("  6  JMP Charts + PPT + HTML from RMT_Extraction.xlsx")
            print("       Already have an Excel file (RMT_Extraction.xlsx or any .xlsx)?")
            print("       Read All_RMT sheet and produce JMP charts, PPT, and HTML.")
            print("       Input : path to a .xlsx file or containing folder")
            print("       Outputs: jmp_charts/<Param>.png + RMT_Summary_JMP_Charts.pptx")
            print("                RMT_Report.html (with embedded JMP charts)")
            print()
            print("  7  PPT from existing JMP charts folder")
            print("       Already ran JMP and have PNG chart images?")
            print("       Skip everything — just build the PowerPoint from those PNGs.")
            print("       Input : path to a folder of .png chart files")
            print("       Outputs: RMT_Summary_JMP_Charts.pptx")
            print()
            print("Chart Fields (stages 3 & 4)")
            print("-" * 60)
            print("  All 8 parameters are charted by default:")
            print("    RecEnDelay  TxDqsDelay  RxDqsDelay  TxDqDelay")
            print("    RxDqVrefByte  TxVref  ClkGrpPi  CmdVref")
            print("  To chart a subset, enter numbers or names when prompted.")
            print()
            print("Axis Customisation")
            print("-" * 60)
            print("  Edit jmp_axis_settings.json in this folder to change:")
            print("    min / max    Y-axis range")
            print("    inc          Major tick increment")
            print("    minor_ticks  Minor tick count (0 = off)")
            print("    ref_line     Value for the reference line")
            print()
            print("Configuration (edit top of this script)")
            print("-" * 60)
            print(f"  JMP executable : {JMP_EXE}")
            print(f"  PPT template   : {PPT_TEMPLATE or '(none - set RMT_PPT_TEMPLATE)'}")
            print(f"  Axis config    : {AXIS_CONFIG}")
            return
        elif sys.argv[1] == "--quick":
            result = simple_preset()
        elif sys.argv[1] == "--cmds":
            cli_quick_commands()
            return
        else:
            print("❌ Unknown option. Use --help for usage.")
            return
    else:
        result = interactive_menu()
    
    if result:
        # Stages 5-7 return a short tuple with a sentinel string as first element
        if isinstance(result, tuple) and result[0] == "__jmp_from_csv__":
            _, csv_path, output_dir, chart_fields = result
            run_jmp_from_csv(csv_path, output_dir, chart_fields)
        elif isinstance(result, tuple) and result[0] == "__jmp_from_csv_multi__":
            _, file_list, output_dir, chart_fields = result
            run_jmp_from_csv(file_list, output_dir, chart_fields)
        elif isinstance(result, tuple) and result[0] == "__jmp_from_csv_separate__":
            _, file_list, output_base, chart_fields = result
            total = len(file_list)
            for idx, f in enumerate(file_list, 1):
                file_outdir = str(Path(output_base) / Path(f).stem)
                print(f"\n{'='*60}")
                print(f"  [{idx}/{total}] Processing: {Path(f).name}")
                print(f"  Output -> {file_outdir}")
                print(f"{'='*60}")
                run_jmp_from_csv(str(f), file_outdir, chart_fields)
            print(f"\n✅ Done. Processed {total} CSV files into separate subfolders under:")
            print(f"   {output_base}")
        elif isinstance(result, tuple) and result[0] == "__jmp_from_excel__":
            _, excel_path, output_dir, chart_fields = result
            run_jmp_from_excel(excel_path, output_dir, chart_fields)
        elif isinstance(result, tuple) and result[0] == "__ppt_from_charts__":
            _, charts_dir, output_dir = result
            run_ppt_from_charts(charts_dir, output_dir)
        elif isinstance(result, tuple) and result[0] == "__multi_separate__":
            _, file_list, output_base, stage, chart_fields = result
            total = len(file_list)
            for idx, f in enumerate(file_list, 1):
                file_outdir = str(Path(output_base) / f.stem)
                print(f"\n{'='*60}")
                print(f"  [{idx}/{total}] Processing: {f.name}")
                print(f"  Output -> {file_outdir}")
                print(f"{'='*60}")
                run_pipeline(str(f), file_outdir, stage, chart_fields, f"*{f.suffix}")
            print(f"\n✅ Done. Processed {total} files into separate subfolders under:")
            print(f"   {output_base}")
        else:
            input_path, output_dir, stage, chart_fields, pattern = result
            run_pipeline(input_path, output_dir, stage, chart_fields, pattern)
    else:
        print("\n❌ Operation cancelled.")


if __name__ == "__main__":
    main()
