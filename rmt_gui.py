#!/usr/bin/env python3
"""
Intel CCG CVE DDR5 RMT Margin Analysis Tool - Novalake HX — Graphical Front-End
==============================================

A neat, industry-standard desktop GUI layered **on top of** the existing
command-line tools (``rmt_log_pipeline.py`` / ``rmt_pipeline_runner.py``).

It lets an engineer, without touching a terminal:

  * Graphically load one or many boot / MRC ``*.txt`` / ``*.log`` files
    (or existing CSV / Excel / PNG-chart folders).
  * Walk the same workflow stages exposed by the CLI (CSV → Excel → JMP
    charts → PPT → HTML) with every CLI flag surfaced as a control.
  * Pick which margin parameters become charts (``--chart-fields``).
  * Edit the JMP plot **X / Y axis** settings (Min, Max, Inc, Minor Ticks,
    Reference Line) per parameter — and **import a reference ``.jrp``** file
    to clone its axis scales.
  * Preview the exact CLI command before running, then run it with live
    streaming output.

No third-party GUI dependencies — pure ``tkinter`` (ships with Python).

Usage
-----
    python rmt_gui.py                 # launch the GUI
    python rmt_gui.py --screenshot out.png   # render once + save a preview PNG
"""
from __future__ import annotations

import fnmatch
import functools
import json
import logging
import os
import queue
import re
import shlex
import subprocess
import sys
import threading
import time
import types
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ──────────────────────────────────────────────────────────────────────────
# Configuration (kept in sync with rmt_pipeline_runner.py)
# ──────────────────────────────────────────────────────────────────────────
HERE            = Path(__file__).resolve().parent


def _resolve_python() -> str:
    """Interpreter used to spawn the pipeline as a subprocess.

    Prefers the tool-local ``.venv`` created by ``setup.bat``. Falls back to
    the running interpreter, mapping ``pythonw.exe`` back to ``python.exe`` so
    the subprocess keeps a working stdout/stderr for log streaming.
    """
    venv_py = HERE / ".venv" / "Scripts" / "python.exe"
    if venv_py.exists():
        return str(venv_py)
    current = sys.executable or r"C:/Program Files/Python314/python.exe"
    console = Path(current).with_name(Path(current).name.replace("pythonw", "python"))
    return str(console) if console.exists() else current


PYTHON_EXE      = _resolve_python()
PIPELINE_SCRIPT = str(HERE / "rmt_log_pipeline.py")
JMP_EXE         = r"C:/Program Files/SAS/JMPPRO/17/jmp.exe"
# Optional default PPT template; set RMT_PPT_TEMPLATE to pre-fill the GUI field.
PPT_TEMPLATE    = os.environ.get("RMT_PPT_TEMPLATE", "")
AXIS_CONFIG     = str(HERE / "jmp_axis_settings.json")

# ──────────────────────────────────────────────────────────────────────────
# Debug logging — every action taken in the app (every method call's entry,
# return value / exception, every subprocess command + its output, every
# messagebox shown, and any uncaught exception) is written in detail to a
# per-session log file under logs/. This is intended to be read by an LLM
# (or an engineer) after the fact to debug issues proactively.
# ──────────────────────────────────────────────────────────────────────────
LOG_DIR = HERE / "logs"


def _setup_debug_logger() -> tuple[logging.Logger, Path]:
    LOG_DIR.mkdir(exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"rmt_gui_debug_{ts}_{os.getpid()}.log"
    logger = logging.getLogger("rmt_gui")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    if not logger.handlers:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(logging.Formatter(
            "%(asctime)s.%(msecs)03d [%(levelname)-7s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"))
        logger.addHandler(handler)

    # Keep the log directory tidy: retain only the most recent 20 sessions.
    try:
        logs = sorted(LOG_DIR.glob("rmt_gui_debug_*.log"), key=lambda p: p.stat().st_mtime)
        for old in logs[:-20]:
            old.unlink(missing_ok=True)
    except Exception:
        pass

    logger.info("=" * 70)
    logger.info("RMT GUI debug session started (pid=%s, python=%s)",
                os.getpid(), sys.version.split()[0])
    logger.info("Log file: %s", log_path)
    logger.info("=" * 70)
    return logger, log_path


LOG, LOG_FILE = _setup_debug_logger()


def _safe_repr(value, limit: int = 300) -> str:
    """repr() that never raises and never floods the log with huge values."""
    try:
        s = repr(value)
    except Exception as e:
        s = f"<unrepr'able: {e}>"
    return s if len(s) <= limit else s[:limit] + "...<truncated>"


# Methods that fire on a fast repeating timer — skip per-call entry/exit
# tracing for these so the debug log stays readable; their real activity
# (e.g. actual subprocess output lines) is still logged explicitly inside.
_NO_TRACE_METHODS = {"_drain_log"}


def _instrument_methods(cls):
    """Wrap every plain method of *cls* so its entry, return value /
    exception are recorded to the debug log — gives full step-by-step
    tracing of the whole app without hand-instrumenting every method."""
    for name, attr in list(vars(cls).items()):
        if name.startswith("__") or name in _NO_TRACE_METHODS:
            continue
        if not isinstance(attr, types.FunctionType):
            continue

        def _make_wrapper(fn, fn_name):
            @functools.wraps(fn)
            def _wrapper(self, *args, **kwargs):
                LOG.debug("-> %s  args=%s kwargs=%s", fn_name,
                          _safe_repr(args), _safe_repr(kwargs))
                t0 = time.perf_counter()
                try:
                    result = fn(self, *args, **kwargs)
                except Exception:
                    LOG.exception("!! %s raised an exception", fn_name)
                    raise
                dt_ms = (time.perf_counter() - t0) * 1000
                LOG.debug("<- %s  returned=%s  (%.1fms)", fn_name,
                          _safe_repr(result), dt_ms)
                return result
            return _wrapper

        setattr(cls, name, _make_wrapper(attr, name))
    return cls


# Wrap tkinter's messagebox popups so every info/warning/error shown to the
# user is captured in the debug log (helpful context for why the app halted
# or what it told the user at any given step).
def _wrap_messagebox(kind: str, orig):
    def _fn(title=None, message=None, **kw):
        LOG.info("[messagebox.%s] title=%r message=%r", kind, title, message)
        return orig(title, message, **kw)
    return _fn


messagebox.showinfo = _wrap_messagebox("showinfo", messagebox.showinfo)
messagebox.showwarning = _wrap_messagebox("showwarning", messagebox.showwarning)
messagebox.showerror = _wrap_messagebox("showerror", messagebox.showerror)


# Catch anything that slips past Tkinter's callback boundary (button
# commands, trace callbacks, etc.) and anything truly uncaught at the top
# level — both are logged with a full traceback instead of being silently
# swallowed (Tkinter) or just printed to a console nobody is watching (GUI).
_orig_excepthook = sys.excepthook


def _log_excepthook(etype, value, tb) -> None:
    LOG.critical("UNCAUGHT EXCEPTION", exc_info=(etype, value, tb))
    _orig_excepthook(etype, value, tb)


sys.excepthook = _log_excepthook


def _tk_report_callback_exception(exc, val, tb) -> None:
    LOG.critical("Unhandled exception in a Tk callback", exc_info=(exc, val, tb))
    try:
        messagebox.showerror("Unexpected Error",
                              f"{val}\n\nDetails were written to the debug log:\n{LOG_FILE}")
    except Exception:
        pass

KNOWN_PARAMS = [
    "RecEnDelay", "TxDqsDelay", "RxDqsDelay", "TxDqDelay",
    "RxDqVrefByte", "TxVref", "ClkGrpPi", "CmdVref",
]
# Constant-value params are unchecked by default (flat across the sweep).
CONSTANT_PARAMS = ["RecEnDelay", "TxDqsDelay"]

JMP_COLORS = [
    "Orange", "Medium Light Red", "Red", "Medium Dark Blue", "Blue",
    "Green", "Black", "Gray",
]

# Source-type identifiers
SRC_LOGS = "logs"
SRC_CSV  = "csv"
SRC_XLSX = "xlsx"
SRC_PPT  = "ppt"

# Experiment profiles. "Non-thermal" is the original behaviour (single RMT per
# frequency/gear). "Thermal (DTR)" adapts log parsing for Dynamic Thermal
# Range experiments, where each log carries a Boot-temperature RMT and a
# Run-temperature RMT (BCRH / BHRC).
PROFILE_NONTHERMAL = "nonthermal"
PROFILE_THERMAL    = "thermal"
PROFILE_LABELS = {
    PROFILE_NONTHERMAL: "Non-Thermal Experiment",
    PROFILE_THERMAL:    "Thermal Experiment (DTR)",
}
PROFILE_BY_LABEL = {v: k for k, v in PROFILE_LABELS.items()}

# Palette (Intel-flavoured)
CLR_PRIMARY = "#0071C5"
CLR_DARK    = "#00305E"
CLR_BG      = "#f0f4f8"
CLR_CARD    = "#ffffff"
CLR_ACCENT  = "#059669"


APP_TITLE = "Intel CCG CVE DDR5 RMT Margin Analysis Tool - Novalake HX"
APP_SUBTITLE = "Client Computing Group  ·  Client Validation Engineering  ·  DDR5 Memory Training"


# ──────────────────────────────────────────────────────────────────────────
# Tooltip helper
# ──────────────────────────────────────────────────────────────────────────
class _Tooltip:
    """A simple hover tooltip that shows multi-line text next to a widget."""
    def __init__(self, widget: tk.Widget, text: str, delay: int = 500) -> None:
        self._w = widget
        self._text = text
        self._delay = delay
        self._id: str | None = None
        self._tw: tk.Toplevel | None = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._cancel)
        widget.bind("<ButtonPress>", self._cancel)

    def _schedule(self, _event=None) -> None:
        self._cancel()
        self._id = self._w.after(self._delay, self._show)

    def _cancel(self, _event=None) -> None:
        if self._id:
            self._w.after_cancel(self._id)
            self._id = None
        if self._tw:
            self._tw.destroy()
            self._tw = None

    def _show(self) -> None:
        x = self._w.winfo_rootx() + 20
        y = self._w.winfo_rooty() + self._w.winfo_height() + 4
        self._tw = tw = tk.Toplevel(self._w)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self._text, justify="left",
                 background="#fffde7", foreground="#1a1a1a",
                 relief="solid", borderwidth=1,
                 font=("Segoe UI", 9), padx=8, pady=6).pack()


# ──────────────────────────────────────────────────────────────────────────
# Reference .jrp axis parser
# ──────────────────────────────────────────────────────────────────────────
def parse_jrp_axis(text: str) -> dict[str, dict]:
    """Extract per-column axis settings from a JMP ``.jrp`` / JSL file.

    Looks for ``"<Column>", ScaleBox, {Min( ) Max( ) Inc( ) Minor Ticks( )
    Add Ref Line( ) ...}`` blocks and returns a mapping::

        { "RxDqsDelay+": {"min": 5, "max": 40, "inc": 2,
                          "minor_ticks": 0, "ref_line": 10}, ... }
    """
    out: dict[str, dict] = {}
    for m in re.finditer(r'"([^"]+)"\s*,\s*ScaleBox\s*,\s*\{', text):
        col = m.group(1)
        nxt = text.find("Dispatch(", m.end())
        seg = text[m.end(): nxt if nxt != -1 else m.end() + 600]

        def _num(pat: str):
            mm = re.search(pat, seg)
            return float(mm.group(1)) if mm else None

        out[col] = {
            "min":         _num(r"Min\(\s*(-?\d+(?:\.\d+)?)\s*\)"),
            "max":         _num(r"Max\(\s*(-?\d+(?:\.\d+)?)\s*\)"),
            "inc":         _num(r"Inc\(\s*(-?\d+(?:\.\d+)?)\s*\)"),
            "minor_ticks": _num(r"Minor Ticks\(\s*(-?\d+)\s*\)"),
            "ref_line":    _num(r"Add Ref Line\(\s*(-?\d+(?:\.\d+)?)"),
        }
    return out


def _fmt(v) -> str:
    """Render a number for an entry box (drop trailing .0)."""
    if v is None:
        return ""
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    except (TypeError, ValueError):
        return str(v)


# ──────────────────────────────────────────────────────────────────────────
# Main application
# ──────────────────────────────────────────────────────────────────────────
class RmtGuiApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("1100x780")
        self.root.minsize(960, 700)
        self.root.configure(bg=CLR_BG)
        # Route any exception raised inside a Tk callback (button command,
        # variable trace, combobox binding, etc.) to the debug log instead
        # of it being silently printed to a console nobody is watching.
        self.root.report_callback_exception = _tk_report_callback_exception

        self._proc: subprocess.Popen | None = None
        self._stop_requested = False
        self._log_q: queue.Queue[str] = queue.Queue()
        self._completed_stages: set[int] = set()
        self._active_stage: int | None = None
        self._run_ok = True
        self._tabs_unlocked = False
        self._stage1_csv: str | None = None
        self.var_csvinfo = tk.StringVar(value="Locked until Stage 1 CSV is generated from Tab 1.")

        # state vars
        self.var_source   = tk.StringVar(value=SRC_LOGS)
        self.var_profile  = tk.StringVar(value=PROFILE_LABELS[PROFILE_NONTHERMAL])
        self.var_pattern  = tk.StringVar(value="*.log;*.txt")
        self.var_outdir   = tk.StringVar()
        self.var_outdir.trace_add("write", lambda *_: self._update_recalc_state())
        self.var_excel    = tk.StringVar(value="RMT_Extraction.xlsx")
        self.var_ppt      = tk.StringVar(value="RMT_Summary.pptx")
        self.var_stage    = tk.IntVar(value=4)
        self.var_jmp      = tk.StringVar(value=JMP_EXE)
        self.var_tmpl     = tk.StringVar(value=PPT_TEMPLATE)
        self.var_jslonly  = tk.BooleanVar(value=False)
        self.var_approval = tk.BooleanVar(value=False)
        self.var_use_tmpl  = tk.BooleanVar(value=False)   # set properly after widget build
        self.var_multimode = tk.StringVar(value="concat")  # concat | separate
        self.var_useaxis  = tk.BooleanVar(value=True)
        self.var_axispath = tk.StringVar(value=AXIS_CONFIG)

        # Per-file axis support (separate mode, multiple files): each file
        # gets its own recalculated Min/Max/Inc, selectable via a dropdown.
        self.var_axis_file = tk.StringVar()
        self.per_file_axis: dict[str, dict] = {}   # file path -> {"defaults": {...}, "parameters": {...}}
        self._axis_file_map: dict[str, str] = {}   # display name -> file path
        self._axis_last_file: str | None = None

        self.param_vars: dict[str, tk.BooleanVar] = {
            p: tk.BooleanVar(value=(p not in CONSTANT_PARAMS)) for p in KNOWN_PARAMS
        }
        # axis grid: (param, side, field) -> StringVar
        self.axis_vars: dict[tuple, tk.StringVar] = {}
        self.def_vars = {
            "inc":          tk.StringVar(value="5"),
            "minor_ticks":  tk.StringVar(value="0"),
            "major_grid":   tk.StringVar(value="1"),
            "minor_grid":   tk.StringVar(value="1"),
            "plus_color":   tk.StringVar(value="Medium Light Red"),
            "minus_color":  tk.StringVar(value="Blue"),
        }

        self._build_style()
        self._build_header()
        self._build_footer()    # pack side=bottom FIRST so it always reserves space
        self._build_notebook()  # nav (side=bottom) then nb (expand fills middle)

        self._load_axis_json(Path(AXIS_CONFIG), silent=True)
        self._on_source_change()
        self._on_profile_change()
        self.root.after(150, self._drain_log)

    # ── styling ──────────────────────────────────────────────────────────
    def _build_style(self) -> None:
        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TFrame", background=CLR_BG)
        st.configure("Card.TFrame", background=CLR_CARD)
        st.configure("TLabel", background=CLR_BG, font=("Segoe UI", 10))
        st.configure("Card.TLabel", background=CLR_CARD, font=("Segoe UI", 10))
        st.configure("Hint.TLabel", background=CLR_CARD, foreground="#6b7280",
                     font=("Segoe UI", 9))
        st.configure("TLabelframe", background=CLR_CARD, relief="solid",
                     borderwidth=1, bordercolor="#d1d5db")
        st.configure("TLabelframe.Label", background=CLR_CARD,
                     foreground=CLR_DARK, font=("Segoe UI Semibold", 10))
        # Dimmed variant used by Multiple-File Handling when inactive
        st.configure("Multi.TLabelframe", background=CLR_CARD, relief="solid",
                     borderwidth=1, bordercolor="#e5e7eb")
        st.configure("Multi.TLabelframe.Label", background=CLR_CARD,
                     foreground="#9ca3af", font=("Segoe UI Semibold", 10))
        st.configure("TButton", font=("Segoe UI", 10), padding=5)
        st.configure("Accent.TButton", font=("Segoe UI Semibold", 10),
                     foreground="#fff", background=CLR_PRIMARY, padding=7)
        st.map("Accent.TButton",
               background=[("active", CLR_DARK), ("disabled", "#9ca3af")])
        st.configure("Run.TButton", font=("Segoe UI Semibold", 11),
                     foreground="#fff", background=CLR_ACCENT, padding=8)
        st.map("Run.TButton",
               background=[("active", "#047857"), ("disabled", "#9ca3af")])
        st.configure("TNotebook", background=CLR_BG, borderwidth=0)
        st.configure("TNotebook.Tab", font=("Segoe UI Semibold", 10),
                     padding=(16, 8))
        st.configure("TCheckbutton", background=CLR_CARD, font=("Segoe UI", 10))
        st.configure("TRadiobutton", background=CLR_CARD, font=("Segoe UI", 10))
        st.configure("TEntry", padding=3)

    # ── header ───────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        hdr = tk.Frame(self.root, bg=CLR_PRIMARY, height=76)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        # Intel logo band (left accent)
        tk.Frame(hdr, bg="#00C7FD", width=6).pack(side="left", fill="y")
        left = tk.Frame(hdr, bg=CLR_PRIMARY)
        left.pack(side="left", padx=(10, 0))
        tk.Label(left, text=APP_TITLE,
                 bg=CLR_PRIMARY, fg="#ffffff",
                 font=("Segoe UI Semibold", 16)).pack(anchor="w")
        tk.Label(left, text=APP_SUBTITLE,
                 bg=CLR_PRIMARY, fg="#93c5fd",
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(hdr, text="MRC RMT Log Pipeline  ",
                 bg=CLR_PRIMARY, fg="#cfe6fb",
                 font=("Segoe UI", 9)).pack(side="right", padx=10)

    # ── notebook ─────────────────────────────────────────────────────────
    def _build_notebook(self) -> None:
        self.nb = ttk.Notebook(self.root)
        self.tab_input = ttk.Frame(self.nb)
        self.tab_axis  = ttk.Frame(self.nb)
        self.tab_run   = ttk.Frame(self.nb)
        self.nb.add(self.tab_input, text="  1 · Input & Workflow  ")
        self.nb.add(self.tab_axis,  text="  2 · Parameters & Axis  ")
        self.nb.add(self.tab_run,   text="  3 · Preview & Run  ")
        self._build_nav()       # nav (side=bottom) BEFORE notebook expand
        self.nb.pack(fill="both", expand=True, padx=12, pady=(10, 6))
        self._build_tab_input()
        self._build_tab_axis()
        self._build_tab_run()
        # Lock downstream tabs until Stage 1 (CSV) is generated from Tab 1.
        for i in (1, 2):
            self.nb.tab(i, state="disabled")

    # ── global Next / Previous navigation ────────────────────────────────
    def _build_nav(self) -> None:
        # side="bottom" keeps this bar always visible regardless of notebook
        # content height.  It stacks above the footer (packed side=bottom first).
        nav = ttk.Frame(self.root)
        nav.pack(fill="x", side="bottom", padx=12, pady=(0, 4))
        self.btn_prev = ttk.Button(nav, text="\u2190 Previous", command=self._prev_tab)
        self.btn_prev.pack(side="left")
        self.lbl_step = ttk.Label(nav, text="")
        self.lbl_step.pack(side="left", expand=True)
        self.btn_next = ttk.Button(nav, text="Next \u2192", style="Accent.TButton",
                                   command=self._next_tab)
        self.btn_next.pack(side="right")
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._update_nav()

    def _on_tab_changed(self, event=None) -> None:
        """Guard entry to the last (Preview & Run) tab: whether reached via
        the 'Next' button or by clicking the tab header directly, every
        Chart-enabled parameter's per-parameter Y-axis values must be fully
        populated first."""
        i = self.nb.index(self.nb.select())
        if i == self._tab_count() - 1 and self._tabs_unlocked:
            missing = self._validate_axis_values()
            if missing:
                self.nb.select(self.tab_axis)
                messagebox.showwarning(
                    "Axis Values Missing",
                    "Please fill in all per-parameter Y-axis values "
                    "(Min / Max / Inc / Ref) for every chart-enabled "
                    "parameter before continuing to Preview & Run:\n\n"
                    + "\n".join(missing))
                return
        self._update_nav()

    def _tab_count(self) -> int:
        return len(self.nb.tabs())

    def _next_tab(self) -> None:
        i = self.nb.index(self.nb.select())
        # From Tab 1: generate the CSV (Stage 1) before unlocking later tabs.
        if i == 0 and not self._tabs_unlocked:
            self._advance_from_input()
            return
        if i < self._tab_count() - 1:
            self.nb.select(i + 1)

    def _prev_tab(self) -> None:
        i = self.nb.index(self.nb.select())
        if i > 0:
            self.nb.select(i - 1)

    def _advance_from_input(self) -> None:
        """Tab 1 'Next': for boot-logs run Stage 1 to make the CSV, then unlock
        downstream tabs which reuse that CSV (no re-extraction later)."""
        if self._proc is not None:
            messagebox.showinfo("Running", "A run is already in progress.")
            return
        if self.var_source.get() != SRC_LOGS:
            # CSV / XLSX / PNG sources already have the data — unlock directly.
            if not self._validate():
                return
            csvs = [p for p in self._all_paths() if Path(p).suffix.lower() == ".csv"]
            self._unlock_tabs(csvs)
            self.nb.select(self.tab_axis)
            return
        # Boot-log source → generate Stage 1 CSV, unlock on completion.
        self._run_stage(1)

    def _find_stage1_csvs(self) -> list[str]:
        """Locate the RMT_Combined_Extended.csv file(s) produced by Stage 1."""
        outdir = Path(self.var_outdir.get().strip())
        found: list[str] = []
        if outdir.is_dir():
            direct = outdir / "RMT_Combined_Extended.csv"
            if direct.exists():
                found.append(str(direct))
            else:
                found.extend(str(p) for p in outdir.glob("*/RMT_Combined_Extended.csv"))
        return found

    def _unlock_tabs(self, csvs: list[str]) -> None:
        """Enable downstream tabs and pin further processing to the CSV(s)."""
        self._stage1_csv = csvs[0] if csvs else None
        self._tabs_unlocked = True
        for i in (1, 2):
            self.nb.tab(i, state="normal")
        if csvs:
            outdir = self.var_outdir.get().strip()
            self.var_csvinfo.set(
                f"Output: {outdir}    |    CSV in use: "
                f"{', '.join(Path(c).name for c in csvs)}")
        self._update_nav()

    def _update_nav(self) -> None:
        i = self.nb.index(self.nb.select())
        n = self._tab_count()
        self.lbl_step.configure(text=f"Step {i + 1} of {n}")
        self.btn_prev.configure(state="normal" if i > 0 else "disabled")
        if i == 0 and not self._tabs_unlocked:
            # Tab 1 acts as Stage 1: generate the CSV, then unlock the rest.
            self.btn_next.configure(text="Generate CSV & Continue \u2192",
                                    state="normal")
        else:
            self.btn_next.configure(text="Next \u2192",
                                    state="normal" if i < n - 1 else "disabled")
        if i == n - 1:
            self._refresh_preview()

    # ── scrollable tab helper ─────────────────────────────────────────────
    def _make_scrollable_tab(self, tab: ttk.Frame) -> ttk.Frame:
        """Wrap *tab* in a Canvas+Scrollbar so its content scrolls when the
        window is too short to show everything at once."""
        canvas = tk.Canvas(tab, bg=CLR_BG, highlightthickness=0)
        vs = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        def _on_inner_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_resize(e):
            # Keep inner frame width in sync with canvas so LabelFrames fill
            # the full width when the window is resized or maximized.
            canvas.itemconfig(win_id, width=e.width)

        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", _on_inner_resize)
        canvas.bind("<Configure>", _on_canvas_resize)
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        def _mw(e):
            canvas.yview_scroll(-1 * (e.delta // 120), "units")

        # Activate mouse-wheel scrolling only while the pointer is over this tab.
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _mw))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return inner

    # ── Tab 1: Input & Workflow ──────────────────────────────────────────
    def _build_tab_input(self) -> None:
        # Wrap in a scrollable canvas so all sections remain reachable at any
        # window height, including small/laptop screens.
        t = self._make_scrollable_tab(self.tab_input)

        # Experiment profile
        prof = ttk.LabelFrame(t, text="Experiment Profile")
        prof.pack(fill="x", padx=10, pady=(10, 6))
        prow = ttk.Frame(prof, style="Card.TFrame")
        prow.pack(fill="x", padx=8, pady=6)
        ttk.Label(prow, text="Profile:", style="Card.TLabel", width=16).pack(side="left")
        self.cbo_profile = ttk.Combobox(
            prow, textvariable=self.var_profile, state="readonly", width=32,
            values=[PROFILE_LABELS[PROFILE_NONTHERMAL],
                    PROFILE_LABELS[PROFILE_THERMAL]])
        self.cbo_profile.pack(side="left", padx=4)
        self.cbo_profile.bind("<<ComboboxSelected>>", self._on_profile_change)
        self.lbl_profile_hint = ttk.Label(prof, text="", style="Hint.TLabel")
        self.lbl_profile_hint.pack(anchor="w", padx=10, pady=(0, 6))

        # Data source
        src = ttk.LabelFrame(t, text="Data Source")
        src.pack(fill="x", padx=10, pady=(10, 6))
        for txt, val in [
            ("Boot / MRC logs  (*.txt, *.log)", SRC_LOGS),
            ("Existing CSV file(s)",            SRC_CSV),
            ("Existing Excel (.xlsx)",          SRC_XLSX),
            ("PNG charts \u2192 PPT only",      SRC_PPT),
        ]:
            ttk.Radiobutton(src, text=txt, value=val, variable=self.var_source,
                            command=self._on_source_change).pack(side="left",
                                                                 padx=8, pady=6)

        # Files
        ff = ttk.LabelFrame(t, text="Files / Folder")
        ff.pack(fill="x", padx=10, pady=6)
        bar = ttk.Frame(ff, style="Card.TFrame")
        bar.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Button(bar, text="Add Files\u2026", command=self._add_files).pack(side="left", padx=3)
        ttk.Button(bar, text="Add Folder\u2026", command=self._add_folder).pack(side="left", padx=3)
        ttk.Button(bar, text="Remove", command=self._remove_sel).pack(side="left", padx=3)
        ttk.Button(bar, text="Clear", command=self._clear_files).pack(side="left", padx=3)
        ttk.Label(bar, text="File Mask:", style="Card.TLabel").pack(side="left", padx=(16, 4))
        ent_pat = ttk.Entry(bar, textvariable=self.var_pattern, width=18)
        ent_pat.pack(side="left")
        _Tooltip(ent_pat,
                 "Filter which files are listed when you click 'Add Folder'.\n"
                 "Separate multiple patterns with  ;  (semicolon).\n"
                 "\n"
                 "  Windows wildcard patterns supported:\n"
                 "    *        matches any number of characters\n"
                 "    ?        matches exactly one character\n"
                 "    [abc]   matches one character in the set\n"
                 "\n"
                 "  Examples:\n"
                 "    *.log             – all .log files\n"
                 "    *.log;*.txt       – .log and .txt files (default)\n"
                 "    boot_*.log        – files starting with 'boot_'\n"
                 "    4800_G?.log       – e.g. 4800_G2.log, 4800_G4.log\n"
                 "    *RMT*;*mrc*.txt   – any file with 'RMT' or 'mrc' in name\n"
                 "    *.log;*.txt;*.dat – three extension types\n"
                 "\n"
                 "  Note: patterns are case-insensitive on Windows.")
        ttk.Button(bar, text="Re-apply",
                   command=self._reapply_pattern).pack(side="left", padx=(4, 0))

        lstwrap = ttk.Frame(ff, style="Card.TFrame")
        lstwrap.pack(fill="x", padx=8, pady=(0, 8))
        self.lst = tk.Listbox(lstwrap, selectmode="extended", height=8,
                              font=("Consolas", 9), activestyle="none")
        sb = ttk.Scrollbar(lstwrap, orient="vertical", command=self.lst.yview)
        self.lst.configure(yscrollcommand=sb.set)
        self.lst.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.lbl_src_hint = ttk.Label(
            ff, text="", style="Hint.TLabel")
        self.lbl_src_hint.pack(anchor="w", padx=10, pady=(0, 6))

        # Multiple-file handling (mirrors the CLI concat / separate prompt)
        self.multi_frame = ttk.LabelFrame(t, text="Multiple-File Handling",
                                          style="Multi.TLabelframe")
        self.multi_frame.pack(fill="x", padx=10, pady=6)
        self.lbl_count = ttk.Label(self.multi_frame, text="No files loaded.",
                                   style="Hint.TLabel")
        self.lbl_count.pack(anchor="w", padx=10, pady=(6, 2))
        rrow = ttk.Frame(self.multi_frame, style="Card.TFrame")
        rrow.pack(fill="x", padx=8, pady=(0, 8))
        self.rb_concat = ttk.Radiobutton(
            rrow, text="Concatenate into ONE combined report  (overlay by source)",
            value="concat", variable=self.var_multimode, command=self._on_multimode_change)
        self.rb_concat.pack(anchor="w", padx=4, pady=1)
        self.rb_separate = ttk.Radiobutton(
            rrow, text="Process each file SEPARATELY  (one subfolder per file)",
            value="separate", variable=self.var_multimode, command=self._on_multimode_change)
        self.rb_separate.pack(anchor="w", padx=4, pady=1)

        # Output + names
        out = ttk.LabelFrame(t, text="Output")
        out.pack(fill="x", padx=10, pady=6)
        row = ttk.Frame(out, style="Card.TFrame"); row.pack(fill="x", padx=8, pady=6)
        ttk.Label(row, text="Output directory:", style="Card.TLabel", width=16).pack(side="left")
        ttk.Entry(row, textvariable=self.var_outdir).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="Browse\u2026", command=self._pick_outdir).pack(side="left")
        row2 = ttk.Frame(out, style="Card.TFrame"); row2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row2, text="Excel name:", style="Card.TLabel", width=16).pack(side="left")
        ttk.Entry(row2, textvariable=self.var_excel, width=24).pack(side="left", padx=4)
        ttk.Label(row2, text="PPT name:", style="Card.TLabel").pack(side="left", padx=(16, 4))
        ttk.Entry(row2, textvariable=self.var_ppt, width=24).pack(side="left", padx=4)

        # PPT Template (optional) — Intel-branded .pptx master slide deck
        ppt_frame = ttk.LabelFrame(t, text="PPT Template  (optional)")
        ppt_frame.pack(fill="x", padx=10, pady=6)
        ppt_top = ttk.Frame(ppt_frame, style="Card.TFrame")
        ppt_top.pack(fill="x", padx=8, pady=(6, 2))
        self.var_use_tmpl = tk.BooleanVar(value=bool(self.var_tmpl.get().strip()))
        self.cb_use_tmpl = ttk.Checkbutton(
            ppt_top,
            text="Use a custom Intel PowerPoint template (.pptx) for branding / slide master",
            variable=self.var_use_tmpl, command=self._on_tmpl_toggle)
        self.cb_use_tmpl.pack(side="left")
        ppt_row = ttk.Frame(ppt_frame, style="Card.TFrame")
        ppt_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(ppt_row, text="Template file:", style="Card.TLabel", width=14).pack(side="left")
        self.ent_tmpl = ttk.Entry(ppt_row, textvariable=self.var_tmpl)
        self.ent_tmpl.pack(side="left", fill="x", expand=True, padx=4)
        self.btn_tmpl = ttk.Button(ppt_row, text="Browse\u2026",
                   command=lambda: self._pick_file(self.var_tmpl, [("PowerPoint", "*.pptx")]))
        self.btn_tmpl.pack(side="left")
        ttk.Label(ppt_row,
                  text="Leave blank to use the pipeline's built-in default layout.",
                  style="Hint.TLabel").pack(side="left", padx=(12, 0))
        # Set initial state
        self.root.after(10, self._on_tmpl_toggle)

        # Workflow stage
        self.stage_frame = ttk.LabelFrame(t, text="Workflow Stage  (boot-log extraction)")
        self.stage_frame.pack(fill="x", padx=10, pady=6)
        stages = [
            (1, "CSV only", "Extract START_RMT blocks to CSV."),
            (2, "CSV + Excel", "Add multi-sheet Excel workbook."),
            (3, "CSV + Excel + JMP charts", "Generate per-parameter JMP PNGs."),
            (4, "CSV + Excel + JMP + PPT  (default)", "Everything + PowerPoint deck."),
        ]
        for n, ttl, desc in stages:
            r = ttk.Frame(self.stage_frame, style="Card.TFrame"); r.pack(fill="x", padx=8, pady=1)
            ttk.Radiobutton(r, text=f"{n})  {ttl}", value=n, variable=self.var_stage,
                            command=self._refresh_preview).pack(side="left")
            ttk.Label(r, text="— " + desc, style="Hint.TLabel").pack(side="left", padx=6)

        # Tools & options
        tools = ttk.LabelFrame(t, text="Tools & Options")
        tools.pack(fill="x", padx=10, pady=(6, 10))
        r = ttk.Frame(tools, style="Card.TFrame"); r.pack(fill="x", padx=8, pady=5)
        ttk.Label(r, text="JMP executable:", style="Card.TLabel", width=16).pack(side="left")
        ttk.Entry(r, textvariable=self.var_jmp).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(r, text="Browse\u2026",
                   command=lambda: self._pick_file(self.var_jmp, [("JMP", "*.exe")])).pack(side="left")
        r = ttk.Frame(tools, style="Card.TFrame"); r.pack(fill="x", padx=8, pady=(5, 8))
        ttk.Checkbutton(r, text="Generate JSL only (don't launch JMP)",
                        variable=self.var_jslonly, command=self._refresh_preview).pack(side="left", padx=6)
        ttk.Checkbutton(r, text="Ask chart approval (CLI prompt)",
                        variable=self.var_approval, command=self._refresh_preview).pack(side="left", padx=6)

    # ── Tab 2: Parameters ────────────────────────────────────────────────
    def _csv_banner(self, parent) -> None:
        """A green banner echoing the Stage-1 output dir + CSV used downstream."""
        bn = tk.Frame(parent, bg="#ecfdf5", highlightbackground="#a7f3d0",
                      highlightthickness=1)
        bn.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(bn, textvariable=self.var_csvinfo, bg="#ecfdf5", fg="#065f46",
                 font=("Segoe UI", 9), anchor="w", justify="left").pack(
            side="left", padx=10, pady=4)

    def _build_tab_axis(self) -> None:
        t = self.tab_axis
        self._csv_banner(t)

        ttk.Label(t,
                  text="Note: these axis values are pre-configured from the log's measured "
                       "values (defaults in jmp_axis_settings.json). Adjust per parameter, "
                       "or import a reference .jrp to clone its scales.",
                  style="Hint.TLabel", wraplength=940, justify="left").pack(anchor="w", padx=14, pady=(10, 0))

        top = ttk.Frame(t); top.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Checkbutton(top, text="Apply axis settings (--jmp-axis-config)",
                        variable=self.var_useaxis, command=self._refresh_preview).pack(side="left", padx=4)
        self.btn_recalc = ttk.Button(top, text="Recalculate from csv file",
                                      style="Accent.TButton", state="disabled",
                                      command=self._recalculate_axis)
        self.btn_recalc.pack(side="left", padx=(12, 3))
        _hint_lbl = ttk.Label(top, text="\u2139", style="Hint.TLabel")
        _hint_lbl.pack(side="left")
        _Tooltip(_hint_lbl,
                 "Re-calculate Min, Max, and Inc from the loaded data files.\n"
                 "+ Ref and \u2212 Ref values are preserved.\n\n"
                 "Works directly from raw .log/.txt files, CSV files loaded\n"
                 "on Tab 1, or a previously extracted Combined CSV.\n\n"
                 "When 'Process each file SEPARATELY' is selected with\n"
                 "multiple files, each file gets its OWN Min/Max/Inc — use\n"
                 "the 'Per-file axis values for:' dropdown below to view them.")
        ttk.Button(top, text="Load .jrp reference\u2026", style="Accent.TButton",
                   command=self._load_jrp).pack(side="right", padx=3)
        ttk.Button(top, text="Load JSON\u2026", command=self._browse_axis_json).pack(side="right", padx=3)
        ttk.Button(top, text="Save JSON\u2026", command=self._save_axis_json).pack(side="right", padx=3)

        # Per-file axis selector — only shown for multiple files in "Process
        # each file SEPARATELY" mode, where each file can have its own
        # recalculated Min/Max/Inc (values naturally differ per log).
        self.axis_file_row = ttk.Frame(t)
        ttk.Label(self.axis_file_row, text="Per-file axis values for:",
                  style="Hint.TLabel").pack(side="left", padx=(4, 4))
        self.cbo_axis_file = ttk.Combobox(self.axis_file_row, textvariable=self.var_axis_file,
                                          state="readonly", width=48)
        self.cbo_axis_file.pack(side="left", padx=4)
        self.cbo_axis_file.bind("<<ComboboxSelected>>", self._on_axis_file_selected)
        _afile_hint = ttk.Label(self.axis_file_row, text="\u2139", style="Hint.TLabel")
        _afile_hint.pack(side="left")
        _Tooltip(_afile_hint,
                 "Each file gets its own INDEPENDENT config \u2014 both the\n"
                 "per-parameter Y-axis grid (Min/Max/Inc/Ref) AND the\n"
                 "Defaults section (Inc/Minor ticks/Major/Minor grid/ref\n"
                 "colors) \u2014 computed by 'Recalculate' and editable per file.\n"
                 "Switch files here to view or hand-edit that file's values.")
        # Not packed here — _refresh_axis_file_selector() shows/hides it.

        # Parameter selection controls (charted params)
        psel = ttk.Frame(t); psel.pack(fill="x", padx=10, pady=(0, 2))
        self._axis_psel_ref = psel  # anchor: per-file row is inserted right before this
        ttk.Label(psel, text="Charted parameters \u2014 tick the box to include; "
                            "unchecked rows are excluded and dimmed:",
                  style="Hint.TLabel").pack(side="left", padx=4)
        ttk.Button(psel, text="Reset Defaults", command=self._reset_params).pack(side="right", padx=3)
        ttk.Button(psel, text="Clear All", command=lambda: self._set_all_params(False)).pack(side="right", padx=3)
        ttk.Button(psel, text="Select All", command=lambda: self._set_all_params(True)).pack(side="right", padx=3)

        # Defaults
        df = ttk.LabelFrame(t, text="Defaults")
        df.pack(fill="x", padx=10, pady=6)
        r = ttk.Frame(df, style="Card.TFrame"); r.pack(fill="x", padx=8, pady=8)
        def _dfield(lbl, key, w=5, combo=None):
            ttk.Label(r, text=lbl, style="Card.TLabel").pack(side="left", padx=(8, 2))
            if combo:
                ttk.Combobox(r, textvariable=self.def_vars[key], values=combo,
                             width=16, state="readonly").pack(side="left")
            else:
                ttk.Entry(r, textvariable=self.def_vars[key], width=w).pack(side="left")
        _dfield("Inc", "inc"); _dfield("Minor ticks", "minor_ticks")
        _dfield("Major grid", "major_grid"); _dfield("Minor grid", "minor_grid")
        r2 = ttk.Frame(df, style="Card.TFrame"); r2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(r2, text="+ ref color", style="Card.TLabel").pack(side="left", padx=(8, 2))
        ttk.Combobox(r2, textvariable=self.def_vars["plus_color"], values=JMP_COLORS,
                     width=18, state="readonly").pack(side="left")
        ttk.Label(r2, text="\u2212 ref color", style="Card.TLabel").pack(side="left", padx=(16, 2))
        ttk.Combobox(r2, textvariable=self.def_vars["minus_color"], values=JMP_COLORS,
                     width=18, state="readonly").pack(side="left")

        # Per-parameter axis grid
        gf = ttk.LabelFrame(t, text="Per-Parameter Y-Axis  (+ panel and \u2212 panel)")
        gf.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        canvas = tk.Canvas(gf, bg=CLR_CARD, highlightthickness=0)
        vs = ttk.Scrollbar(gf, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        vs.pack(side="right", fill="y", pady=6)

        hdr = ["Chart", "Parameter", "+Min", "+Max", "+Inc", "+Ref",
               "\u2212Min", "\u2212Max", "\u2212Inc", "\u2212Ref", "Preview"]
        for c, h in enumerate(hdr):
            ttk.Label(inner, text=h, style="Card.TLabel",
                      font=("Segoe UI Semibold", 9)).grid(row=0, column=c, padx=4, pady=(4, 6))
        fields = ["min", "max", "inc", "ref_line"]
        self.axis_preview_canvases: dict[str, tk.Canvas] = {}
        self.axis_row_entries: dict[str, list] = {}
        self.axis_row_labels: dict[str, ttk.Label] = {}
        for r, p in enumerate(KNOWN_PARAMS, start=1):
            cb = ttk.Checkbutton(inner, variable=self.param_vars[p],
                                 command=lambda _p=p: self._on_param_toggle(_p))
            cb.grid(row=r, column=0, padx=2, pady=2)
            tag = "  (optional)" if p in CONSTANT_PARAMS else ""
            lab = ttk.Label(inner, text=p + tag, style="Card.TLabel")
            lab.grid(row=r, column=1, sticky="w", padx=6, pady=2)
            self.axis_row_labels[p] = lab
            entries = []
            c = 2
            for side in ("plus", "minus"):
                for f in fields:
                    v = tk.StringVar()
                    self.axis_vars[(p, side, f)] = v
                    e = ttk.Entry(inner, textvariable=v, width=6, justify="center")
                    e.grid(row=r, column=c, padx=3, pady=2)
                    v.trace_add("write", lambda *_, _p=p: self._redraw_axis_preview(_p))
                    entries.append(e)
                    c += 1
            self.axis_row_entries[p] = entries
            # Mini axis preview canvas
            pv = tk.Canvas(inner, width=130, height=72, bg=CLR_CARD,
                           highlightthickness=1, highlightbackground="#e5e7eb")
            pv.grid(row=r, column=10, padx=(8, 4), pady=2)
            self.axis_preview_canvases[p] = pv
        # Initial draw after layout settles
        inner.after(100, self._sync_param_rows)
        inner.after(120, self._redraw_all_axis_previews)

    # ── Tab 4: Preview & Run ─────────────────────────────────────────────
    def _build_tab_run(self) -> None:
        t = self.tab_run
        self._csv_banner(t)
        pf = ttk.LabelFrame(t, text="Command Preview  (editable — what you see here is what runs)")
        pf.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(pf, text="Edit the command directly if needed. Use 'Reset to Auto' to "
                           "regenerate it from the current settings.",
                  style="Hint.TLabel").pack(anchor="w", padx=10, pady=(6, 0))
        self.txt_cmd = tk.Text(pf, height=6, wrap="word", font=("Consolas", 9),
                               bg="#0f172a", fg="#e2e8f0", insertbackground="#e2e8f0",
                               relief="flat", padx=8, pady=8)
        self.txt_cmd.pack(fill="x", padx=8, pady=8)

        bar = ttk.Frame(t); bar.pack(fill="x", padx=12, pady=4)
        ttk.Button(bar, text="\u21BB Reset to Auto", command=self._refresh_preview).pack(side="left", padx=3)
        ttk.Button(bar, text="Open Output Folder", command=self._open_outdir).pack(side="left", padx=3)
        self.btn_run = ttk.Button(bar, text="\u25B6  Run Pipeline", style="Run.TButton", command=self._run)
        self.btn_run.pack(side="right", padx=3)
        self.btn_stop = ttk.Button(bar, text="\u25A0 Stop", command=self._stop, state="disabled")
        self.btn_stop.pack(side="right", padx=3)

        lf = ttk.LabelFrame(t, text="Output Log")
        lf.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.txt_log = tk.Text(lf, wrap="word", font=("Consolas", 9),
                               bg="#111827", fg="#d1d5db", relief="flat", padx=8, pady=6)
        slog = ttk.Scrollbar(lf, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=slog.set)
        self.txt_log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        slog.pack(side="right", fill="y", pady=8)

    # ── footer / status ──────────────────────────────────────────────────
    def _build_footer(self) -> None:
        self.status = tk.StringVar(value="Ready.")
        bar = tk.Frame(self.root, bg=CLR_DARK, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self.status, bg=CLR_DARK, fg="#cfe6fb",
                 font=("Segoe UI", 9)).pack(side="left", padx=10)

    # ──────────────────────────────────────────────────────────────────
    # File handling
    # ──────────────────────────────────────────────────────────────────
    def _add_files(self) -> None:
        src = self.var_source.get()
        if src == SRC_CSV:
            ft = [("CSV", "*.csv"), ("All", "*.*")]
        elif src == SRC_XLSX:
            ft = [("Excel", "*.xlsx"), ("All", "*.*")]
        elif src == SRC_PPT:
            ft = [("PNG", "*.png"), ("All", "*.*")]
        else:
            ft = [("Logs", "*.txt *.log"), ("All", "*.*")]
        paths = filedialog.askopenfilenames(title="Select input file(s)", filetypes=ft)
        existing = set(self._all_paths())
        for p in paths:
            if p not in existing:
                self.lst.insert("end", p)
                existing.add(p)
        self._after_files_changed()

    def _folder_patterns(self) -> list[str]:
        """Glob patterns to expand a chosen folder, based on the source type."""
        src = self.var_source.get()
        if src == SRC_CSV:
            return ["*.csv"]
        if src == SRC_XLSX:
            return ["*.xlsx"]
        if src == SRC_PPT:
            return ["*.png"]
        raw = self.var_pattern.get().strip()
        if not raw:
            # Empty field → use the full default so Add Folder still works.
            return ["*.log", "*.txt"]
        pats = [p.strip() for p in re.split(r"[;,]", raw) if p.strip()]
        return pats or ["*.log", "*.txt"]

    def _matches_patterns(self, filepath: str) -> bool:
        """Return True if the filename matches any of the current file mask patterns."""
        name = Path(filepath).name.lower()
        for pat in self._folder_patterns():
            if fnmatch.fnmatch(name, pat.lower()):
                return True
        return False

    def _reapply_pattern(self) -> None:
        """Remove entries in the listbox that no longer match the current pattern,
        then update the multi-file counter and preview.  Only acts on log-source
        mode (CSV / Excel / PPT have fixed patterns)."""
        if self.var_source.get() not in (SRC_LOGS,):
            self._after_files_changed()
            return
        # If the pattern field is blank, treat as "no filter" — keep everything.
        if not self.var_pattern.get().strip():
            self._after_files_changed()
            return
        # Remove non-matching entries (iterate in reverse to keep indices stable)
        for i in range(self.lst.size() - 1, -1, -1):
            entry = self.lst.get(i)
            if Path(entry).is_dir():
                continue          # folder entries are always kept
            if not self._matches_patterns(entry):
                self.lst.delete(i)
        self._after_files_changed()

    def _add_folder(self) -> None:
        d = filedialog.askdirectory(title="Select folder")
        if not d:
            return
        if self.var_source.get() == SRC_PPT:
            # PPT mode needs the charts folder itself, not its individual PNGs.
            self.lst.insert("end", d)
            self._after_files_changed()
            return
        folder = Path(d)
        found: list[Path] = []
        for pat in self._folder_patterns():
            found.extend(sorted(folder.rglob(pat)))
        # de-duplicate while preserving order
        seen, files = set(), []
        for f in found:
            r = f.resolve()
            if r not in seen:
                seen.add(r); files.append(f)
        existing = set(self._all_paths())
        added = 0
        for f in files:
            if str(f) not in existing:
                self.lst.insert("end", str(f))
                added += 1
        if added == 0 and not files:
            messagebox.showinfo(
                "No files",
                f"No files matching {', '.join(self._folder_patterns())} were "
                f"found in:\n{d}")
        self._after_files_changed()

    def _remove_sel(self) -> None:
        for i in reversed(self.lst.curselection()):
            self.lst.delete(i)
        self._after_files_changed()

    def _clear_files(self) -> None:
        self.lst.delete(0, "end")
        self._after_files_changed()

    def _after_files_changed(self) -> None:
        self._autofill_outdir()
        self._update_multi_state()
        self._update_recalc_state()
        self._refresh_axis_file_selector()
        self._refresh_preview()

    def _update_recalc_state(self) -> None:
        """Enable the Re-calculate button when CSV or raw log/txt data is
        available — raw logs are parsed directly via parse_rmt_from_text."""
        has_data = any(
            Path(p).suffix.lower() in (".csv", ".log", ".txt") and Path(p).is_file()
            for p in self._all_paths()
        )
        if not has_data:
            outd = Path(self.var_outdir.get().strip())
            if outd.is_dir():
                has_data = any(
                    (outd / name).exists()
                    for name in ("RMT_Combined_Extended.csv", "RMT_Combined_Similar.csv")
                )
        try:
            self.btn_recalc.configure(state="normal" if has_data else "disabled")
        except AttributeError:
            pass  # button not yet built

    def _update_multi_state(self) -> None:
        n = len(self._all_paths())
        src = self.var_source.get()
        multi_capable = src in (SRC_LOGS, SRC_CSV)
        active = multi_capable and n > 1

        if not multi_capable:
            msg = "Single-source mode (this input type processes one dataset)."
        elif n == 0:
            msg = "No files loaded — add two or more files to enable."
        elif n == 1:
            msg = "1 file loaded — add a second file to enable multi-file options."
        else:
            msg = f"{n} files loaded — choose how to process them:"

        # Colour the count label: grey when inactive, normal dark when active
        self.lbl_count.configure(text=msg,
                                 foreground="#6b7280" if not active else "#111827")

        # Enable / disable radio buttons
        rb_state = "normal" if active else "disabled"
        for w in (self.rb_concat, self.rb_separate):
            w.configure(state=rb_state)

        # Visually dim the whole section when inactive by switching the frame style
        self.multi_frame.configure(
            style="TLabelframe" if active else "Multi.TLabelframe")

    def _on_multimode_change(self) -> None:
        self._refresh_axis_file_selector()
        self._refresh_preview()

    def _all_paths(self) -> list[str]:
        return list(self.lst.get(0, "end"))

    def _autofill_outdir(self) -> None:
        if self.var_outdir.get().strip():
            return
        items = self._all_paths()
        if not items:
            return
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        first = Path(items[0])
        if len(items) == 1 and first.is_file():
            # Single file selected: <filename>_RMT_OUTPUT_<timestamp>
            self.var_outdir.set(str(first.parent / f"{first.stem}_RMT_OUTPUT_{ts}"))
            return
        # Multiple files (or a folder entry): existing naming convention.
        base = first if first.is_dir() else first.parent
        self.var_outdir.set(str(base / f"RMT_Output_{ts}"))

    # ── pickers ──────────────────────────────────────────────────────────
    def _pick_outdir(self) -> None:
        d = filedialog.askdirectory(title="Select output directory")
        if d:
            self.var_outdir.set(d)
            self._refresh_preview()

    def _pick_file(self, var: tk.StringVar, ft) -> None:
        p = filedialog.askopenfilename(filetypes=ft + [("All", "*.*")])
        if p:
            var.set(p)
            self._refresh_preview()

    def _open_outdir(self) -> None:
        d = self.var_outdir.get().strip()
        if d and Path(d).is_dir():
            try:
                subprocess.Popen(["explorer", str(Path(d))])
            except Exception:
                pass
        else:
            messagebox.showinfo("Output", "Output folder does not exist yet.")

    # ── source toggle ──────────────────────────────────────────────────
    def _profile_is_thermal(self) -> bool:
        return PROFILE_BY_LABEL.get(self.var_profile.get()) == PROFILE_THERMAL

    def _on_profile_change(self, event=None) -> None:
        """React to Experiment Profile selection.

        Thermal (DTR) experiments are always parsed from raw logs, so the
        source is forced to logs and a helpful hint is shown. The chosen
        profile only affects log parsing (adds ``--dtr``); all downstream
        CSV/Excel/HTML/JMP generation is shared with the non-thermal path.
        """
        if self._profile_is_thermal():
            self.var_source.set(SRC_LOGS)
            hint = ("DTR (Dynamic Thermal Range): each log yields a Boot-temp RMT "
                    "(Temp Drift RMT Test, before Hit-Enter) and a Run-temp RMT "
                    "(Rank Margin Tool, after Hit-Enter). Boot/Run temps are read "
                    "from 'PHY Temperature1' and snapped to 0 / 90. Add BCRH / BHRC logs.")
            self._on_source_change()
        else:
            hint = "Standard (non-thermal) RMT extraction — original behaviour, unchanged."
        if hasattr(self, "lbl_profile_hint"):
            self.lbl_profile_hint.configure(text=hint)
        self._refresh_preview()

    def _on_source_change(self) -> None:
        src = self.var_source.get()
        hints = {
            SRC_LOGS: "Add boot / MRC *.txt or *.log files, or a folder (uses the pattern).",
            SRC_CSV:  "Add one or more Extended-layout CSV files. Multiple files are "
                      "concatenated and tagged by SourceFile for colour overlay.",
            SRC_XLSX: "Add a single .xlsx (All_RMT sheet) to chart from.",
            SRC_PPT:  "Add the folder that already contains JMP PNG charts; a PPT is built from them.",
        }
        self.lbl_src_hint.configure(text=hints.get(src, ""))
        # Stage selection only meaningful for log extraction
        state = "normal" if src == SRC_LOGS else "disabled"
        for child in self.stage_frame.winfo_children():
            for w in child.winfo_children():
                try:
                    w.configure(state=state)
                except tk.TclError:
                    pass
        self._update_multi_state()
        self._refresh_preview()

    def _on_tmpl_toggle(self) -> None:
        """Enable/disable the PPT template entry and browse button."""
        enabled = self.var_use_tmpl.get()
        state = "normal" if enabled else "disabled"
        try:
            self.ent_tmpl.configure(state=state)
            self.btn_tmpl.configure(state=state)
        except AttributeError:
            pass  # widgets not yet built
        if not enabled:
            # Clear the path so it's not passed to the pipeline accidentally.
            self.var_tmpl.set("")
        self._refresh_preview()

    # ── params helpers ─────────────────────────────────────────────────
    def _set_all_params(self, on: bool) -> None:
        for v in self.param_vars.values():
            v.set(on)
        self._sync_param_rows()
        self._refresh_preview()

    def _reset_params(self) -> None:
        """'Reset Defaults': resets BOTH which parameters are charted AND
        every per-parameter Y-axis value (Min/Max/Inc/+Ref/-Ref) + the
        Defaults section back to the default JSON (jmp_axis_settings.json)
        — this is the single Reset control (the separate 'Reset' button was
        removed as a duplicate of this one)."""
        for p, v in self.param_vars.items():
            v.set(p not in CONSTANT_PARAMS)
        self._sync_param_rows()

        if self._is_separate() and self.per_file_axis:
            # Apply the same reset Chart-enabled state to EVERY file's stored
            # config — the dropdown's file list is left untouched, only the
            # chart-enabled selection is reset back to defaults for all files.
            reset_enabled = self._param_enabled_snapshot()
            for f in list(self.per_file_axis.keys()):
                self.per_file_axis[f]["_chart_enabled"] = dict(reset_enabled)

        self._reset_axis()  # also resets Min/Max/Inc/Ref + Defaults (calls _refresh_preview)

    def _on_param_toggle(self, param: str) -> None:
        self._sync_param_rows(param)
        self._refresh_preview()

    def _sync_param_rows(self, param: str | None = None) -> None:
        """Enable/dim each axis row's fields based on its chart checkbox."""
        if not hasattr(self, "axis_row_entries"):
            return
        targets = [param] if param else list(self.param_vars)
        for p in targets:
            on = self.param_vars[p].get()
            for e in self.axis_row_entries.get(p, []):
                e.configure(state="normal" if on else "disabled")
            lab = self.axis_row_labels.get(p)
            if lab is not None:
                lab.configure(foreground="#111827" if on else "#9ca3af")
            self._redraw_axis_preview(p)

    def _selected_params(self) -> list[str]:
        return [p for p in KNOWN_PARAMS if self.param_vars[p].get()]

    def _validate_axis_values(self) -> list[str]:
        """Return a list of human-readable 'file: Param (fields)' strings for
        every Chart-enabled parameter that is still missing one or more
        Min/Max/Inc/Ref values. An empty list means everything is filled in.

        In separate-file mode this checks EVERY file's own stored per-file
        axis data (falling back to the shared grid for any file that has
        not been through Recalculate yet); otherwise it checks the single
        shared grid."""
        # Make sure the currently-displayed grid is captured first (mirrors
        # the same "capture live edits" snapshot build_commands() performs).
        if self._axis_last_file:
            snap = self._axis_to_dict()
            snap["_chart_enabled"] = self._param_enabled_snapshot()
            self.per_file_axis[self._axis_last_file] = snap

        field_label = {"min": "Min", "max": "Max", "inc": "Inc", "ref_line": "Ref"}
        side_prefix = {"plus": "+", "minus": "-"}

        def _missing_for(params: dict, enabled: dict) -> list[str]:
            out = []
            for p in KNOWN_PARAMS:
                if not enabled.get(p, False):
                    continue
                pc = params.get(p, {})
                missing_fields = [
                    f"{side_prefix[side]}{field_label[f]}"
                    for side in ("plus", "minus")
                    for f in ("min", "max", "inc", "ref_line")
                    if pc.get(side, {}).get(f) is None
                ]
                if missing_fields:
                    out.append(f"{p} ({', '.join(missing_fields)})")
            return out

        problems: list[str] = []
        if self._is_separate() and self.per_file_axis:
            for f in self._all_paths():
                stored = self.per_file_axis.get(f)
                if stored is not None:
                    params = stored.get("parameters", {})
                    enabled = stored.get("_chart_enabled", {})
                else:
                    params = self._axis_to_dict()["parameters"]
                    enabled = self._param_enabled_snapshot()
                missing = _missing_for(params, enabled)
                if missing:
                    problems.append(f"{Path(f).name}: " + "; ".join(missing))
        else:
            data = self._axis_to_dict()
            missing = _missing_for(data["parameters"], self._param_enabled_snapshot())
            if missing:
                problems.append("; ".join(missing))
        return problems

    # ──────────────────────────────────────────────────────────────────
    # Axis settings: load / save / jrp
    # ──────────────────────────────────────────────────────────────────
    def _set_axis_from_dict(self, data: dict) -> None:
        defs = data.get("defaults", {})
        if "inc" in defs: self.def_vars["inc"].set(_fmt(defs["inc"]))
        if "minor_ticks" in defs: self.def_vars["minor_ticks"].set(_fmt(defs["minor_ticks"]))
        if "show_major_grid" in defs: self.def_vars["major_grid"].set(_fmt(defs["show_major_grid"]))
        if "show_minor_grid" in defs: self.def_vars["minor_grid"].set(_fmt(defs["show_minor_grid"]))
        if "ref_line_plus_color" in defs: self.def_vars["plus_color"].set(defs["ref_line_plus_color"])
        if "ref_line_minus_color" in defs: self.def_vars["minus_color"].set(defs["ref_line_minus_color"])
        params = data.get("parameters", {})
        for p in KNOWN_PARAMS:
            pc = params.get(p, {})
            for side in ("plus", "minus"):
                sc = pc.get(side, {})
                for f in ("min", "max", "inc", "ref_line"):
                    if f in sc and (p, side, f) in self.axis_vars:
                        self.axis_vars[(p, side, f)].set(_fmt(sc[f]))
        # Refresh previews after bulk load
        self.root.after(50, self._redraw_all_axis_previews)

    def _load_axis_json(self, path: Path, silent: bool = False) -> None:
        if not path.exists():
            if not silent:
                messagebox.showwarning("Axis", f"File not found:\n{path}")
            return
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            self._set_axis_from_dict(data)
            if not silent:
                self.status.set(f"Loaded axis settings from {path.name}")
        except Exception as e:
            if not silent:
                messagebox.showerror("Axis", f"Could not parse JSON:\n{e}")

    def _browse_axis_json(self) -> None:
        p = filedialog.askopenfilename(title="Load axis JSON",
                                       filetypes=[("JSON", "*.json"), ("All", "*.*")])
        if p:
            self.var_axispath.set(p)
            self._load_axis_json(Path(p))
            self._refresh_preview()

    def _default_axis_dict(self) -> dict:
        """Load the full default axis config (Defaults + per-parameter
        Min/Max/Inc/+Ref/-Ref for every KNOWN_PARAM) from the default JSON
        file on disk (AXIS_CONFIG, i.e. jmp_axis_settings.json). Falls back
        to a minimal built-in default (no per-parameter overrides) if the
        file is missing or unreadable."""
        try:
            return json.loads(Path(AXIS_CONFIG).read_text(encoding="utf-8"))
        except Exception:
            return {
                "_comment": "JMP Chart Axis Settings (generated by rmt_gui.py)",
                "defaults": {
                    "inc": 5, "minor_ticks": 0,
                    "show_major_grid": 1, "show_minor_grid": 1,
                    "ref_line_plus_color": "Medium Light Red",
                    "ref_line_minus_color": "Blue",
                },
                "parameters": {},
            }

    def _axis_to_dict(self) -> dict:
        def _n(s, default=None):
            s = (s or "").strip()
            if s == "":
                return default
            try:
                f = float(s)
                return int(f) if f.is_integer() else f
            except ValueError:
                return default
        out = {
            "_comment": "JMP Chart Axis Settings (generated by rmt_gui.py)",
            "defaults": {
                "inc": _n(self.def_vars["inc"].get(), 5),
                "minor_ticks": _n(self.def_vars["minor_ticks"].get(), 0),
                "show_major_grid": _n(self.def_vars["major_grid"].get(), 1),
                "show_minor_grid": _n(self.def_vars["minor_grid"].get(), 1),
                "ref_line_plus_color": self.def_vars["plus_color"].get(),
                "ref_line_minus_color": self.def_vars["minus_color"].get(),
            },
            "parameters": {},
        }
        for p in KNOWN_PARAMS:
            entry = {}
            for side in ("plus", "minus"):
                sc = {}
                for f in ("min", "max", "inc", "ref_line"):
                    val = _n(self.axis_vars[(p, side, f)].get())
                    if val is not None:
                        sc[f] = val
                if sc:
                    entry[side] = sc
            if entry:
                out["parameters"][p] = entry
        return out

    def _save_axis_json(self) -> None:
        p = filedialog.asksaveasfilename(title="Save axis JSON", defaultextension=".json",
                                         initialfile="jmp_axis_settings.json",
                                         filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            Path(p).write_text(json.dumps(self._axis_to_dict(), indent=2), encoding="utf-8")
            self.var_axispath.set(p)
            self.status.set(f"Saved axis settings to {Path(p).name}")
            self._refresh_preview()
        except Exception as e:
            messagebox.showerror("Axis", f"Could not save:\n{e}")

    def _reset_axis(self) -> None:
        """Reset every per-parameter Y-axis field (Min/Max/Inc/+Ref/-Ref) AND
        the shared Defaults back to the values in the default axis JSON file
        (AXIS_CONFIG / jmp_axis_settings.json) — not just a handful of blank
        fields, so +Ref/-Ref (and every other configured value) come back
        pre-filled from that file, matching 'Reset' == 'reload defaults'."""
        for v in self.axis_vars.values():
            v.set("")
        self.def_vars["inc"].set("5"); self.def_vars["minor_ticks"].set("0")
        self.def_vars["major_grid"].set("1"); self.def_vars["minor_grid"].set("1")
        self.def_vars["plus_color"].set("Medium Light Red")
        self.def_vars["minus_color"].set("Blue")

        default_dict = self._default_axis_dict()
        self._set_axis_from_dict(default_dict)

        if self._is_separate() and self.per_file_axis:
            # Reset EVERY file's stored Defaults + per-parameter Y-axis
            # values back to this SAME default config — the dropdown's file
            # list (and each file's own Chart-enabled selection) is left
            # completely untouched; only the axis VALUES are reset.
            for f in list(self.per_file_axis.keys()):
                existing = self.per_file_axis[f]
                self.per_file_axis[f] = {
                    **default_dict,
                    "_chart_enabled": dict(existing.get("_chart_enabled", {})),
                }
            self._refresh_axis_file_selector()

        self._refresh_preview()

    # ──────────────────────────────────────────────────────────────────
    # Axis preview canvas drawing
    # ──────────────────────────────────────────────────────────────────
    def _redraw_all_axis_previews(self) -> None:
        for p in KNOWN_PARAMS:
            self._redraw_axis_preview(p)
        self._refresh_preview()

    def _redraw_axis_preview(self, param: str) -> None:
        """Draw a horizontal eye-diagram for *param* on its preview canvas.

        Layout (left → right on x-axis):
          m_min ── m_ref ── m_max · 0 · p_min ── p_ref ── p_max
          ← minus half (blue) →   gap  ← plus half  (red)  →

        The eye body opens at the ref lines (tallest) and tapers to a point
        at the outer extremes (min/max), exactly like a signal-eye diagram.
        The inner gap (m_max → p_min) represents the 0-crossing dead-band
        and is shown as a light-grey "eye-opening" zone.
        """
        cv = getattr(self, "axis_preview_canvases", {}).get(param)
        if cv is None:
            return
        cv.delete("all")
        W, H = 130, 72
        PX, PY = 5, 5          # padding

        def _v(side: str, field: str):
            var = self.axis_vars.get((param, side, field))
            try:
                return float(var.get()) if var else None
            except ValueError:
                return None

        p_min = _v("plus",  "min");  p_max = _v("plus",  "max")
        p_ref = _v("plus",  "ref_line")
        m_min = _v("minus", "min");  m_max = _v("minus", "max")
        m_ref = _v("minus", "ref_line")
        inc   = _v("plus",  "inc")

        # Background
        cv.create_rectangle(0, 0, W, H, fill=CLR_CARD, outline="")

        # Dim when the parameter is not selected for charting
        if param in self.param_vars and not self.param_vars[param].get():
            cv.create_rectangle(0, 0, W, H, fill="#f3f4f6", outline="")
            cv.create_text(W // 2, H // 2, text="off", fill="#9ca3af",
                           font=("Segoe UI", 8, "italic"))
            return

        # Validate
        if None in (p_min, p_max, m_min, m_max) or p_max <= p_min or m_max <= m_min:
            cv.create_text(W // 2, H // 2, text="—", fill="#9ca3af",
                           font=("Segoe UI", 9))
            return

        # Full horizontal range: minus outer edge → plus outer edge
        v_lo, v_hi = m_min, p_max
        v_rng = v_hi - v_lo
        if v_rng <= 0:
            return

        draw_w = W - 2 * PX
        y_mid  = H // 2
        h_eye  = (H - 2 * PY) // 2 - 1   # half-height of widest part of eye

        def tx(v: float) -> float:
            """Map a data value to an x-pixel coordinate."""
            return PX + (v - v_lo) / v_rng * draw_w

        # Key x positions
        xA = tx(m_min)                                          # minus outer point
        xB = tx(m_ref) if m_ref is not None else tx(m_min * 0.3)  # minus ref line
        xC = tx(m_max)                                          # minus inner edge
        xD = tx(p_min)                                          # plus  inner edge
        xE = tx(p_ref) if p_ref is not None else tx(p_max * 0.3)  # plus  ref line
        xF = tx(p_max)                                          # plus  outer point

        # Clamp ref-line positions to be within the drawn range
        xB = max(xA, min(xC, xB))
        xE = max(xD, min(xF, xE))

        # ── Minus eye half (left, blue) ──────────────────────────────
        # Shape: tapers from xA (height 0) → widens to full h at xB →
        #        stays full until xC (inner edge, right side is flat wall)
        poly_m = [
            xA, y_mid,               # left apex
            xB, y_mid - h_eye,       # top opens at ref line
            xC, y_mid - h_eye,       # top flat to inner edge
            xC, y_mid + h_eye,       # bottom flat at inner edge
            xB, y_mid + h_eye,       # bottom closes at ref line
        ]
        cv.create_polygon(poly_m, fill="#bfdbfe", outline="#3b82f6",
                          width=1, smooth=False)

        # ── Plus eye half (right, red) ───────────────────────────────
        # Shape: flat wall at xD → wide until xE → tapers to xF (apex)
        poly_p = [
            xD, y_mid - h_eye,       # top flat at inner edge
            xE, y_mid - h_eye,       # top flat to ref line
            xF, y_mid,               # right apex
            xE, y_mid + h_eye,       # bottom closes at ref line
            xD, y_mid + h_eye,       # bottom flat at inner edge
        ]
        cv.create_polygon(poly_p, fill="#fecaca", outline="#ef4444",
                          width=1, smooth=False)

        # ── Eye opening (dead-band between m_max and p_min) ──────────
        if xD > xC + 1:
            cv.create_rectangle(xC, y_mid - h_eye, xD, y_mid + h_eye,
                                 fill="#f3f4f6", outline="#d1d5db", width=1)

        # ── Horizontal centre axis ────────────────────────────────────
        cv.create_line(PX, y_mid, W - PX, y_mid, fill="#cbd5e1", width=1)

        # ── Zero marker (thin dashed vertical) ───────────────────────
        if v_lo < 0 < v_hi:
            x0 = tx(0.0)
            cv.create_line(x0, PY, x0, H - PY,
                           fill="#64748b", width=1, dash=(2, 3))
            cv.create_text(x0, H - 2, text="0", fill="#64748b",
                           font=("Segoe UI", 6), anchor="s")

        # ── Ref-line markers (dashed vertical) ───────────────────────
        if m_ref is not None:
            cv.create_line(xB, PY + 1, xB, H - PY - 1,
                           fill="#1d4ed8", width=1, dash=(4, 2))
        if p_ref is not None:
            cv.create_line(xE, PY + 1, xE, H - PY - 1,
                           fill="#b91c1c", width=1, dash=(4, 2))

        # ── Inc tick marks along the centre axis ─────────────────────
        if inc and inc > 0:
            v, cnt = 0.0, 0
            while v <= p_max + 1e-6 and cnt < 50:
                xt = tx(v)
                cv.create_line(xt, y_mid - 3, xt, y_mid + 3, fill="#475569")
                v += inc; cnt += 1
            v, cnt = -inc, 0
            while v >= m_min - 1e-6 and cnt < 50:
                xt = tx(v)
                cv.create_line(xt, y_mid - 3, xt, y_mid + 3, fill="#475569")
                v -= inc; cnt += 1

        # ── Outer-edge value labels ───────────────────────────────────
        def _lbl(val: float) -> str:
            return str(int(val)) if val == int(val) else f"{val:.1f}"
        cv.create_text(xA + 1, PY + 1, text=_lbl(m_min), fill="#1d4ed8",
                       font=("Segoe UI", 6), anchor="nw")
        cv.create_text(xF - 1, PY + 1, text=_lbl(p_max), fill="#b91c1c",
                       font=("Segoe UI", 6), anchor="ne")

        # ── Ref-line value labels ─────────────────────────────────────
        if m_ref is not None:
            cv.create_text(xB, H - PY, text=_lbl(m_ref), fill="#1d4ed8",
                           font=("Segoe UI", 6), anchor="s")
        if p_ref is not None:
            cv.create_text(xE, H - PY, text=_lbl(p_ref), fill="#b91c1c",
                           font=("Segoe UI", 6), anchor="s")

    # ──────────────────────────────────────────────────────────────────
    # Per-file axis selector (separate mode, multiple files)
    # ──────────────────────────────────────────────────────────────────
    def _axis_file_labels(self, files: list[str]) -> dict[str, str]:
        """Return {unique_display_label: file_path} for the dropdown.

        Plain filenames collide when multiple per-file subfolders each
        contain a same-named CSV (e.g. every Stage-1 output is named
        "RMT_Combined_Extended.csv") — without disambiguation the dropdown
        (and its underlying dict) would silently collapse every file down to
        one entry (last one wins). Disambiguate using the immediate parent
        folder name whenever a bare filename is not unique across the set.
        """
        names = [Path(f).name for f in files]
        dupe_names = {n for n in names if names.count(n) > 1}
        labels: list[str] = []
        for f in files:
            p = Path(f)
            labels.append(f"{p.parent.name}/{p.name}" if p.name in dupe_names else p.name)
        # Final guard in case disambiguated labels still collide.
        seen: dict[str, int] = {}
        unique_labels: list[str] = []
        for lbl in labels:
            seen[lbl] = seen.get(lbl, 0) + 1
            unique_labels.append(lbl if seen[lbl] == 1 else f"{lbl} ({seen[lbl]})")
        return dict(zip(unique_labels, files))

    def _refresh_axis_file_selector(self) -> None:
        """Show/hide and (re)populate the 'Per-file axis values for:' dropdown.

        Only meaningful when processing multiple files SEPARATELY — each file
        gets its own subfolder/report, so its axis Min/Max/Inc can (and
        usually should) differ from the others. The underlying file list is
        NEVER altered here — only the dropdown's display labels/mapping are
        (re)built from whatever files are currently loaded.
        """
        if not hasattr(self, "cbo_axis_file"):
            return
        is_sep = self._is_separate()
        files = self._all_paths() if is_sep else []
        self._axis_file_map = self._axis_file_labels(files)
        values = list(self._axis_file_map.keys())
        self.cbo_axis_file["values"] = values

        if is_sep and values:
            self.axis_file_row.pack(fill="x", padx=10, pady=(0, 6), before=self._axis_psel_ref)
            cur = self.var_axis_file.get()
            if cur not in values:
                cur = values[0]
                self.var_axis_file.set(cur)
                last_file = self._axis_file_map.get(cur)
                self._axis_last_file = last_file
                stored = self.per_file_axis.get(last_file) if last_file else None
                if stored:
                    self._set_axis_from_dict(stored)
                    self._apply_param_enabled_snapshot(stored.get("_chart_enabled", {}))
                    self._redraw_all_axis_previews()
        else:
            self.axis_file_row.pack_forget()
            self._axis_last_file = None

    def _param_enabled_snapshot(self) -> dict[str, bool]:
        """Return the current 'Chart' checkbox state for every known param."""
        return {p: bool(self.param_vars[p].get()) for p in KNOWN_PARAMS}

    def _apply_param_enabled_snapshot(self, snap: dict) -> None:
        """Apply a stored per-file 'Chart' checkbox snapshot to the grid."""
        if not snap:
            return
        for p in KNOWN_PARAMS:
            if p in snap:
                self.param_vars[p].set(bool(snap[p]))
        self._sync_param_rows()

    def _on_axis_file_selected(self, event=None) -> None:
        """Save the currently-displayed grid (Defaults + per-parameter Y-axis
        + which parameters are charted) to the previous file's slot, then
        load the newly-selected file's own stored (or default) values — each
        file has a fully independent Defaults + Parameters + Chart-enabled
        config."""
        new_key = self.var_axis_file.get()
        new_file = self._axis_file_map.get(new_key)
        if self._axis_last_file:
            snap = self._axis_to_dict()
            snap["_chart_enabled"] = self._param_enabled_snapshot()
            self.per_file_axis[self._axis_last_file] = snap
        if new_file:
            stored = self.per_file_axis.get(new_file)
            if stored:
                self._set_axis_from_dict(stored)
                self._apply_param_enabled_snapshot(stored.get("_chart_enabled", {}))
            # else: leave the grid as-is (shared/default values) until the
            # user runs Recalculate for this batch.
        self._axis_last_file = new_file
        self._redraw_all_axis_previews()

    def _rows_for_file(self, f: str) -> list[dict]:
        """Return RMT rows for a single file — CSV rows for .csv, or parsed
        rows straight from a raw .log/.txt file (no Stage-1 extraction needed)."""
        import importlib, csv as _csv
        p = Path(f)
        try:
            if p.suffix.lower() == ".csv":
                with p.open(encoding="utf-8", errors="replace", newline="") as fh:
                    return list(_csv.DictReader(fh))
            pipeline = importlib.import_module("rmt_log_pipeline")
            text = pipeline.read_text_file(p)
            return pipeline.parse_rmt_from_text(text, p.name)
        except Exception:
            return []

    # ──────────────────────────────────────────────────────────────────
    # Re-calculate axis ranges from loaded data
    # ──────────────────────────────────────────────────────────────────
    def _recalculate_axis(self) -> None:
        """Derive Min / Max / Inc from the loaded data files.
        Ref-line values (+Ref / -Ref) are intentionally preserved.

        When "Process each file SEPARATELY" is active with multiple files,
        each file's Min/Max/Inc is computed from ITS OWN data only (not the
        combined set) and stored per-file — switch between them with the
        "Per-file axis values for:" dropdown above the parameter grid.
        """
        import importlib

        try:
            pipeline = importlib.import_module("rmt_log_pipeline")
            _car = pipeline._compute_axis_range
        except Exception:
            messagebox.showerror("Re-calculate",
                                 "Could not import rmt_log_pipeline.\n"
                                 "Ensure it is in the same folder.")
            return

        if self._is_separate():
            files = [f for f in self._all_paths()
                     if Path(f).suffix.lower() in (".csv", ".log", ".txt") and Path(f).is_file()]
            if not files:
                messagebox.showinfo(
                    "Re-calculate",
                    "No CSV / .log / .txt files found in the file list.")
                return

            row_counts: dict[str, int] = {}
            base_defaults = self._axis_to_dict()["defaults"]  # current shared UI defaults
            base_chart_enabled = self._param_enabled_snapshot()  # current shared Chart checkboxes
            for f in files:
                rows = self._rows_for_file(f)
                row_counts[f] = len(rows)
                if not rows:
                    continue
                params_dict: dict = {}
                for p in KNOWN_PARAMS:
                    try:
                        pl_min, pl_max, mi_min, mi_max, inc = _car(rows, p)
                    except Exception:
                        continue
                    ref_p = self.axis_vars[(p, "plus", "ref_line")].get()
                    ref_m = self.axis_vars[(p, "minus", "ref_line")].get()
                    entry = {
                        "plus":  {"min": pl_min, "max": pl_max, "inc": float(inc)},
                        "minus": {"min": mi_min, "max": mi_max, "inc": float(inc)},
                    }
                    if ref_p.strip():
                        entry["plus"]["ref_line"] = float(ref_p)
                    if ref_m.strip():
                        entry["minus"]["ref_line"] = float(ref_m)
                    params_dict[p] = entry
                # Each file gets its OWN independent Defaults + Parameters +
                # Chart-enabled config (starts as a copy of the current
                # shared UI state, but is stored/edited separately from here
                # on). The file list itself is never modified by Recalculate.
                existing = self.per_file_axis.get(f, {})
                self.per_file_axis[f] = {
                    "_comment": "JMP Chart Axis Settings (generated by rmt_gui.py)",
                    "defaults": dict(existing.get("defaults", base_defaults)),
                    "parameters": params_dict,
                    "_chart_enabled": dict(existing.get("_chart_enabled", base_chart_enabled)),
                }

            self._refresh_axis_file_selector()
            cur_file = self._axis_file_map.get(self.var_axis_file.get())
            if cur_file and cur_file in self.per_file_axis:
                stored = self.per_file_axis[cur_file]
                self._set_axis_from_dict(stored)
                self._apply_param_enabled_snapshot(stored.get("_chart_enabled", {}))
            self._redraw_all_axis_previews()

            n_ok = sum(1 for f in files if row_counts.get(f, 0) > 0)
            self.status.set(
                f"Re-calculated per-file axis ranges for {n_ok}/{len(files)} file(s). "
                f"+Ref / \u2212Ref preserved. Files list unchanged.")
            details = "\n".join(
                f"  {Path(f).name}: {row_counts.get(f, 0)} row(s)" for f in files)
            messagebox.showinfo(
                "Re-calculate complete (per file)",
                "Computed unique Min/Max/Inc per file:\n\n" + details + "\n\n"
                "Each file also has its own independent Defaults section "
                "(Inc/Minor ticks/Major/Minor grid/ref colors) and its own "
                "Chart-enabled selection \u2014 edit them after selecting a file "
                "from the dropdown above.\n"
                "+ Ref / \u2212 Ref values were not changed.")
            return

        # ── Combined mode (concat / single file / non-separate) ──────────
        import csv as _csv

        # Collect CSV candidate files from the listbox
        csv_files: list[Path] = []
        log_files: list[Path] = []
        for item in self._all_paths():
            p = Path(item)
            if not p.is_file():
                continue
            if p.suffix.lower() == ".csv":
                csv_files.append(p)
            elif p.suffix.lower() in (".log", ".txt"):
                log_files.append(p)

        # If no CSVs in listbox, look for Combined Extended in output dir
        if not csv_files and not log_files:
            outd = Path(self.var_outdir.get().strip())
            if outd.is_dir():
                for cand in ["RMT_Combined_Extended.csv", "RMT_Combined_Similar.csv"]:
                    fp = outd / cand
                    if fp.exists():
                        csv_files.append(fp); break

        if not csv_files and not log_files:
            messagebox.showinfo(
                "Re-calculate",
                "No CSV / .log / .txt data found.\n\n"
                "Load CSV files on Tab 1 (Data Source → Existing CSV),\n"
                "or add your .log/.txt files directly — Re-calculate parses\n"
                "them without needing Stage 1 extraction first.")
            return

        # Load rows from all CSVs and/or raw logs
        rows: list[dict] = []
        for fp in csv_files:
            try:
                with fp.open(encoding="utf-8", errors="replace", newline="") as fh:
                    rows.extend(list(_csv.DictReader(fh)))
            except Exception as e:
                messagebox.showerror("Re-calculate", f"Could not read {fp.name}:\n{e}")
                return
        for fp in log_files:
            rows.extend(self._rows_for_file(str(fp)))

        source_files = csv_files + log_files
        if not rows:
            messagebox.showinfo("Re-calculate", "No RMT rows found in the selected file(s).")
            return

        updated = 0
        for p in KNOWN_PARAMS:
            try:
                pl_min, pl_max, mi_min, mi_max, inc = _car(rows, p)
            except Exception:
                continue

            # Only update min/max/inc — preserve ref_line
            for side, field, val in [
                ("plus",  "min", pl_min), ("plus",  "max", pl_max),
                ("minus", "min", mi_min), ("minus", "max", mi_max),
                ("plus",  "inc", float(inc)), ("minus", "inc", float(inc)),
            ]:
                if (p, side, field) in self.axis_vars:
                    self.axis_vars[(p, side, field)].set(_fmt(val))
            updated += 1

        self._redraw_all_axis_previews()
        self.status.set(
            f"Re-calculated axis ranges for {updated} parameters from "
            f"{len(source_files)} file(s).  +Ref / \u2212Ref preserved.")
        messagebox.showinfo(
            "Re-calculate complete",
            f"Updated Min / Max / Inc for {updated} parameters.\n"
            f"Source: {', '.join(f.name for f in source_files[:3])}"
            + (" \u2026" if len(source_files) > 3 else "") +
            "\n\n+ Ref and \u2212 Ref values were not changed.")

    def _load_jrp(self) -> None:
        p = filedialog.askopenfilename(title="Load reference .jrp",
                                       filetypes=[("JMP journal/script", "*.jrp *.jsl"), ("All", "*.*")])
        if not p:
            return
        try:
            text = Path(p).read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            messagebox.showerror(".jrp", f"Could not read file:\n{e}")
            return
        cols = parse_jrp_axis(text)
        if not cols:
            messagebox.showwarning(".jrp", "No ScaleBox axis blocks were found in this file.")
            return
        applied = 0
        for col, vals in cols.items():
            m = re.match(r"^(.*?)([+\-])$", col.strip())
            if not m:
                continue
            base, sign = m.group(1), m.group(2)
            if base not in KNOWN_PARAMS:
                continue
            side = "plus" if sign == "+" else "minus"
            for f in ("min", "max", "inc", "ref_line"):
                if vals.get(f) is not None and (base, side, f) in self.axis_vars:
                    self.axis_vars[(base, side, f)].set(_fmt(vals[f]))
            applied += 1
        self.var_useaxis.set(True)
        self.status.set(f"Imported axis scales from {Path(p).name} ({applied} columns).")
        messagebox.showinfo(".jrp imported",
                            f"Imported axis scales for {applied} column(s) from:\n{Path(p).name}")
        self._refresh_preview()

    # ──────────────────────────────────────────────────────────────────
    # Command building
    # ──────────────────────────────────────────────────────────────────
    def _write_axis_temp(self) -> str | None:
        """Persist current axis grid to the configured JSON path; return it."""
        try:
            path = Path(self.var_axispath.get().strip() or AXIS_CONFIG)
            path.write_text(json.dumps(self._axis_to_dict(), indent=2), encoding="utf-8")
            return str(path)
        except Exception:
            return None

    def _axis_path_for_file(self, f: str, shared_axis_path: str, persist: bool) -> str:
        """Return the --jmp-axis-config path to use for a single file.

        If per-file axis data was computed for *f* (via Recalculate in
        "separate" mode), its own independent Defaults + Parameters config is
        written to its own JSON (named after the file), and that path is
        returned. Otherwise the shared axis_path is used unchanged (backward
        compatible).
        """
        stored = self.per_file_axis.get(f)
        if not stored:
            return shared_axis_path
        base = Path(shared_axis_path or AXIS_CONFIG)
        safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(f).stem)
        path = base.with_name(f"{base.stem}__{safe_stem}{base.suffix}")
        if persist:
            try:
                # "_chart_enabled" is GUI-internal bookkeeping (which params
                # are ticked) — not part of the axis JSON schema consumed by
                # the pipeline, so it is excluded from the written file.
                data = {k: v for k, v in stored.items() if k != "_chart_enabled"}
                path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            except Exception:
                return shared_axis_path
        return str(path)

    def _chart_fields_for_file(self, f: str) -> str | None:
        """Return this file's own '--chart-fields' value if per-file Chart
        checkbox data was computed for it (Recalculate in separate mode);
        None if it should fall back to the currently-displayed selection."""
        stored = self.per_file_axis.get(f)
        if not stored or "_chart_enabled" not in stored:
            return None
        enabled = stored["_chart_enabled"]
        return ",".join(p for p in KNOWN_PARAMS if enabled.get(p, False))

    def _build_single_command(self, files: list[str], outdir: str,
                              axis_path: str, fields_override: str | None = None) -> list[str]:
        src    = self.var_source.get()
        cmd = [PYTHON_EXE, PIPELINE_SCRIPT]
        fields = fields_override if fields_override is not None else ",".join(self._selected_params())

        def _axis():
            if self.var_useaxis.get() and axis_path:
                return ["--jmp-axis-config", axis_path]
            return []

        if src == SRC_PPT:
            charts = files[0] if files else ""
            cmd += ["--ppt-from-charts", charts, "--outdir", outdir,
                    "--ppt-name", self.var_ppt.get()]
            if self.var_tmpl.get().strip():
                cmd += ["--ppt-template", self.var_tmpl.get().strip()]
            return cmd

        if src == SRC_XLSX:
            xlsx = files[0] if files else ""
            cmd += ["--jmp-from-excel", xlsx, "--outdir", outdir,
                    "--jmp-exe", self.var_jmp.get().strip(),
                    "--excel-name", self.var_excel.get(),
                    "--ppt-name", self.var_ppt.get()]
            if fields:
                cmd += ["--chart-fields", fields]
            if self.var_tmpl.get().strip():
                cmd += ["--ppt-template", self.var_tmpl.get().strip()]
            cmd += _axis()
            if self.var_jslonly.get():
                cmd += ["--jmp-jsl-only"]
            return cmd

        if src == SRC_CSV:
            cmd += ["--jmp-from-csv", *files, "--outdir", outdir,
                    "--jmp-exe", self.var_jmp.get().strip(),
                    "--excel-name", self.var_excel.get(),
                    "--ppt-name", self.var_ppt.get()]
            if fields:
                cmd += ["--chart-fields", fields]
            if self.var_tmpl.get().strip():
                cmd += ["--ppt-template", self.var_tmpl.get().strip()]
            cmd += _axis()
            if self.var_jslonly.get():
                cmd += ["--jmp-jsl-only"]
            return cmd

        # SRC_LOGS — staged extraction
        cmd += ["--input", *files, "--pattern", self.var_pattern.get(),
                "--outdir", outdir,
                "--excel-name", self.var_excel.get(),
                "--ppt-name", self.var_ppt.get()]
        if self._profile_is_thermal():
            cmd += ["--dtr"]
        cmd += [] if self.var_approval.get() else ["--no-ask-chart-approval"]
        stage = self.var_stage.get()
        if stage in (1, 2):
            cmd += ["--chart-fields", ""]
        else:
            cmd += ["--chart-fields", fields, "--generate-jmp-charts",
                    "--jmp-exe", self.var_jmp.get().strip()]
            cmd += _axis()
            if stage == 4 and self.var_tmpl.get().strip():
                cmd += ["--ppt-template", self.var_tmpl.get().strip()]
            if self.var_jslonly.get():
                cmd += ["--jmp-jsl-only"]
        return cmd

    def _is_separate(self) -> bool:
        """True when each file should be processed into its own subfolder."""
        return (self.var_source.get() in (SRC_LOGS, SRC_CSV)
                and self.var_multimode.get() == "separate"
                and len(self._all_paths()) > 1)

    def _unique_stems(self, files: list[str]) -> dict[str, str]:
        """Return {file_path: disambiguated_stem} for per-file output folders.

        Plain stems collide when multiple per-file subfolders each contain a
        same-named CSV (e.g. every Stage-1 output is named
        "RMT_Combined_Extended.csv") — without disambiguation every file's
        separate-mode output subfolder would collapse to the SAME name, so
        a later file's run silently overwrites an earlier one's charts/PPT
        in that shared folder instead of getting its own. Disambiguate using
        the immediate parent folder name whenever a bare stem is not unique
        across the set (mirrors `_axis_file_labels`'s approach).
        """
        stems = [Path(f).stem for f in files]
        dupes = {s for s in stems if stems.count(s) > 1}
        result: dict[str, str] = {}
        seen: dict[str, int] = {}
        for f in files:
            p = Path(f)
            stem = f"{p.parent.name}_{p.stem}" if p.stem in dupes else p.stem
            seen[stem] = seen.get(stem, 0) + 1
            result[f] = stem if seen[stem] == 1 else f"{stem}_{seen[stem]}"
        return result

    def build_commands(self, persist_axis: bool = False) -> list[tuple[str, list[str]]]:
        """Return a list of (label, argv) commands to run sequentially.

        Concat / single -> one command for all files, into ``outdir`` as-is
        (existing ``RMT_Output_<timestamp>`` naming from `_autofill_outdir`).
        Separate -> one command per file, each into its own subfolder named
        ``<file-stem>_RMT_OUTPUT_<timestamp>`` (mirrors the single-file naming
        convention), all subfolders sharing one timestamp per batch. Each
        file uses ITS OWN independent Defaults + Parameters axis JSON, and
        its own Chart-enabled parameter selection, when per-file axis data
        has been computed via Recalculate (`self.per_file_axis`); otherwise
        it falls back to the shared axis config / current selection,
        unchanged from prior behavior.
        """
        files  = self._all_paths()
        outdir = self.var_outdir.get().strip()
        axis_path = self.var_axispath.get().strip()
        if persist_axis and self.var_useaxis.get():
            axis_path = self._write_axis_temp() or axis_path

        # Capture any live edits (Defaults + per-parameter Y-axis + Chart
        # checkboxes) to the currently-displayed file's grid before
        # resolving per-file axis paths / chart fields below.
        if self._axis_last_file:
            snap = self._axis_to_dict()
            snap["_chart_enabled"] = self._param_enabled_snapshot()
            self.per_file_axis[self._axis_last_file] = snap

        if self._is_separate():
            from datetime import datetime
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            stems = self._unique_stems(files)
            cmds = []
            for f in files:
                sub = str(Path(outdir) / f"{stems[f]}_RMT_OUTPUT_{ts}") if outdir else ""
                file_axis_path = self._axis_path_for_file(f, axis_path, persist_axis)
                file_fields = self._chart_fields_for_file(f)
                cmds.append((stems[f],
                            self._build_single_command([f], sub, file_axis_path, file_fields)))
            return cmds
        return [("all", self._build_single_command(files, outdir, axis_path))]

    def _quote(self, parts: list[str]) -> str:
        out = []
        for p in parts:
            out.append(f'"{p}"' if (" " in p or not p) else p)
        # readable multi-line layout: break before each leading-dash flag
        line, lines = [], []
        for tok in out:
            if tok.startswith("--") and line:
                lines.append(" ".join(line)); line = [tok]
            else:
                line.append(tok)
        if line:
            lines.append(" ".join(line))
        return " ^\n    ".join(lines)

    def _refresh_preview(self) -> None:
        try:
            cmds = self.build_commands(persist_axis=False)
            self.txt_cmd.delete("1.0", "end")
            if len(cmds) == 1:
                self.txt_cmd.insert("1.0", self._quote(cmds[0][1]))
            else:
                blocks = []
                for i, (label, argv) in enumerate(cmds, 1):
                    blocks.append(f"REM [{i}/{len(cmds)}] {label}\n" + self._quote(argv))
                self.txt_cmd.insert("1.0", "\n\n".join(blocks))
        except Exception as e:
            self.txt_cmd.delete("1.0", "end")
            self.txt_cmd.insert("1.0", f"[preview error] {e}")

    # ──────────────────────────────────────────────────────────────────
    # Run
    # ──────────────────────────────────────────────────────────────────
    def _validate(self) -> bool:
        if not self._all_paths():
            messagebox.showwarning("Missing input", "Add at least one input file or folder.")
            return False
        if not self.var_outdir.get().strip():
            messagebox.showwarning("Missing output", "Choose an output directory.")
            return False
        src = self.var_source.get()
        needs_jmp = (src in (SRC_CSV, SRC_XLSX)) or \
                    (src == SRC_LOGS and self.var_stage.get() >= 3)
        if needs_jmp and not self.var_jslonly.get():
            jmp = self.var_jmp.get().strip()
            if not jmp or not Path(jmp).exists():
                if not messagebox.askyesno(
                        "JMP not found",
                        "JMP executable was not found. Continue anyway?"):
                    return False
        return True

    def _commands_from_preview(self) -> list[list[str]]:
        """Parse the (possibly hand-edited) preview text into one or more argv
        lists. Commands are delimited by lines starting with REM, or by the
        appearance of a fresh python-exe token."""
        raw = self.txt_cmd.get("1.0", "end")
        # Drop REM comment lines, collapse caret line-continuations.
        kept = []
        for ln in raw.splitlines():
            if ln.strip().upper().startswith("REM"):
                kept.append("\x00")        # command separator marker
            else:
                kept.append(ln)
        text = "\n".join(kept).replace("^\n", " ").replace("\n", " ")

        def _clean(tokens: list[str]) -> list[str]:
            out = []
            for tok in tokens:
                if len(tok) >= 2 and tok[0] == tok[-1] == '"':
                    tok = tok[1:-1]
                out.append(tok)
            return out

        commands: list[list[str]] = []
        for chunk in text.split("\x00"):
            chunk = chunk.strip()
            if not chunk:
                continue
            toks = _clean(shlex.split(chunk, posix=False))
            if toks:
                commands.append(toks)
        return commands

    def _run(self) -> None:
        if self._proc is not None:
            messagebox.showinfo("Running", "A pipeline run is already in progress.")
            return
        if not self._validate():
            return
        # Persist axis edits so --jmp-axis-config (in the preview) points to fresh data.
        if self.var_useaxis.get():
            self._write_axis_temp()
        commands = self._commands_from_preview()
        if not commands:
            messagebox.showwarning("Empty command", "The command preview is empty.")
            return
        self._active_stage = None
        self.txt_log.delete("1.0", "end")
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.status.set(f"Running pipeline ({len(commands)} command(s))\u2026")
        self.nb.select(self.tab_run)
        threading.Thread(target=self._worker, args=(commands,), daemon=True).start()

    def _run_stage(self, n: int) -> None:
        """Run the boot-log pipeline up to stage *n* only, then stop."""
        if self._proc is not None:
            messagebox.showinfo("Running", "A pipeline run is already in progress.")
            return
        if self.var_source.get() != SRC_LOGS:
            messagebox.showinfo(
                "Stage runner",
                "Stage-by-stage running applies to the boot-log source.\n"
                "Use Run Pipeline for CSV / Excel / PNG inputs.")
            return
        if not self._validate():
            return
        prev = self.var_stage.get()
        self.var_stage.set(n)
        if self.var_useaxis.get():
            self._write_axis_temp()
        self._refresh_preview()
        commands = self._commands_from_preview()
        self.var_stage.set(prev)
        self._refresh_preview()
        if not commands:
            messagebox.showwarning("Empty command", "The command preview is empty.")
            return
        self._active_stage = n
        self.txt_log.delete("1.0", "end")
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.status.set(f"Running stage {n}\u2026")
        self.nb.select(self.tab_run)
        threading.Thread(target=self._worker, args=(commands,), daemon=True).start()


    def _worker(self, commands: list[list[str]]) -> None:
        # Force UTF-8 I/O in the child so emoji / unicode prints don't crash
        # under the default Windows cp1252 codec when stdout is a pipe.
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        self._stop_requested = False
        total = len(commands)
        ok = True
        LOG.info("Worker starting: %d command(s) to run", total)
        try:
            for idx, cmd in enumerate(commands, 1):
                if self._stop_requested:
                    LOG.info("Worker stopped by user before command %d/%d", idx, total)
                    break
                LOG.info("[cmd %d/%d] %s", idx, total, " ".join(cmd))
                if total > 1:
                    self._log_q.put(f"\n{'='*60}\n[{idx}/{total}] {' '.join(cmd)}\n{'='*60}\n")
                else:
                    self._log_q.put("$ " + " ".join(cmd) + "\n\n")
                self._proc = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, cwd=str(HERE), env=env,
                    encoding="utf-8", errors="replace")
                assert self._proc.stdout is not None
                for line in self._proc.stdout:
                    LOG.debug("[subprocess %d/%d] %s", idx, total, line.rstrip())
                    self._log_q.put(line)
                self._proc.wait()
                rc = self._proc.returncode
                LOG.info("[cmd %d/%d] exited with code %s", idx, total, rc)
                self._log_q.put(f"\n[exit code {rc}]\n")
                if rc != 0:
                    ok = False
                    LOG.error("[cmd %d/%d] failed (rc=%s) — stopping remaining commands", idx, total, rc)
                    self._log_q.put("\n[stopping: previous command failed]\n")
                    break
        except Exception as e:
            ok = False
            LOG.exception("Worker raised an exception")
            self._log_q.put(f"\n[error] {e}\n")
        finally:
            self._proc = None
            self._run_ok = ok and not self._stop_requested
            LOG.info("Worker finished: ok=%s stop_requested=%s", ok, self._stop_requested)
            self._log_q.put("__DONE__")
            self._log_q.put("__DONE__")

    def _stop(self) -> None:
        self._stop_requested = True
        if self._proc and self._proc.poll() is None:
            try:
                self._proc.terminate()
                self._log("\n[stopped by user]\n")
            except Exception:
                pass

    def _on_run_complete(self) -> None:
        """Mark stage complete, refresh recalc/eye-diagram state, prompt next step."""
        n = self._active_stage
        if n is not None and self._run_ok:
            # All stages up to n are now satisfied
            self._completed_stages.update(range(1, n + 1))
            self.status.set(f"Stage {n} complete.")
            self._update_recalc_state()
            if n == 1:
                # Stage 1 produced the CSV — switch downstream to reuse it so
                # later stages never re-extract logs, then unlock tabs 2–4.
                csvs = self._find_stage1_csvs()
                if csvs and not self._tabs_unlocked:
                    self.var_source.set(SRC_CSV)
                    self.lst.delete(0, "end")
                    for c in csvs:
                        self.lst.insert("end", c)
                    self._update_recalc_state()
                self._unlock_tabs(csvs)
                self.status.set("Stage 1 complete \u2014 CSV ready. "
                                "Re-calculate axis on JMP Axis Settings, then Run.")
                self.nb.select(self.tab_axis)
                self.root.after(150, self._redraw_all_axis_previews)
        else:
            self.status.set("Done." if self._run_ok else "Stopped / failed.")
        self._active_stage = None

    def _drain_log(self) -> None:
        try:
            while True:
                item = self._log_q.get_nowait()
                if item == "__DONE__":
                    self._proc = None
                    self.btn_run.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self._on_run_complete()
                else:
                    self._log(item)
        except queue.Empty:
            pass
        self.root.after(120, self._drain_log)

    def _log(self, text: str) -> None:
        self.txt_log.insert("end", text)
        self.txt_log.see("end")

    # ──────────────────────────────────────────────────────────────────
    def run(self) -> None:
        self._refresh_preview()
        self.root.mainloop()


# Instrument every method above with detailed entry/exit/exception debug
# logging (see _instrument_methods near the top of the file).
RmtGuiApp = _instrument_methods(RmtGuiApp)


# ──────────────────────────────────────────────────────────────────────────
def _screenshot(out_path: str) -> None:
    """Render the window once and save a PNG preview (best-effort)."""
    app = RmtGuiApp()
    # populate a representative example so the preview looks real
    app.lst.insert("end", r"C:/logs/boot_mrc_4800_G2.txt")
    app.lst.insert("end", r"C:/logs/boot_mrc_4800_G4.txt")
    app.var_outdir.set(r"C:/logs/RMT_Output")
    app._refresh_preview()
    app.root.update_idletasks()
    app.root.update()
    app.root.after(500, lambda: _grab(app, out_path))
    app.root.mainloop()


def _grab(app: "RmtGuiApp", out_path: str) -> None:
    try:
        from PIL import ImageGrab  # type: ignore
        app.root.update_idletasks()
        x = app.root.winfo_rootx(); y = app.root.winfo_rooty()
        w = app.root.winfo_width(); h = app.root.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        img.save(out_path)
        print(f"Saved preview: {out_path}")
    except Exception as e:
        print(f"[screenshot failed] {e}")
    finally:
        app.root.destroy()


if __name__ == "__main__":
    if "--screenshot" in sys.argv:
        i = sys.argv.index("--screenshot")
        target = sys.argv[i + 1] if i + 1 < len(sys.argv) else "rmt_gui_preview.png"
        _screenshot(target)
    else:
        RmtGuiApp().run()
