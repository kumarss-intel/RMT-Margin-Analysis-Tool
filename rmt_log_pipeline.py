#!/usr/bin/env python3
"""
Parse RMT blocks from MRC TXT logs and generate:
1) CSV outputs (combined + per source file),
2) Excel workbook with multiple sheets,
3) PPT summary deck.

Expected RMT block format in logs:

START_RMT
Params:  RecEnDelay TxDqsDelay RxDqsDelay TxDqDelay RxDqVrefByte TxVref ClkGrpPi CmdVref
Mc0.C0.R0: -32.0 32.0 -31.0 31.0 ...
...

Usage:
  python rmt_log_pipeline.py --input C:/logs --outdir C:/logs/rmt_out
"""

from __future__ import annotations

import argparse
import csv
import html as _html
import re
import statistics
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font

from rmt_project import (
    BUILTIN_PARAM_ALIASES,
    TOOL_NAME,
    TOOL_SUBTITLE,
    TOOL_VERSION,
    get_project,
    load_projects,
    project_label,
    project_param_aliases,
    resolve_project_key,
)

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches
except ImportError:  # pragma: no cover - optional dependency at runtime
    Presentation = None
    CategoryChartData = None
    XL_CHART_TYPE = None
    Inches = None

PARAMS = [
    "RecEnDelay",
    "TxDqsDelay",
    "RxDqsDelay",
    "TxDqDelay",
    "RxDqVrefByte",
    "TxVref",
    "ClkGrpPi",
    "CmdVref",
]

CSV_COLUMNS = [
    "Params",
    "BootTemp",
    "RunTemp",
    "RecEnDelay-",
    "RecEnDelay+",
    "TxDqsDelay-",
    "TxDqsDelay+",
    "RxDqsDelay-",
    "RxDqsDelay+",
    "TxDqDelay-",
    "TxDqDelay+",
    "RxDqVrefByte-",
    "RxDqVrefByte+",
    "TxVref-",
    "TxVref+",
    "ClkGrpPi-",
    "ClkGrpPi+",
    "CmdVref-",
    "CmdVref+",
]

EXTENDED_COLUMNS = [
    "SourceFile",
    "Frequency",
    "Gear",
    "BlockIndex",
] + CSV_COLUMNS

FLOAT_PATTERN = re.compile(r"-?\d+(?:\.\d+)?")
# Rank rows (NVL: "Mc0.C0.R0") and per-byte rows (WCL: "Mc0.C0.B0.R0").
ROW_PATTERN = re.compile(r"^(Mc\d+\.C\d+(?:\.B\d+)?\.R\d+)\s*:\s*(.+)$")
HEADER_PATTERN = re.compile(r"^Params\s*:\s*(.+)$", re.IGNORECASE)

# Header alias map used by the parsers; main() extends it with the selected
# project's aliases from projects.json.
_ACTIVE_PARAM_ALIASES: dict[str, str] = dict(BUILTIN_PARAM_ALIASES)
# Project label shown in the HTML / PPT reports; set by main() from --project.
_ACTIVE_PROJECT_NAME: str | None = None


def set_param_aliases(aliases: dict[str, str] | None) -> None:
    """Replace the START_RMT header alias map (builtin aliases always kept)."""
    _ACTIVE_PARAM_ALIASES.clear()
    _ACTIVE_PARAM_ALIASES.update(BUILTIN_PARAM_ALIASES)
    _ACTIVE_PARAM_ALIASES.update(aliases or {})


def _header_param_order(header_line: str | None) -> list[str | None]:
    """Canonical parameter per column pair of a ``Params:`` header.

    Unknown names map to ``None`` (their values are skipped). Returns the
    fixed PARAMS order when there is no usable header.
    """
    if not header_line:
        return list(PARAMS)
    m = HEADER_PATTERN.match(header_line.strip())
    if not m:
        return list(PARAMS)
    names = m.group(1).split()
    order: list[str | None] = []
    for name in names:
        canon = _ACTIVE_PARAM_ALIASES.get(name, name)
        order.append(canon if canon in PARAMS else None)
    return order if any(order) else list(PARAMS)


def _assign_param_values(record: dict[str, Any], values: list[float],
                         order: list[str | None]) -> bool:
    """Fill ``<param>-`` / ``<param>+`` from a row's value pairs.

    Returns False when the row holds fewer value pairs than the header.
    """
    if len(values) < 2 * len(order):
        return False
    for idx, param in enumerate(order):
        if param is None:
            continue
        record[f"{param}-"] = values[2 * idx]
        record[f"{param}+"] = values[2 * idx + 1]
    return True
# Matches "Setting boot frequency to 4800" (KIRK/legacy format)
FREQ_PATTERN = re.compile(r"Setting\s+boot\s+frequency\s+to\s+(\d+)", re.IGNORECASE)
# Matches "Requested/actual ratio 144/144, Frequency=4800, GearMode=1" (both logs)
# This is the most accurate per-block source — updates both freq and gear.
FREQ_GEAR_PATTERN = re.compile(
    r"Requested/actual ratio[^,]+,\s*Frequency=(\d+),\s*GearMode=(\d+)",
    re.IGNORECASE,
)
# Matches "Setting gear ratio to 4" (KIRK) or "--> Gear: 2" (MiniBIOS interpreter)
GEAR_PATTERN = re.compile(
    r"Setting\s+gear\s+ratio\s+to\s+(\d+)"   # KIRK: explicit gear setting
    r"|-+>\s*Gear\s*:\s*(\d+)",               # MiniBIOS interpreter: --> Gear: 2
    re.IGNORECASE,
)
BOOT_TEMP_PATTERN = re.compile(
    r"(?:BootTemp|Boot\s*Temp|Setting\s*boot\s*temp(?:erature)?\s*to)\s*[:=]?\s*(-?\d+(?:\.\d+)?)",
    re.IGNORECASE,
)
# NOTE: rtt\s*wr\s*= intentionally REMOVED — it matched ODT resistance values
# (e.g. "RttWr=240") which are not temperature readings.
RUN_TEMP_PATTERN = re.compile(
    r"(?:RunTemp|Run\s*Temp|Run\s+Temperature|Setting\s*run\s*temp(?:erature)?\s*to)\s*[:=]?\s*(-?\d+(?:\.\d+)?)",
    re.IGNORECASE,
)
# Bounds the RMT parsing window to the actual "Rank Margin Tool" GREEN MRC task
# execution, so any stray RMT-looking blocks outside the task (e.g. from a
# prior/next task, or leftover debug prints) are ignored.
RMT_TASK_START_PATTERN = re.compile(
    r"GREEN\s+MRC\s+task\s*(?:--+\s*)?Rank\s+Margin\s+Tool\s*--+\s*Started",
    re.IGNORECASE,
)
RMT_TASK_END_PATTERN = re.compile(
    r"GREEN\s+MRC\s+task\s*(?:--+\s*)?Rank\s+Margin\s+Tool\s*--+\s*SUCCEEDED",
    re.IGNORECASE,
)

# ── DTR (Dynamic Thermal Range) experiment markers ───────────────────────
# A DTR log captures TWO RMT data sets for a single frequency/gear:
#   1. Boot-temperature RMT, under the "Temp Drift RMT Test" GREEN MRC task,
#      captured BEFORE the "Hit Enter" pause (operator changes chamber temp).
#   2. Run-temperature RMT, under the "Rank Margin Tool" GREEN MRC task,
#      captured AFTER the "Hit Enter" pause (chamber now at the run temp).
# Two files are collected per freq/gear: BCRH (Boot Cold Run Hot) and
# BHRC (Boot Hot Run Cold).
DTR_BOOT_TASK_PATTERN = re.compile(
    r"GREEN\s+MRC\s+task\s*(?:--+\s*)?Temp\s+Drift\s+RMT\s+Test\s*--+\s*Started",
    re.IGNORECASE,
)
DTR_RUN_TASK_PATTERN = re.compile(
    r"GREEN\s+MRC\s+task\s*(?:--+\s*)?Rank\s+Margin\s+Tool\s*--+\s*Started",
    re.IGNORECASE,
)
DTR_HIT_ENTER_PATTERN = re.compile(r"Hit\s+Enter", re.IGNORECASE)
# "MemSs[0] PHY Temperature1: 90 C" → captures the numeric temperature.
PHY_TEMP_PATTERN = re.compile(
    r"PHY\s+Temperature1\s*:\s*(-?\d+(?:\.\d+)?)", re.IGNORECASE
)



@dataclass
class ParseContext:
    frequency: int | None = None
    gear: int | None = None
    boot_temp: float | None = None
    run_temp: float | None = None



def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", name)
    return cleaned[:31] if len(cleaned) > 31 else cleaned



def read_text_file(path: Path) -> str:
    encodings = ["utf-8", "utf-16", "latin-1"]
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")



def parse_rmt_from_text(text: str, source_name: str) -> list[dict[str, Any]]:
    # ── Bound RMT block capture to the "Rank Margin Tool" GREEN MRC task ──
    # Real logs typically contain the Started/SUCCEEDED marker pair once per
    # frequency/gear/temperature iteration (i.e. multiple times per file), and
    # Frequency/Gear/Temp context lines almost always appear *before* the
    # first "Started" marker (set during cold boot). So rather than slicing
    # the file down to a single window (which would drop every iteration but
    # the first, and lose all preceding frequency/gear context), we scan the
    # whole file as before for context, and only gate START_RMT *block
    # capture* on whether the current line falls inside any Started..
    # SUCCEEDED window. Logs without these markers at all fall back to the
    # prior behavior (capture RMT blocks anywhere).
    lines = text.splitlines()
    _has_task_markers = bool(RMT_TASK_START_PATTERN.search(text))

    # Precompute, per line index, whether that line falls inside a Rank Margin
    # Tool Started..SUCCEEDED window. This MUST be a separate pass rather than
    # incremental toggling during the main scan below, because the main scan's
    # inner "consume a START_RMT block" loop advances `i` directly and skips
    # lines without re-checking them — a Started marker sitting inside a
    # *different* (e.g. stray/composite-UI) RMT block's consumed range would
    # otherwise silently never be seen.
    in_task_window = [False] * len(lines)
    _in_task = False
    for _idx, _raw in enumerate(lines):
        _l = _raw.strip()
        if RMT_TASK_START_PATTERN.search(_l):
            _in_task = True
        in_task_window[_idx] = _in_task
        if RMT_TASK_END_PATTERN.search(_l):
            _in_task = False

    rows: list[dict[str, Any]] = []
    ctx = ParseContext()
    block_index = 0

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        freq_match = FREQ_PATTERN.search(line)
        if freq_match:
            ctx.frequency = int(freq_match.group(1))

        # Preferred source: "Requested/actual ratio .., Frequency=X, GearMode=Y"
        # Sets both frequency and gear in one shot; works for all known log formats.
        fg_match = FREQ_GEAR_PATTERN.search(line)
        if fg_match:
            ctx.frequency = int(fg_match.group(1))
            gear_mode = int(fg_match.group(2))
            # GearMode=0 → Gear2, GearMode=1 → Gear4, GearMode=2 → Gear8 …
            ctx.gear = (gear_mode + 1) * 2

        gear_match = GEAR_PATTERN.search(line)
        if gear_match:
            ctx.gear = int(gear_match.group(1) or gear_match.group(2))

        boot_match = BOOT_TEMP_PATTERN.search(line)
        if boot_match:
            ctx.boot_temp = float(boot_match.group(1))

        run_match = RUN_TEMP_PATTERN.search(line)
        if run_match:
            ctx.run_temp = float(run_match.group(1))

        if line.upper().startswith("START_RMT"):
            capture_block = (not _has_task_markers) or in_task_window[i]
            block_index += 1
            header_line: str | None = None
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("Mc"):
                if HEADER_PATTERN.match(lines[i].strip()):
                    header_line = lines[i].strip()
                i += 1
            param_order = _header_param_order(header_line)

            while i < len(lines):
                current = lines[i].strip()
                if not current:
                    i += 1
                    continue
                if current.upper().startswith("START_RMT"):
                    i -= 1
                    break
                if current.startswith("Setting ") or current.startswith("$"):
                    break

                row_match = ROW_PATTERN.match(current)
                if not row_match:
                    if current.startswith("Mc"):
                        break
                    i += 1
                    continue

                if not capture_block:
                    i += 1
                    continue

                rank = row_match.group(1)
                values = [float(v) for v in FLOAT_PATTERN.findall(row_match.group(2))]

                record: dict[str, Any] = {
                    "SourceFile": source_name,
                    "Frequency": ctx.frequency,
                    "Gear": ctx.gear,
                    "BlockIndex": block_index,
                    "Params": rank,
                    "BootTemp": ctx.boot_temp,
                    "RunTemp": ctx.run_temp,  # None when no temperature sweep in log
                }

                if not _assign_param_values(record, values, param_order):
                    i += 1
                    continue

                rows.append(record)
                i += 1
            # Continue scanning after block
        i += 1

    # ── Fallback: extract Frequency / Gear from source filename ──────
    # Consolidated / interpreter-snippet logs (e.g. "4800_G2_rtt_wr_120.log")
    # are bare RMT captures without a boot header, so FREQ_PATTERN and
    # GEAR_PATTERN never fire.  When rows exist but frequency or gear is still
    # None, try to recover both values from the filename itself.
    _rows_need_freq = [r for r in rows if r["Frequency"] is None]
    _rows_need_gear = [r for r in rows if r["Gear"] is None]
    if _rows_need_freq or _rows_need_gear:
        _stem = re.sub(r"\.[^.]+$", "", source_name)
        _fb_freq: int | None = None
        _fb_gear: int | None = None
        # Match patterns like "4800_G2", "5600_G4", "6400_G2" anywhere in name
        _fg = re.search(r"(?<!\d)(\d{4,5})[_\-]G(\d)(?:[_\-\s.]|$)",
                        _stem, re.IGNORECASE)
        if _fg:
            _fb_freq = int(_fg.group(1))
            _fb_gear = int(_fg.group(2))
        else:
            # Try separate extraction: freq as a plausible DDR multiple of 200
            _ff = re.search(r"(?<!\d)(\d{4,5})(?!\d)", _stem)
            if _ff:
                _fv = int(_ff.group(1))
                if 1600 <= _fv <= 12800 and _fv % 200 == 0:
                    _fb_freq = _fv
            _fg2 = re.search(r"[_\-\s\.]G([248])(?:[_\-\s.]|$)", _stem, re.IGNORECASE)
            if _fg2:
                _fb_gear = int(_fg2.group(1))
        if _fb_freq is not None:
            for r in _rows_need_freq:
                r["Frequency"] = _fb_freq
            print(f"  [info] Frequency={_fb_freq} inferred from filename: {source_name}")
        if _fb_gear is not None:
            for r in _rows_need_gear:
                r["Gear"] = _fb_gear
            print(f"  [info] Gear={_fb_gear} inferred from filename: {source_name}")

    return rows


def _normalize_dtr_temp(raw: float | None) -> float | None:
    """Snap a raw PHY temperature to the DTR experiment's nominal setpoint.

    DTR runs use two chamber setpoints — cold (0 °C) and hot (90 °C) — but the
    on-die PHY reading drifts a few degrees around them. Per the DTR spec:
    a reading in [-10, +10] represents the cold setpoint (report 0), and a
    reading in [80, 100] represents the hot setpoint (report 90). Anything
    outside those bands is left as-is (rounded) so unexpected values stay
    visible rather than being silently forced to a setpoint.
    """
    if raw is None:
        return None
    if -10 <= raw <= 10:
        return 0.0
    if 80 <= raw <= 100:
        return 90.0
    return round(raw)


def _parse_one_rmt_block(
    lines: list[str], start_idx: int
) -> tuple[list[tuple[str, list[float]]], int, list[str | None]]:
    """Parse a single ``START_RMT`` .. ``STOP_RMT`` block.

    ``start_idx`` must be the index of the ``START_RMT`` line. Returns
    ``(ranks, next_idx, param_order)`` where ``ranks`` is a list of
    ``(rank_label, values)`` tuples, ``next_idx`` is the index just past the
    terminating ``STOP_RMT`` (or where scanning stopped) and ``param_order``
    maps each value pair to its canonical parameter (from the header).
    """
    ranks: list[tuple[str, list[float]]] = []
    header_line: str | None = None
    i = start_idx + 1
    while i < len(lines):
        cur = lines[i].strip()
        if cur.upper().startswith("STOP_RMT"):
            i += 1
            break
        if cur.upper().startswith("START_RMT"):
            break
        if header_line is None and HEADER_PATTERN.match(cur):
            header_line = cur
        m = ROW_PATTERN.match(cur)
        if m:
            values = [float(v) for v in FLOAT_PATTERN.findall(m.group(2))]
            ranks.append((m.group(1), values))
        i += 1
    return ranks, i, _header_param_order(header_line)


def _phy_temp_after(lines: list[str], idx: int, window: int = 30) -> float | None:
    """Return the first ``PHY Temperature1`` value at/after line ``idx``."""
    for j in range(idx, min(idx + window, len(lines))):
        m = PHY_TEMP_PATTERN.search(lines[j])
        if m:
            return float(m.group(1))
    return None


def _dtr_freq_gear_from_name(source_name: str) -> tuple[int | None, int | None]:
    """Infer (frequency, gear) from a DTR filename such as
    ``BCRH_Hynix_2Rx8_3200g2_147.log`` (note: freq and gear may be adjacent
    with no separator, e.g. ``3200g2``)."""
    stem = re.sub(r"\.[^.]+$", "", source_name)
    m = re.search(r"(?<!\d)(\d{4,5})[_\-]?[gG](\d)", stem)
    if m:
        return int(m.group(1)), int(m.group(2))
    freq: int | None = None
    gear: int | None = None
    mf = re.search(r"(?<!\d)(\d{4,5})(?!\d)", stem)
    if mf:
        fv = int(mf.group(1))
        if 1600 <= fv <= 12800 and fv % 200 == 0:
            freq = fv
    mg = re.search(r"[gG](\d)(?:[_\-\s.]|$)", stem)
    if mg:
        gear = int(mg.group(1))
    return freq, gear


def parse_dtr_rmt_from_text(text: str, source_name: str) -> list[dict[str, Any]]:
    """Parse a DTR (Dynamic Thermal Range) log into exactly two RMT data sets.

    A DTR log holds a Boot-temperature RMT (under the "Temp Drift RMT Test"
    task, captured *before* the "Hit Enter" pause) and a Run-temperature RMT
    (under the "Rank Margin Tool" task, captured *after* the pause). Each
    per-rank row gets:
      * ``BootTemp`` = the first PHY Temperature1 of the experiment (the boot
        RMT's reading), normalised to the nominal setpoint, applied to BOTH
        data sets;
      * ``RunTemp``  = that data set's own PHY Temperature1, normalised.

    Falls back to :func:`parse_rmt_from_text` when the DTR task markers are
    absent (so a mislabelled / non-DTR file still yields something sensible).
    """
    lines = text.splitlines()

    # Locate the boot task, the Hit-Enter pause, and the run task in order.
    boot_task_idx = hit_enter_idx = run_task_idx = None
    for i, raw in enumerate(lines):
        l = raw.strip()
        if boot_task_idx is None:
            if DTR_BOOT_TASK_PATTERN.search(l):
                boot_task_idx = i
            continue
        if hit_enter_idx is None:
            if DTR_HIT_ENTER_PATTERN.search(l):
                hit_enter_idx = i
            continue
        if run_task_idx is None and DTR_RUN_TASK_PATTERN.search(l):
            run_task_idx = i
            break

    if boot_task_idx is None or run_task_idx is None:
        print(f"  [info] No DTR task markers in {source_name}; using standard RMT parser.")
        return parse_rmt_from_text(text, source_name)

    # Boot RMT = first START_RMT after the boot task marker but before the
    # Hit-Enter pause. Run RMT = first START_RMT after the run task marker.
    boot_limit = hit_enter_idx if hit_enter_idx is not None else run_task_idx
    boot_start = next(
        (i for i in range(boot_task_idx, boot_limit)
         if lines[i].strip().upper().startswith("START_RMT")),
        None,
    )
    run_start = next(
        (i for i in range(run_task_idx, len(lines))
         if lines[i].strip().upper().startswith("START_RMT")),
        None,
    )

    # Frequency / Gear: prefer the filename (authoritative label for DTR runs;
    # the log body's "Requested/actual ratio" often reports an early boot
    # frequency that differs from the DTR test setpoint), then fall back to
    # the body context.
    freq, gear = _dtr_freq_gear_from_name(source_name)
    if freq is None or gear is None:
        for line in lines:
            fg = FREQ_GEAR_PATTERN.search(line)
            if fg:
                if freq is None:
                    freq = int(fg.group(1))
                if gear is None:
                    gear = (int(fg.group(2)) + 1) * 2
                break

    rows: list[dict[str, Any]] = []
    global_boot_temp: float | None = None
    for block_index, (start, is_boot) in enumerate(
        ((boot_start, True), (run_start, False)), start=1
    ):
        if start is None:
            continue
        ranks, after, param_order = _parse_one_rmt_block(lines, start)
        phy = _phy_temp_after(lines, after)
        if is_boot:
            global_boot_temp = phy
        run_temp = _normalize_dtr_temp(phy)
        for rank, values in ranks:
            record: dict[str, Any] = {
                "SourceFile": source_name,
                "Frequency": freq,
                "Gear": gear,
                "BlockIndex": block_index,
                "Params": rank,
                "BootTemp": None,  # filled below once the global boot temp is known
                "RunTemp": run_temp,
            }
            if _assign_param_values(record, values, param_order):
                rows.append(record)

    norm_boot = _normalize_dtr_temp(global_boot_temp)
    for r in rows:
        r["BootTemp"] = norm_boot

    if not rows:
        print(f"  [warn] DTR markers present but no RMT rows parsed in {source_name}.")
    else:
        print(
            f"  [info] DTR: parsed {len(rows)} rows from {source_name} "
            f"(Freq={freq}, Gear={gear}, BootTemp={norm_boot})"
        )
    return rows



def collect_input_files(inputs: list[str], pattern: str) -> list[Path]:
    # Support multiple glob patterns separated by ';' or ',' (e.g. "*.log;*.txt").
    patterns = [pat.strip() for pat in re.split(r"[;,]", pattern or "") if pat.strip()]
    if not patterns:
        patterns = ["*.txt"]
    files: list[Path] = []
    for item in inputs:
        p = Path(item)
        if p.is_file():
            files.append(p)
        elif p.is_dir():
            for pat in patterns:
                files.extend(sorted(p.rglob(pat)))
    seen: set[Path] = set()
    unique_files: list[Path] = []
    for f in files:
        resolved = f.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        unique_files.append(f)
    return unique_files



def write_csv(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in columns})



def write_excel(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    def add_sheet(sheet_name: str, data_rows: Iterable[dict[str, Any]], columns: list[str]) -> None:
        ws = wb.create_sheet(title=safe_sheet_name(sheet_name))
        ws.append(columns)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for row in data_rows:
            ws.append([row.get(c, "") for c in columns])
        ws.freeze_panes = "A2"

    # Remove default sheet and then create explicit tabs.
    del wb[wb.sheetnames[0]]

    add_sheet("All_RMT", rows, EXTENDED_COLUMNS)

    by_file: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_freq: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for row in rows:
        by_file[str(row.get("SourceFile", "unknown"))].append(row)
        freq = row.get("Frequency")
        freq_key = f"Freq_{freq}" if freq is not None else "Freq_Unknown"
        by_freq[freq_key].append(row)

    for source_name, source_rows in sorted(by_file.items()):
        add_sheet(f"File_{Path(source_name).stem}", source_rows, EXTENDED_COLUMNS)

    for freq_name, freq_rows in sorted(by_freq.items()):
        add_sheet(freq_name, freq_rows, EXTENDED_COLUMNS)

    wb.save(path)



def calc_window_width(row: dict[str, Any], param: str) -> float | None:
    low = row.get(f"{param}-")
    high = row.get(f"{param}+")
    if low is None or high is None:
        return None
    try:
        return float(high) - float(low)
    except (TypeError, ValueError):
        return None


def average_widths(rows: list[dict[str, Any]]) -> dict[str, float]:
    averages: dict[str, float] = {}
    for param in PARAMS:
        widths = [calc_window_width(r, param) for r in rows]
        widths = [w for w in widths if w is not None]
        if widths:
            averages[param] = statistics.mean(widths)
    return averages


def _escape_jsl_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def load_axis_config(path: Path | None) -> dict:
    """Load JMP axis settings JSON.  Returns {} on missing/bad file."""
    import json as _json
    if path is None or not path.exists():
        return {}
    try:
        return _json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"Warning: could not load axis config {path}: {e}")
        return {}


def _fmt_num(v: float) -> str:
    """Render a number for JSL (drop a trailing .0)."""
    f = float(v)
    return str(int(f)) if f.is_integer() else str(f)


def _default_ref(defs: dict | None, is_plus: bool) -> float:
    """Default +Ref / -Ref from the axis JSON ``defaults`` (fallback +/-10)."""
    key = "ref_line_plus" if is_plus else "ref_line_minus"
    val = (defs or {}).get(key)
    try:
        mag = abs(float(val)) if val is not None else 10.0
    except (TypeError, ValueError):
        mag = 10.0
    mag = int(mag) if mag.is_integer() else mag
    return mag if is_plus else -mag


def _param_ref(param: str, axis_config: dict | None, is_plus: bool) -> float:
    """Per-parameter ref line, falling back to the default +Ref / -Ref."""
    cfg = axis_config or {}
    side = "plus" if is_plus else "minus"
    ref = cfg.get("parameters", {}).get(param, {}).get(side, {}).get("ref_line")
    if ref is not None:
        try:
            return float(ref)
        except (TypeError, ValueError):
            pass
    return _default_ref(cfg.get("defaults"), is_plus)


def _scale_box_dispatch(col_name: str, cfg: dict, defs: dict, is_plus: bool) -> list[str]:
    """Return JSL lines for one Dispatch(ScaleBox) block (no trailing comma)."""
    items: list[str] = []
    if "min" in cfg:
        items.append(f"Min( {cfg['min']} )")
    if "max" in cfg:
        items.append(f"Max( {cfg['max']} )")
    items.append(f"Inc( {cfg.get('inc', defs.get('inc', 5))} )")
    items.append(f"Minor Ticks( {cfg.get('minor_ticks', defs.get('minor_ticks', 0))} )")
    ref = cfg.get("ref_line", _default_ref(defs, is_plus))
    clr_key = "ref_line_plus_color" if is_plus else "ref_line_minus_color"
    clr = defs.get(clr_key, "Medium Light Red" if is_plus else "Blue")
    items.append(f'Add Ref Line( {ref}, "Solid", "{clr}", "", 1 )')
    mg  = defs.get("show_major_grid", 1)
    mng = defs.get("show_minor_grid", 1)
    items.append(f"Label Row( {{Show Major Grid( {mg} ), Show Minor Grid( {mng} )}} )")
    sep  = ",\n                    "
    body = sep.join(items)
    return [
        "                Dispatch(",
        "                    {},",
        f'                    "{col_name}",',
        "                    ScaleBox,",
        f"                    {{{body}}}",
        "                )",
    ]


def _has_temp_data(rows: list[dict[str, Any]]) -> bool:
    """Return True if any row has a non-zero, non-None BootTemp or RunTemp value."""
    for r in rows:
        try:
            if (r.get("BootTemp") is not None and float(r["BootTemp"]) != 0):
                return True
            if (r.get("RunTemp") is not None and float(r["RunTemp"]) != 0):
                return True
        except (TypeError, ValueError):
            pass
    return False


def _compute_axis_range(
    rows: list[dict[str, Any]] | None, param: str, ref_mag: float = 10.0,
    ref_mag_lo: float | None = None,
) -> tuple[float, float, float, float, int]:
    """Compute mirrored, shared Y-axis bounds for *param*+ and *param*-.

    Mirrors the reference .jrp layout where the + panel shows a positive
    magnitude band and the - panel shows the same band negated, so both
    panels share an identical magnitude span (e.g. +[5,40] / -[-40,-5]).
    The shared band always includes the reference-line magnitudes
    (*ref_mag* and, for asymmetric +Ref / -Ref, *ref_mag_lo*; default 10)
    so the pass/fail lines stay visible on both panels.

    Returns ``(plus_min, plus_max, minus_min, minus_max, inc)`` where:
      * ``plus_min``/``plus_max``   apply to the *param*+ panel.
      * ``minus_min``/``minus_max`` are the negated mirror for the *param*- panel.
    """
    import math as _m

    ref_mag = abs(float(ref_mag))
    ref_lo = abs(float(ref_mag_lo)) if ref_mag_lo is not None else ref_mag

    plus_abs:  list[float] = []
    minus_abs: list[float] = []
    for r in (rows or []):
        vp = r.get(f"{param}+")
        if vp is not None:
            try:
                plus_abs.append(abs(float(vp)))
            except (ValueError, TypeError):
                pass
        vm = r.get(f"{param}-")
        if vm is not None:
            try:
                minus_abs.append(abs(float(vm)))
            except (ValueError, TypeError):
                pass

    all_abs = plus_abs + minus_abs
    if not all_abs:
        return 5.0, 40.0, -40.0, -5.0, 5

    # Include the reference-line magnitudes so they are always visible on the axis
    abs_min = min(min(all_abs), ref_mag, ref_lo)
    abs_max = max(max(all_abs), ref_mag, ref_lo)
    abs_rng = abs_max - abs_min

    # Choose a sensible Inc based on the magnitude span
    if abs_rng <= 10:
        inc = 1
    elif abs_rng <= 20:
        inc = 2
    elif abs_rng <= 50:
        inc = 5
    elif abs_rng <= 100:
        inc = 10
    else:
        inc = 20

    # Snap magnitude band outward to Inc boundaries (shared by both panels)
    mag_lo = float(_m.floor((abs_min - inc * 0.15) / inc) * inc)
    mag_hi = float(_m.ceil((abs_max + inc * 0.15) / inc) * inc)
    if mag_lo < 0:
        mag_lo = 0.0

    plus_min,  plus_max  = mag_lo, mag_hi
    minus_min, minus_max = -mag_hi, -mag_lo
    return plus_min, plus_max, minus_min, minus_max, inc


def _dataset_overlay_col(rows: list[dict[str, Any]] | None) -> str | None:
    """Return the column to colour-overlay by when several data sets are merged.

    When multiple CSV files are concatenated each row carries a distinct
    ``SourceFile`` value; overlaying by it lets the JMP charts differentiate
    the merged data sets (e.g. gear-2 vs gear-4 captures).  Returns ``None``
    when the data comes from a single source.
    """
    if not rows:
        return None
    srcs = {
        str(r.get("SourceFile")).strip()
        for r in rows
        if r.get("SourceFile") not in (None, "")
    }
    return "SourceFile" if len(srcs) > 1 else None


def generate_jmp_jsl(
    csv_path: Path,
    output_dir: Path,
    chart_params: list[str],
    axis_config: dict | None = None,
    rows: list[dict[str, Any]] | None = None,
    interactive: bool = False,
) -> Path:
    """Generate JMP JSL for the given CSV.

    Temperature mode (when *rows* contains non-zero BootTemp / RunTemp):
      Group X  : BootTemp, RunTemp
      Y-axis   : dynamic Min/Max computed from actual data; common bounds for
                 param+ and param- so both panels share the same scale.
      Ref lines: per-parameter +Ref / -Ref from the axis config (falling back
                 to the default +Ref / -Ref, then +/-10).

    Standard mode (no temperature sweep):
      Group X  : Frequency (side-by-side facets) with Gear colour overlay.
      Y-axis   : driven by jmp_axis_settings.json when supplied.

    Multi data-set mode (concatenated CSVs with several ``SourceFile`` values):
      An ``Overlay( :SourceFile )`` is added so each merged source is drawn in
      a distinct colour, allowing the captures to be told apart.

    Interactive mode (*interactive* = True) writes ``rmt_graph_builder.jsl``
    which leaves every Graph Builder window open (no PNG export, no Quit) so
    the engineer can fine-tune the charts in JMP and save a .jrp/.jsl.
    """
    import math as _m

    output_dir.mkdir(parents=True, exist_ok=True)
    jsl_path  = output_dir / ("rmt_graph_builder.jsl" if interactive else "rmt_jmp_charts.jsl")
    image_dir = output_dir / "jmp_charts"
    if not interactive:
        image_dir.mkdir(parents=True, exist_ok=True)

    csv_jsl     = _escape_jsl_string(str(csv_path.resolve()))
    img_dir_jsl = _escape_jsl_string(str(image_dir.resolve()))

    cfg_params = (axis_config or {}).get("parameters", {})
    cfg_defs   = (axis_config or {}).get("defaults", {})

    temp_mode   = bool(rows) and _has_temp_data(rows)
    overlay_col = _dataset_overlay_col(rows)

    # ── Preamble ─────────────────────────────────────────────────────────
    lines: list[str] = [
        "Names Default To Here(1);",
        f'dt = Open("{csv_jsl}");',
        'If( Is Empty( dt ), Throw("Failed to open CSV") );',
        "",
        'col_names = dt << Get Column Names( String );',
        'Write("Columns: " || Concat Items(col_names, ", ") || "\\!N");',
        "",
        'If( Contains(col_names, "Params"),',
        '    dt << Get Column("Params") << Set Modeling Type("Nominal");',
        ');',
        "",
    ]
    if temp_mode:
        lines += [
            'If( Contains(col_names, "BootTemp"),',
            '    dt << Get Column("BootTemp") << Set Modeling Type("Nominal");',
            ');',
            'If( Contains(col_names, "RunTemp"),',
            '    dt << Get Column("RunTemp") << Set Modeling Type("Nominal");',
            ');',
            "",
        ]
    else:
        lines += [
            'If( Contains(col_names, "Frequency"),',
            '    dt << Get Column("Frequency") << Set Modeling Type("Nominal");',
            ');',
            'If( Contains(col_names, "Gear"),',
            '    dt << Get Column("Gear") << Set Modeling Type("Nominal");',
            ');',
            "",
        ]
    if overlay_col:
        lines += [
            f'If( Contains(col_names, "{overlay_col}"),',
            f'    dt << Get Column("{overlay_col}") << Set Modeling Type("Nominal");',
            ');',
            "",
        ]

    # ── Per-parameter charts ──────────────────────────────────────────────
    for legend_idx, param in enumerate(chart_params, start=1):
        param_plus  = f"{param}+"
        param_minus = f"{param}-"
        out_png     = _escape_jsl_string(str((image_dir / f"{param}.png").resolve()))
        leg_a = legend_idx * 2 - 1
        leg_b = legend_idx * 2

        if temp_mode:
            # Mirrored, shared Y-axis bounds for + and - panels
            ref_p = _param_ref(param, axis_config, is_plus=True)
            ref_m = _param_ref(param, axis_config, is_plus=False)
            p_min, p_max, m_min, m_max, inc = _compute_axis_range(
                rows, param, ref_mag=max(abs(ref_p), abs(ref_m)),
                ref_mag_lo=min(abs(ref_p), abs(ref_m)))
            clr_p = cfg_defs.get("ref_line_plus_color", "Orange")
            clr_m = cfg_defs.get("ref_line_minus_color", "Medium Dark Blue")
            _ovl = [f"            Overlay( :{overlay_col} )"] if overlay_col else []
            _grp_sep = "," if overlay_col else ""
            lines += [
                f"// ---- Chart: {param} [temperature mode] ----",
                f'If( Contains(col_names, "{param_plus}") & Contains(col_names, "{param_minus}"),',
                f"    (gb = dt << Graph Builder(",
                f"        Size( 842, 756 ),",
                f"        Variables(",
                f"            X( :Params ),",
                f'            Y( :"{param_plus}"n ),',
                f'            Y( :"{param_minus}"n ),',
                f"            Group X( :BootTemp ),",
                f"            Group X( :RunTemp ){_grp_sep}",
                *_ovl,
                f"        ),",
                f"        Elements( Position( 1, 1 ), Points( X, Y, Legend( {leg_a} ) ) ),",
                f"        Elements( Position( 1, 2 ), Points( X, Y, Legend( {leg_b} ) ) ),",
                f"        SendToReport(",
                f"            Dispatch(",
                f"                {{}},",
                f'                "{param_plus}",',
                f"                ScaleBox,",
                f"                {{Min( {p_min} ), Max( {p_max} ), Inc( {inc} ),",
                f"                Minor Ticks( 0 ),",
                f'                Add Ref Line( {_fmt_num(ref_p)}, "Solid", "{clr_p}", "", 1 ),',
                f"                Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}",
                f"            ),",
                f"            Dispatch(",
                f"                {{}},",
                f'                "{param_minus}",',
                f"                ScaleBox,",
                f"                {{Min( {m_min} ), Max( {m_max} ), Inc( {inc} ),",
                f"                Minor Ticks( 0 ),",
                f'                Add Ref Line( {_fmt_num(ref_m)}, "Solid", "{clr_m}", "", 1 ),',
                f"                Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}",
                f"            )",
                f"        )",
                f"    );",
                f'    gb << Save Picture("{out_png}", "png");',
                f"    gb << Close Window;)",
                f");",
                "",
            ]
        else:
            pcfg      = cfg_params.get(param, {})
            plus_cfg  = pcfg.get("plus",  {})
            minus_cfg = pcfg.get("minus", {})
            disp_p    = _scale_box_dispatch(param_plus,  plus_cfg,  cfg_defs, is_plus=True)
            disp_m    = _scale_box_dispatch(param_minus, minus_cfg, cfg_defs, is_plus=False)
            disp_p_c  = disp_p[:-1] + [disp_p[-1] + ","]
            send_to   = (
                ["            SendToReport("]
                + disp_p_c
                + disp_m
                + ["            )"]
            )

            def _gb(with_freq: bool) -> list[str]:
                _ov = overlay_col if overlay_col else "Gear"
                grp   = ["                Group X( :Frequency ),",
                         f"                Overlay( :{_ov} )"] if with_freq else (
                         [f"                Overlay( :{overlay_col} )"] if overlay_col else [])
                sep   = "," if (with_freq or overlay_col) else ""
                return [
                    f"        gb = dt << Graph Builder(",
                    f"            Size( {'700' if with_freq else '600'}, 820 ),",
                    f"            Variables(",
                    f"                X( :Params ),",
                    f'                Y( :"{param_plus}"n ),',
                    f'                Y( :"{param_minus}"n ){sep}',
                    *grp,
                    f"            ),",
                    f"            Elements( Position( 1, 1 ), Points( X, Y, Legend( {leg_a} ) ) ),",
                    f"            Elements( Position( 1, 2 ), Points( X, Y, Legend( {leg_b} ) ) ),",
                    *send_to,
                    f"        );",
                ]

            lines += [
                f"// ---- Chart: {param} ----",
                f'If( Contains(col_names, "{param_plus}") & Contains(col_names, "{param_minus}"),',
                f'    (If( Contains(col_names, "Frequency"),',
                *["    " + l for l in _gb(with_freq=True)],
                f"    ,",
                *["    " + l for l in _gb(with_freq=False)],
                f"    );",
                f'    gb << Save Picture("{out_png}", "png");',
                f"    gb << Close Window;)",
                f");",
                "",
            ]

    # ── Cleanup ──────────────────────────────────────────────────────────
    # NOTE: The aggregated "RMT Dashboard" PNG is intentionally NOT generated
    # here. Dashboard-style multi-parameter / multi-source analysis now lives
    # in the interactive HTML report (RMT_Report.html), which renders the same
    # information with selectable sources and parameters.
    if interactive:
        lines = [
            ln for ln in lines
            if "Save Picture(" not in ln
        ]
        lines = [ln.replace("gb << Close Window;)", ")") for ln in lines]
        lines += [
            "",
            'Write("Graph Builder windows are open for interactive tuning. '
            'Use File > Save Script / Save As .jrp, then Load .jrp in MarginIQ.\\!N");',
        ]
    else:
        lines += [
            "Close( dt, No Save );",
            "",
            f'Write("JMP chart export completed.\\!N");',
            f'Write("Output folder: {img_dir_jsl}\\!N");',
            "",
            "// Auto-close JMP after all charts are saved",
            "Quit();",
        ]

    jsl_path.write_text("\n".join(lines), encoding="utf-8")
    return jsl_path


def _terminate_running_jmp(jmp_exe: Path) -> None:
    """Best-effort: force-close any already-running JMP process.

    JMP is a single-instance Windows application: if a PRIOR invocation (an
    earlier file in a "process each file separately" batch, or a JMP window
    left open from a previous session) is still running when the next
    ``jmp_exe jsl_script`` launch happens, Windows silently FORWARDS the new
    invocation's command line to that already-open instance instead of
    starting a fresh process. The forwarded launch's ``subprocess.run(...)``
    call then returns almost immediately (the thin launcher process exits
    right after forwarding) — long before the actual chart generation for
    THIS file's script has happened, or it silently races with whatever the
    existing instance was still doing. This is exactly why JMP charts fail
    to appear for the 2nd, 3rd, ... file when running "Process each file
    SEPARATELY" over multiple files back-to-back. Killing any stray
    instance first guarantees every file gets its own clean, standalone
    JMP run instead.
    """
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", jmp_exe.name],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=10, check=False,
        )
        # Give Windows a moment to fully release the process / DDE server
        # registration before the next instance is launched.
        import time as _time
        _time.sleep(1)
    except Exception:
        pass


def run_jmp_script(jmp_exe: Path, jsl_script: Path, extra_wait_s: int = 30) -> None:
    if not jmp_exe.exists():
        raise FileNotFoundError(f"JMP executable not found: {jmp_exe}")
    if not jsl_script.exists():
        raise FileNotFoundError(f"JSL script not found: {jsl_script}")

    # Force-close any stray/leftover JMP instance so this launch always
    # starts a clean, standalone process instead of being silently forwarded
    # to an already-open one — see _terminate_running_jmp() docstring.
    _terminate_running_jmp(jmp_exe)

    print(f"\n📊 Running JMP charts generation...")
    print(f"   JSL Script: {jsl_script}")
    print(f"   JMP Exe: {jmp_exe}")

    jmp_charts_dir = jsl_script.parent / "jmp_charts"

    # JMP runs JSL when script path is passed as an argument
    # Note: JMP may display a GUI window, which is normal behavior
    try:
        result = subprocess.run([str(jmp_exe), str(jsl_script)], check=False, timeout=120)
        if result.returncode == 0:
            print("✅ JMP chart generation completed successfully")
        else:
            print(f"⚠️  JMP exited with code {result.returncode}")
    except subprocess.TimeoutExpired:
        print("⚠️  JMP timeout (exceeded 2 minutes) - charts may still be generating")
    except Exception as e:
        print(f"❌ JMP execution failed: {e}")
        raise

    # ── Verify PNGs actually landed in THIS run's output folder ──────────
    # A clean standalone instance was just launched (any stray JMP process
    # was force-closed above), so this is now a normal "give the app a
    # moment to finish writing" wait, not a workaround for cross-instance
    # forwarding.
    import time as _time
    waited = 0
    while waited < extra_wait_s and not (
        jmp_charts_dir.exists() and any(jmp_charts_dir.glob("*.png"))
    ):
        _time.sleep(1)
        waited += 1
    if jmp_charts_dir.exists() and any(jmp_charts_dir.glob("*.png")):
        n = len(list(jmp_charts_dir.glob("*.png")))
        print(f"✓ Verified {n} PNG chart(s) in {jmp_charts_dir}"
              + (f" (waited {waited}s)" if waited else ""))
    else:
        print(
            f"⚠️  No PNG charts found in {jmp_charts_dir} after waiting {waited}s.\n"
            "   Check the JSL script and JMP for errors (a dialog box may be "
            "waiting for input, blocking the script from completing)."
        )


def parse_chart_fields(raw_value: str | None) -> list[str]:
    if not raw_value:
        return PARAMS.copy()
    requested = [item.strip() for item in raw_value.split(",") if item.strip()]
    valid = [item for item in requested if item in PARAMS]
    # Keep user-provided order and deduplicate.
    deduped: list[str] = []
    seen: set[str] = set()
    for item in valid:
        if item not in seen:
            seen.add(item)
            deduped.append(item)
    return deduped


def ask_user_chart_approval(default_fields: list[str]) -> tuple[bool, list[str]]:
    print("\nPPT chart field approval required.")
    print("Available fields:")
    for idx, param in enumerate(PARAMS, start=1):
        print(f"  {idx}. {param}")
    print(
        "Select fields by number or name (comma-separated), type 'all' for all fields, or 'none' to skip PPT charts."
    )
    raw = input(
        f"Chart fields [{', '.join(default_fields)}]: "
    ).strip()

    if not raw:
        selected = default_fields
    elif raw.lower() == "all":
        selected = PARAMS.copy()
    elif raw.lower() == "none":
        selected = []
    else:
        selected = []
        tokens = [token.strip() for token in raw.split(",") if token.strip()]
        for token in tokens:
            if token.isdigit():
                index = int(token)
                if 1 <= index <= len(PARAMS):
                    selected.append(PARAMS[index - 1])
                continue
            if token in PARAMS:
                selected.append(token)

        # De-duplicate while preserving order.
        deduped: list[str] = []
        seen: set[str] = set()
        for item in selected:
            if item not in seen:
                seen.add(item)
                deduped.append(item)
        selected = deduped

    if selected:
        print(f"Proposed chart fields: {', '.join(selected)}")
    else:
        print("Proposed chart fields: none (no chart slides)")

    approval = input("Approve these chart fields? (y/n): ").strip().lower()
    if approval != "y":
        print("Chart generation not approved. PPT chart slides will be skipped.")
        return False, []
    return True, selected


def _delete_all_slides(prs: "Presentation") -> None:  # type: ignore[name-defined]
    """Remove all existing slides from a presentation (preserves master/layouts)."""
    from pptx.oxml.ns import qn as _qn
    prs_part   = prs.part
    sldIdLst   = prs_part._element.find(_qn("p:sldIdLst"))
    if sldIdLst is None:
        return
    rid_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    # Collect rIds before modifying the list
    rids = [sldId.get(rid_attr) for sldId in list(sldIdLst)]
    # Remove all sldId elements from the XML list
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)
    # Drop each relationship so the slide XML files are removed from the package
    for rid in rids:
        if rid:
            try:
                prs_part.drop_rel(rid)
            except Exception:
                pass


def build_ppt_from_jmp_charts(
    path: Path,
    jmp_charts_dir: Path,
    template_path: Path | None = None,
) -> int:
    """Create a PowerPoint from the JMP-generated PNG charts.

    Each PNG gets its own slide with the image filling the slide (with margins).
    If *template_path* is given the presentation theme/master is copied from it;
    otherwise a blank widescreen presentation is used.

    Returns the number of slides added.
    """
    if Presentation is None or Inches is None:
        raise RuntimeError("python-pptx is not installed. Run: pip install python-pptx")

    png_files = sorted(jmp_charts_dir.glob("*.png"))
    if not png_files:
        print(f"No PNG files found in {jmp_charts_dir}")
        return 0

    if template_path and template_path.exists():
        prs = Presentation(str(template_path))
        _delete_all_slides(prs)
        print(f"Using template: {template_path}")
    else:
        prs = Presentation()
        prs.slide_width  = int(13.33 * 914400)   # 13.33 inches widescreen
        prs.slide_height = int(7.5  * 914400)

    # Use the Blank layout (index 6 in standard themes, fall back to 0)
    blank_layout = None
    for layout in prs.slide_layouts:
        if "blank" in layout.name.lower():
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]

    # Use the exact image position from the reference template
    # (x=4.406", y=0.687", w=4.522", h=6.126" in EMU)
    # This places the chart in the content area next to the slide-master branding.
    IMG_LEFT   = 4028969
    IMG_TOP    = 628506
    IMG_WIDTH  = 4134062
    IMG_HEIGHT = 5600988

    for png in png_files:
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            str(png),
            IMG_LEFT, IMG_TOP,
            IMG_WIDTH, IMG_HEIGHT,
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    print(f"Saved JMP chart PPT ({len(png_files)} slides): {path}")
    return len(png_files)


def build_ppt(path: Path, rows: list[dict[str, Any]], chart_params: list[str]) -> None:
    if Presentation is None or CategoryChartData is None or XL_CHART_TYPE is None or Inches is None:
        raise RuntimeError(
            "python-pptx is not installed. Install it with: pip install python-pptx"
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()

    title_slide_layout = prs.slide_layouts[0]
    content_layout = prs.slide_layouts[1]

    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = "RMT Extraction Summary"
    subtitle = slide.placeholders[1]

    frequencies = sorted({r.get("Frequency") for r in rows if r.get("Frequency") is not None})
    files = sorted({r.get("SourceFile") for r in rows})
    subtitle.text = (
        f"Files: {len(files)} | RMT rows: {len(rows)} | "
        f"Frequencies: {', '.join(str(f) for f in frequencies) if frequencies else 'N/A'}"
    )

    # Overall metric slide
    metric_slide = prs.slides.add_slide(content_layout)
    metric_slide.shapes.title.text = "Average Window Width By Parameter"
    tf = metric_slide.placeholders[1].text_frame
    tf.clear()

    overall_avg = average_widths(rows)
    for param in PARAMS:
        avg_width = overall_avg.get(param)
        if avg_width is None:
            continue
        p = tf.add_paragraph()
        p.text = f"{param}: {avg_width:.2f}"

    if chart_params:
        # Overall chart slide.
        overall_chart_slide = prs.slides.add_slide(content_layout)
        overall_chart_slide.shapes.title.text = "Overall Avg Window Width (Chart)"
        chart_data = CategoryChartData()
        chart_data.categories = chart_params
        chart_data.add_series(
            "All Rows",
            [overall_avg.get(param, 0.0) for param in chart_params],
        )
        overall_chart = overall_chart_slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.7),
            Inches(1.6),
            Inches(12.0),
            Inches(5.0),
            chart_data,
        ).chart
        overall_chart.has_legend = False
        overall_chart.value_axis.has_major_gridlines = True

    # Frequency breakdown slides
    by_freq: dict[Any, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        by_freq[row.get("Frequency")].append(row)

    if chart_params:
        # Cross-frequency comparison chart with one series per frequency.
        compare_slide = prs.slides.add_slide(content_layout)
        compare_slide.shapes.title.text = "Avg Width By Frequency (Chart)"
        compare_data = CategoryChartData()
        compare_data.categories = chart_params
        for freq, freq_rows in sorted(by_freq.items(), key=lambda item: (item[0] is None, item[0])):
            freq_avg = average_widths(freq_rows)
            compare_data.add_series(
                f"Freq {freq if freq is not None else 'Unknown'}",
                [freq_avg.get(param, 0.0) for param in chart_params],
            )
        compare_chart = compare_slide.shapes.add_chart(
            XL_CHART_TYPE.LINE_MARKERS,
            Inches(0.7),
            Inches(1.6),
            Inches(12.0),
            Inches(5.0),
            compare_data,
        ).chart
        compare_chart.has_legend = True
        compare_chart.value_axis.has_major_gridlines = True

        # RunTemp trend charts by rank for approved fields.
        for param in chart_params:
            runtemp_points: dict[str, dict[float, list[float]]] = defaultdict(lambda: defaultdict(list))
            for row in rows:
                run_temp = row.get("RunTemp")
                rank = str(row.get("Params", "Unknown"))
                width = calc_window_width(row, param)
                if run_temp is None or width is None:
                    continue
                try:
                    run_temp_value = float(run_temp)
                except (ValueError, TypeError):
                    continue
                runtemp_points[rank][run_temp_value].append(width)

            if not runtemp_points:
                continue

            all_temps = sorted({temp for by_temp in runtemp_points.values() for temp in by_temp.keys()})
            if not all_temps:
                continue

            trend_slide = prs.slides.add_slide(content_layout)
            trend_slide.shapes.title.text = f"{param} Width vs RunTemp (By Rank)"

            trend_data = CategoryChartData()
            trend_data.categories = [str(temp) for temp in all_temps]
            for rank, by_temp in sorted(runtemp_points.items()):
                series_values = []
                for temp in all_temps:
                    values = by_temp.get(temp)
                    series_values.append(statistics.mean(values) if values else None)
                trend_data.add_series(rank, series_values)

            trend_chart = trend_slide.shapes.add_chart(
                XL_CHART_TYPE.LINE_MARKERS,
                Inches(0.7),
                Inches(1.6),
                Inches(12.0),
                Inches(5.0),
                trend_data,
            ).chart
            trend_chart.has_legend = True
            trend_chart.value_axis.has_major_gridlines = True

    for freq, freq_rows in sorted(by_freq.items(), key=lambda item: (item[0] is None, item[0])):
        s = prs.slides.add_slide(content_layout)
        s.shapes.title.text = f"Frequency {freq if freq is not None else 'Unknown'}"
        body = s.placeholders[1].text_frame
        body.clear()

        p = body.add_paragraph()
        p.text = f"Rows: {len(freq_rows)}"

        for param in PARAMS:
            avg_width = average_widths(freq_rows).get(param)
            if avg_width is None:
                continue
            line = body.add_paragraph()
            line.text = f"{param} avg width: {avg_width:.2f}"

    prs.save(path)


# ---------------------------------------------------------------------------
# Platform / DIMM info extraction
# ---------------------------------------------------------------------------

#: JEDEC JEP106 manufacturer lookup: (bank_byte, mfr_byte) -> name
_JEDEC_MFR: dict[tuple, str] = {
    (0x80, 0xCE): "Samsung",
    (0x80, 0x2C): "Micron",
    (0x80, 0xAD): "SK Hynix",
    (0x80, 0x9E): "Kingston",
    (0x80, 0x43): "Ramaxel",
    (0x80, 0x04): "Nanya",
    (0x80, 0xF1): "G.Skill",
    (0x80, 0x5E): "TeamGroup",
    (0x80, 0x7F): "Corsair",
    (0x80, 0x01): "AMD",
}


def _mfr_from_jedec_code(code_str: str) -> str:
    """Decode a JEDEC ID hex string (e.g. '0xCE80') to manufacturer name."""
    try:
        c = int(code_str, 16)
        hi, lo = (c >> 8) & 0xFF, c & 0xFF
        return (_JEDEC_MFR.get((lo, hi)) or _JEDEC_MFR.get((hi, lo)) or code_str)
    except (ValueError, TypeError):
        return code_str


def _date_from_code(code_str: str) -> str:
    """Decode SPD date code '0xYYWW' (BCD) to 'Wk{ww}/{yyyy}'."""
    try:
        c = int(code_str, 16)
        yy, ww = (c >> 8) & 0xFF, c & 0xFF
        year = (yy >> 4) * 10 + (yy & 0xF) + 2000
        week = (ww >> 4) * 10 + (ww & 0xF)
        return f"Wk{week:02d}/{year}"
    except (ValueError, TypeError):
        return code_str


def parse_platform_info(text: str, source_name: str) -> dict[str, Any]:
    """Extract board/RVP, CPU, and per-DIMM details from a MRC log file.

    Returns a dict with keys:
        source_file, board, board_id, fab_id, cpu, phy_version, dimms (list)

    Each dimm dict has:
        slot, populated, module_type, spd_rev, manufacturer, part_number,
        size_mb, ranks, sdram_width, bus_width, sdram_capacity_mb,
        banks, bank_groups, ecc, pmic_type,
        dram_id_code, date_code, date_display, density
    """
    lines = text.splitlines()

    info: dict[str, Any] = {
        "source_file":  source_name,
        "board":        None,
        "board_id":     None,
        "fab_id":       None,
        "cpu":          None,
        "phy_version":  None,
        "dimms":        [],
    }

    # ── Compiled patterns ─────────────────────────────────────────
    P_BOARD  = re.compile(r"Detected\s+board[:\s]+(.+)", re.I)
    P_BRDID  = re.compile(
        r"BoardID from EC:\s*(0x[0-9A-Fa-f]+).*?FAB ID:\s*(\d+).*?BOM ID:\s*(\d+)", re.I
    )
    P_CPU    = re.compile(r"^CPU:\s+(.+)")
    P_PHY    = re.compile(r"PHY IP Version:\s+(.+)")
    P_SPDAD  = re.compile(
        r"Controller,\s*Channel,\s*Dimm\s+(\d+/\d+/\d+)\s+SpdAddress:\s*([0-9A-Fa-f]+)h"
    )
    P_SPDLN  = re.compile(r"\s+([0-9A-Fa-f]+)h\(\s*\d+\)\s*:\s*((?:[0-9A-Fa-f]{2}\s*)+)")
    P_SPDHD  = re.compile(r"^SPD:\s+00\s+01")
    P_SPDREC = re.compile(r"SPD\s+Dimm\s+recognition.*?Dimm\s+(\d+/\d+/\d+)", re.I)
    # No re.I flag intentionally: DIMM type keywords (CSODIMM, SODIMM, RDIMM…)
    # are ALL-CAPS while "detected" is lowercase — this separates them cleanly.
    P_DDRTY  = re.compile(r"(DDR\d(?:\s+[A-Z][A-Z0-9]+)?)\s+detected.*?Rev:\s*([\d.]+)")
    P_NODIM  = re.compile(r"No\s+DIMM\s+detected", re.I)
    # Per-DIMM field patterns are anchored to the start of the line: the SPD
    # block has no end marker, so unanchored patterns would also match later
    # lines such as "Outputs->MaxRanks: 2" and corrupt the last slot.
    P_RANKS  = re.compile(r"^\s*Ranks:\s*(\d+)")
    P_SCAP   = re.compile(r"^\s*SDRAM\s+Capacity:\s*(\d+)\s*Mb")
    P_DSIZ   = re.compile(r"^\s*DIMM\s+size:\s*(\d+)\s*MByte")
    P_SDWID  = re.compile(r"^\s*SDRAM\s+device\s+width:\s*(\d+)")
    P_BUSWID = re.compile(r"^\s*Primary\s+bus\s+width:\s*(\d+)")
    P_BANKS  = re.compile(r"^\s*(\d+)\s+Banks\s+in\s+(\d+)\s+groups")
    P_ECC    = re.compile(r"^\s*ECC\s+is\s+(not\s+)?supported")
    P_PMIC   = re.compile(r"^\s*PMIC\s+type:\s*(\d+)")
    P_CTCD   = re.compile(r"Controller/Channel/Dimm:\s*(\d+/\d+/\d+)")
    P_DRID   = re.compile(r"DramIdCode:\s*(0x[0-9A-Fa-f]+)")
    P_DTCD   = re.compile(r"DateCode:\s*(0x[0-9A-Fa-f]+)")
    P_DENS   = re.compile(r"^\s+Density:\s*(\d+)")

    # ── State ─────────────────────────────────────────────────────
    spd_raw: dict[str, dict[int, int]] = {}     # slot -> {offset: byte}
    cur_spd: str | None = None
    in_spd:  bool = False

    dimm_hr: dict[str, dict] = {}               # slot -> human-readable block
    cur_rec: str | None = None

    dimm_ex: dict[str, dict] = {}               # slot -> {dram_id, date_code, density}
    cur_ex:  str | None = None

    for line in lines:

        # ── Global system fields ───────────────────────────────────
        if info["board"] is None:
            m = P_BOARD.search(line)
            if m:
                info["board"] = m.group(1).strip()

        if info["board_id"] is None:
            m = P_BRDID.search(line)
            if m:
                info["board_id"] = m.group(1)
                info["fab_id"]   = m.group(2)

        if info["cpu"] is None:
            m = P_CPU.match(line)
            if m:
                info["cpu"] = m.group(1).strip()

        if info["phy_version"] is None:
            m = P_PHY.search(line)
            if m:
                info["phy_version"] = m.group(1).strip()

        # ── Raw SPD hex dump ───────────────────────────────────────
        m = P_SPDAD.search(line)
        if m:
            slot, addr = m.group(1), m.group(2)
            if addr != "0":
                cur_spd = slot
                spd_raw[slot] = {}
            else:
                cur_spd = None
            in_spd = False
            continue

        if cur_spd and P_SPDHD.match(line):
            in_spd = True
            continue

        if in_spd and cur_spd:
            m = P_SPDLN.match(line)
            if m:
                base = int(m.group(1), 16)
                for j, b in enumerate(int(x, 16) for x in m.group(2).split() if x):
                    spd_raw[cur_spd][base + j] = b
            else:
                in_spd = False
                cur_spd = None

        # ── SPD Dimm recognition (human-readable details) ──────────
        m = P_SPDREC.search(line)
        if m:
            cur_rec = m.group(1)
            dimm_hr[cur_rec] = {
                "slot": cur_rec, "populated": False,
                "module_type": None, "spd_rev": None,
                "size_mb": None, "ranks": None,
                "sdram_width": None, "bus_width": None,
                "sdram_capacity_mb": None,
                "banks": None, "bank_groups": None,
                "ecc": None, "pmic_type": None,
            }
            continue

        if cur_rec and cur_rec in dimm_hr:
            d = dimm_hr[cur_rec]
            if P_NODIM.search(line):
                cur_rec = None
                continue
            m = P_DDRTY.search(line)
            if m:
                d["populated"] = True
                d["module_type"] = m.group(1).strip()
                d["spd_rev"] = m.group(2)
                continue
            # First value wins: each block is re-initialised on its header, so
            # any later match can only come from unrelated log lines.
            for pat, key, cvt in [
                (P_RANKS,  "ranks",            int),
                (P_SCAP,   "sdram_capacity_mb", int),
                (P_DSIZ,   "size_mb",           int),
                (P_SDWID,  "sdram_width",       int),
                (P_BUSWID, "bus_width",         int),
                (P_PMIC,   "pmic_type",         str),
            ]:
                if d[key] is None:
                    mm = pat.search(line)
                    if mm:
                        d[key] = cvt(mm.group(1))
            mm = P_BANKS.search(line) if d["banks"] is None else None
            if mm:
                d["banks"] = int(mm.group(1))
                d["bank_groups"] = int(mm.group(2))
            mm = P_ECC.search(line) if d["ecc"] is None else None
            if mm:
                d["ecc"] = mm.group(1) is None   # None = "ECC is supported"

        # ── Controller/Channel/Dimm extra block ────────────────────
        m = P_CTCD.search(line)
        if m:
            cur_ex = m.group(1)
            dimm_ex.setdefault(cur_ex, {})
            continue

        if cur_ex:
            m = P_DRID.search(line)
            if m:
                dimm_ex[cur_ex]["dram_id"] = m.group(1)
            m = P_DTCD.search(line)
            if m:
                dimm_ex[cur_ex]["date_code"] = m.group(1)
            m = P_DENS.match(line)
            if m:
                dimm_ex[cur_ex]["density"] = int(m.group(1))

    # ── Assemble final dimms list ──────────────────────────────────
    def _sort_key(s: str) -> list:
        try:
            return [int(x) for x in s.split("/")]
        except ValueError:
            return [0, 0, 0]

    all_slots = sorted(
        set(list(dimm_hr.keys()) + list(spd_raw.keys())),
        key=_sort_key,
    )

    for slot in all_slots:
        d: dict[str, Any] = dict(dimm_hr.get(slot, {"slot": slot, "populated": False}))

        # Extract part number + manufacturer from raw SPD bytes
        spd = spd_raw.get(slot, {})
        if spd:
            d["populated"] = True
            b0 = spd.get(0x200, 0)
            b1 = spd.get(0x201, 0)
            d.setdefault("manufacturer",
                         _JEDEC_MFR.get((b0, b1), f"ID:0x{b0:02X}{b1:02X}"))
            try:
                raw_pn = bytes([spd.get(k, 0x20) for k in range(0x209, 0x226)])
                pn = "".join(
                    c for c in raw_pn.decode("ascii", "replace")
                    if c.isprintable()
                ).strip()
                d.setdefault("part_number", pn or None)
            except Exception:
                d.setdefault("part_number", None)
        else:
            d.setdefault("manufacturer", None)
            d.setdefault("part_number", None)

        # Merge extra info (DramIdCode, DateCode, Density)
        ex = dimm_ex.get(slot, {})
        dram_id  = ex.get("dram_id")
        date_cd  = ex.get("date_code")
        density  = ex.get("density")
        d["dram_id_code"] = dram_id
        d["date_code"]    = date_cd
        d["date_display"] = _date_from_code(date_cd) if date_cd else None
        d["density"]      = density

        # Fallback manufacturer from DramIdCode if SPD bytes not available
        if not d.get("manufacturer") and dram_id:
            d["manufacturer"] = _mfr_from_jedec_code(dram_id)

        info["dimms"].append(d)

    return info


# ---------------------------------------------------------------------------
# Training step extraction
# ---------------------------------------------------------------------------

def parse_training_steps(text: str, source_name: str) -> list[dict[str, Any]]:
    """Parse MRC task start/end events from a MRC boot log.

    Returns a list of dicts:
        task_type   : "MRC" or "GREEN MRC"
        name        : task name string
        status      : "SUCCEEDED" | "FAILED" | "RUNNING"
        status_code : e.g. "0h", "23h" (None if still running)
        start_line  : 1-based line number of the Start event
        end_line    : 1-based line number of the end event (None if open)
        iteration   : 0-based iteration counter (increments at Pre-Training)
        source      : source_name
    """
    import re as _re

    START_RE = _re.compile(
        r"^(?:(GREEN) )?MRC task -- (.+?) -- Started\.?$", _re.IGNORECASE
    )
    END_RE = _re.compile(
        r"^(?:(GREEN) )?MRC task (.+?) -- (SUCCEEDED|FAILED), Status = (\w+)",
        _re.IGNORECASE,
    )

    tasks: list[dict[str, Any]] = []
    # open_tasks maps name -> task dict so we can back-fill end info
    open_tasks: dict[str, dict[str, Any]] = {}
    iteration = 0
    seen_pre_training = False

    for lineno, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()

        m = START_RE.match(line)
        if m:
            is_green = bool(m.group(1))
            name = m.group(2).strip()
            # Each "Pre-Training" starts a new sweep/iteration (skip the very first)
            if name == "Pre-Training":
                if seen_pre_training:
                    iteration += 1
                seen_pre_training = True
            task: dict[str, Any] = {
                "task_type":   "GREEN MRC" if is_green else "MRC",
                "name":        name,
                "status":      "RUNNING",
                "status_code": None,
                "start_line":  lineno,
                "end_line":    None,
                "iteration":   iteration,
                "source":      source_name,
            }
            open_tasks[name] = task
            tasks.append(task)
            continue

        m = END_RE.match(line)
        if m:
            name = m.group(2).strip()
            status = m.group(3).upper()
            code = m.group(4)
            if name in open_tasks:
                t = open_tasks.pop(name)
                t["status"] = status
                t["status_code"] = code
                t["end_line"] = lineno

    return tasks


# ---------------------------------------------------------------------------
# Mode Register / ODT extraction
# ---------------------------------------------------------------------------
# Final per-rank MR table printed during "SAGV Finalization":
#   MC0.C0.R0	Data	Delay (nCK)
#    MR 34:		0x14	 12
_MR_TABLE_HDR = re.compile(r"^(MC\d+\.C\d+\.R\d+)\s+Data\s+Delay", re.IGNORECASE)
_MR_TABLE_ROW = re.compile(r"^MR\s*(\d+)\s*:\s+0x([0-9A-Fa-f]+)")
# JEDEC-reset MR writes: " MC0 C0 R0 MrAddr =  34 MrIndex = 31 Value = 0x24 ..."
_MR_INIT_START = re.compile(r"^InitMrwDdr5\s*:", re.IGNORECASE)
_MR_INIT_ROW = re.compile(
    r"^MC(\d+)\s+C(\d+)\s+R(\d+)\s+MrAddr\s*=\s*(\d+)\s+MrIndex\s*=\s*\d+\s+"
    r"Value\s*=\s*0x([0-9A-Fa-f]+)",
    re.IGNORECASE,
)
_ODT_SUM_START = re.compile(r"^DIMM\s+ODT\s+summary", re.IGNORECASE)
_ODT_SUM_ROW = re.compile(r"^(Mc\d+C\d+D\d+)\s*:\s*(.+)$")
_CPU_RD_ODT = re.compile(
    r"CPU\s+Summary:\s*MC\s*=\s*(\d+)\s+Channel\s*=\s*(\d+)\s+Read\s+ODT\s*=\s*(\d+)",
    re.IGNORECASE,
)
_ODTL_ROW = re.compile(r"^(Mc\d+)\.Ch(\d+)\.(R\d+)\s*:\s*((?:Ddr5\w+\s*:\s*-?\d+\s*)+)$")
_ODTL_KV = re.compile(r"Ddr5(\w+?)\s*:\s*(-?\d+)")
_ODT_INPUT_HDR = re.compile(r"^(Dimm\d+)\s+ODT\s+Values\s*:", re.IGNORECASE)
_ODT_INPUT_ROW = re.compile(r"^([A-Za-z][\w ]*?)\s*:\s*(\S.*)$")
_RCOMP_RDODT = re.compile(r"RcompTarget\[RdOdt\]\s*:\s*(\d+)")
_DDRIO_ODT_MODE = re.compile(r"DDRIO\s+ODT\s+Mode\s*:\s*(\S+)", re.IGNORECASE)
_PRE_TRAINING_START = re.compile(r"MRC\s+task\s+--\s+Pre-Training\s+--\s+Started", re.IGNORECASE)
_ANY_TASK_START = re.compile(r"MRC\s+task\s+--\s+.+?\s+--\s+Started", re.IGNORECASE)


def _new_mr_snapshot(index: int) -> dict[str, Any]:
    return {
        "index": index, "frequency": None, "gear": None,
        "final_mrs": {}, "init_mrs": {},
        "odt_fields": [], "odt_summary": {}, "cpu_read_odt": {}, "odtl": {},
    }


def parse_mr_odt_info(text: str, source_name: str) -> dict[str, Any]:
    """Extract DDR5 Mode Register and ODT settings from an MRC log.

    Returns ``{"source_file", "odt_inputs", "rcomp_rd_odt", "ddrio_odt_mode",
    "snapshots": [...]}``. One snapshot is produced per MRC training pass
    (split at "Pre-Training -- Started") and holds:

      * ``final_mrs``   — rank -> {MR: value}, from the SAGV Finalization table
      * ``init_mrs``    — rank -> {MR: value}, from InitMrwDdr5 (JEDEC reset)
      * ``odt_fields`` / ``odt_summary`` — "DIMM ODT summary" (Ohm)
      * ``cpu_read_odt`` — "Mc0.C0" -> CPU read ODT (Ohm)
      * ``odtl``        — rank -> ODT latency offsets (DDR5 ODT Timing Config)

    Rank labels are normalised to the RMT style (``Mc0.C0.R0``). Only
    snapshots holding data are returned.
    """
    info: dict[str, Any] = {
        "source_file": source_name,
        "odt_inputs": {},
        "rcomp_rd_odt": None,
        "ddrio_odt_mode": None,
        "snapshots": [],
    }
    snaps: list[dict[str, Any]] = []
    cur: dict[str, Any] | None = None
    freq: int | None = None
    gear: int | None = None

    mode: str | None = None          # "init" | "mrtable" | "odtsum" | "odtin"
    table_rank: str | None = None
    odt_in_dimm: str | None = None
    odt_hdr_pending = False

    def _snap() -> dict[str, Any]:
        nonlocal cur
        if cur is None:
            cur = _new_mr_snapshot(len(snaps) + 1)
            snaps.append(cur)
        return cur

    def _tag(s: dict[str, Any]) -> None:
        s["frequency"] = freq
        s["gear"] = gear

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            if mode in ("odtin",):
                mode = None
            continue

        fg = FREQ_GEAR_PATTERN.search(line)
        if fg:
            freq = int(fg.group(1))
            gear = (int(fg.group(2)) + 1) * 2

        if _PRE_TRAINING_START.search(line):
            cur = _new_mr_snapshot(len(snaps) + 1)
            snaps.append(cur)
            mode = None
            continue

        # ── Continue an open multi-line section ──────────────────────
        if mode == "mrtable":
            m = _MR_TABLE_ROW.match(line)
            if m and table_rank:
                _snap()["final_mrs"].setdefault(table_rank, {})[int(m.group(1))] = int(m.group(2), 16)
                continue
            if line.startswith("->"):
                continue
            hdr = _MR_TABLE_HDR.match(line)
            if hdr:
                table_rank = "Mc" + hdr.group(1)[2:]
                _tag(_snap())
                continue
            mode = None
        elif mode == "init":
            m = _MR_INIT_ROW.match(line)
            if m:
                rank = f"Mc{m.group(1)}.C{m.group(2)}.R{m.group(3)}"
                _snap()["init_mrs"].setdefault(rank, {})[int(m.group(4))] = int(m.group(5), 16)
                continue
            if "MRC task" in line:
                mode = None
            else:
                continue  # other prints (e.g. "Ddr5OdtTable...") inside the init block
        elif mode == "odtsum":
            s = _snap()
            if odt_hdr_pending:
                s["odt_fields"] = line.split()
                odt_hdr_pending = False
                continue
            m = _ODT_SUM_ROW.match(line)
            if m:
                vals = m.group(2).split()
                s["odt_summary"][m.group(1)] = {
                    f: (int(v) if v.lstrip("-").isdigit() else v)
                    for f, v in zip(s["odt_fields"], vals)
                }
                continue
            m = _CPU_RD_ODT.search(line)
            if m:
                s["cpu_read_odt"][f"Mc{m.group(1)}.C{m.group(2)}"] = int(m.group(3))
                continue
            mode = None
        elif mode == "odtin":
            m = _ODT_INPUT_ROW.match(line)
            if m and not _ODT_INPUT_HDR.match(line):
                info["odt_inputs"].setdefault(odt_in_dimm, {})[m.group(1).strip()] = m.group(2).strip()
                continue
            mode = None

        # ── Section starts / single-line facts ───────────────────────
        hdr = _MR_TABLE_HDR.match(line)
        if hdr:
            mode = "mrtable"
            table_rank = "Mc" + hdr.group(1)[2:]
            _tag(_snap())
            continue
        if _MR_INIT_START.match(line):
            mode = "init"
            continue
        if _ODT_SUM_START.match(line):
            mode = "odtsum"
            odt_hdr_pending = True
            _tag(_snap())
            continue
        m = _ODT_INPUT_HDR.match(line)
        if m:
            odt_in_dimm = m.group(1)
            if odt_in_dimm not in info["odt_inputs"]:
                mode = "odtin"
                info["odt_inputs"][odt_in_dimm] = {}
            continue
        m = _ODTL_ROW.match(line)
        if m:
            rank = f"{m.group(1)}.C{m.group(2)}.{m.group(3)}"
            _snap()["odtl"][rank] = {k: int(v) for k, v in _ODTL_KV.findall(m.group(4))}
            continue
        m = _RCOMP_RDODT.search(line)
        if m and int(m.group(1)) != 0:
            info["rcomp_rd_odt"] = int(m.group(1))
            continue
        m = _DDRIO_ODT_MODE.search(line)
        if m:
            info["ddrio_odt_mode"] = m.group(1)

    info["snapshots"] = [
        s for s in snaps
        if s["final_mrs"] or s["init_mrs"] or s["odt_summary"] or s["odtl"]
    ]
    for n, s in enumerate(info["snapshots"], start=1):
        s["index"] = n
    return info


# ── JEDEC DDR5 (JESD79-5) Mode Register decoding ────────────────────────
# MR5/MR34/MR35/MR37-39 decodes were cross-checked against MRC's own
# "DIMM ODT summary" and "DDR5 ODT Timing Config" prints.
_RTT_OHMS = {0: "Off", 1: "240", 2: "120", 3: "80", 4: "60", 5: "48", 6: "40", 7: "34"}
_CA_ODT_OHMS = {0: "Off", 1: "480", 2: "240", 3: "120", 4: "80", 5: "60", 6: "RFU", 7: "40"}
_RON_OHMS = {0: "34", 1: "40", 2: "48", 3: "RFU"}
_ODTL_ON = {1: -4, 2: -3, 3: -2, 4: -1, 5: 0, 6: 1}
_ODTL_OFF = {1: 4, 2: 3, 3: 2, 4: 1, 5: 0, 6: -1}

MR_DESCRIPTIONS: dict[int, str] = {
    0: "Burst Length & CAS Latency",
    2: "Functional Modes",
    3: "DQS Training",
    4: "Refresh Settings",
    5: "IO Settings (Ron)",
    6: "tWR / tRTP",
    8: "Preamble / Postamble",
    10: "VrefDQ",
    11: "VrefCA",
    12: "VrefCS",
    13: "tCCD_L",
    32: "CK / CS ODT",
    33: "CA ODT / DQS_RTT_PARK",
    34: "RTT_PARK / RTT_WR",
    35: "RTT_NOM_WR / RTT_NOM_RD",
    36: "RTT Loopback",
    37: "ODTL WR offsets",
    38: "ODTL NT WR offsets",
    39: "ODTL NT RD offsets",
}


def _ohm(code_map: dict[int, str], code: int) -> str:
    v = code_map.get(code, "?")
    return v if v in ("Off", "RFU", "?") else f"{v}\u03a9"


def decode_ddr5_mr(mr: int, val: int) -> str:
    """Human-readable decode of the DDR5 MRs relevant to margin debug.

    Returns an empty string for MRs that are not decoded.
    """
    if mr == 0:
        bl = {0: "BL16", 1: "BC8 OTF", 2: "BL32", 3: "BL32 OTF"}[val & 0x3]
        return f"CL={22 + 2 * ((val >> 2) & 0x1F)}, {bl}"
    if mr == 5:
        txt = f"Ron PU={_ohm(_RON_OHMS, (val >> 1) & 0x3)}, PD={_ohm(_RON_OHMS, (val >> 6) & 0x3)}"
        return txt + (", DQ output disabled" if val & 0x1 else "")
    if mr in (10, 11, 12):
        name = {10: "VrefDQ", 11: "VrefCA", 12: "VrefCS"}[mr]
        code = val & 0x7F
        if code > 0x7D:
            return f"{name} code 0x{code:02X} (RFU)"
        return f"{name}={97.5 - 0.5 * code:.1f}% VDDQ"
    if mr == 32:
        return f"CK ODT={_ohm(_CA_ODT_OHMS, val & 0x7)}, CS ODT={_ohm(_CA_ODT_OHMS, (val >> 3) & 0x7)}"
    if mr == 33:
        return f"CA ODT={_ohm(_CA_ODT_OHMS, val & 0x7)}, DQS_RTT_PARK={_ohm(_RTT_OHMS, (val >> 3) & 0x7)}"
    if mr == 34:
        return f"RTT_WR={_ohm(_RTT_OHMS, (val >> 3) & 0x7)}, RTT_PARK={_ohm(_RTT_OHMS, val & 0x7)}"
    if mr == 35:
        return f"RTT_NOM_WR={_ohm(_RTT_OHMS, val & 0x7)}, RTT_NOM_RD={_ohm(_RTT_OHMS, (val >> 3) & 0x7)}"
    if mr == 36:
        return f"RTT_LOOPBACK={_ohm(_RTT_OHMS, val & 0x7)}"
    if mr in (37, 38, 39):
        on = _ODTL_ON.get(val & 0x7)
        off = _ODTL_OFF.get((val >> 3) & 0x7)
        on_s = f"{on:+d}" if on is not None else "RFU"
        off_s = f"{off:+d}" if off is not None else "RFU"
        return f"ODTL on {on_s} / off {off_s} tCK"
    return ""


def _mr_description(mr: int) -> str:
    if mr in MR_DESCRIPTIONS:
        return MR_DESCRIPTIONS[mr]
    if 128 <= mr <= 255:
        return "Per-DQ settings"
    if mr >= 256:
        return "Extended / vendor"
    return ""


# ---------------------------------------------------------------------------
# Statistics helpers (shared by professional PPT and HTML report)
# ---------------------------------------------------------------------------

def _get_threshold(param: str, axis_config: dict | None) -> float:
    """Return the positive reference-line threshold for *param*.

    Uses the per-parameter +Ref, then the default +Ref, then 10."""
    return abs(float(_param_ref(param, axis_config, is_plus=True)))


def _get_threshold_minus(param: str, axis_config: dict | None) -> float:
    """Return the negative reference-line threshold for *param*.

    Uses the per-parameter -Ref, then the default -Ref, then -10. Used to
    flag raw values whose negative margin is weaker than the configured -Ref
    (i.e. the value lies above -Ref, closer to 0)."""
    return -abs(float(_param_ref(param, axis_config, is_plus=False)))


def _margin_status(avg_width: float, threshold: float) -> str:
    """Classify a margin as PASS / WARN / FAIL relative to *threshold*."""
    if avg_width >= 2 * threshold:
        return "PASS"
    if avg_width >= threshold:
        return "WARN"
    return "FAIL"


def _compute_param_stats(
    rows: list[dict[str, Any]], params: list[str]
) -> dict[str, dict[str, float]]:
    """Per-parameter summary statistics across *rows*."""
    result: dict[str, dict] = {}
    for param in params:
        plus_v  = [float(r[f"{param}+"]) for r in rows if r.get(f"{param}+") is not None]
        minus_v = [float(r[f"{param}-"]) for r in rows if r.get(f"{param}-") is not None]
        widths  = [w for w in (calc_window_width(r, param) for r in rows) if w is not None]
        if not widths:
            continue
        result[param] = {
            "count":     len(widths),
            "avg_plus":  statistics.mean(plus_v)  if plus_v  else 0.0,
            "avg_minus": statistics.mean(minus_v) if minus_v else 0.0,
            "avg_width": statistics.mean(widths),
            "min_width": min(widths),
            "max_width": max(widths),
            "median":    statistics.median(widths),
            "std_width": statistics.stdev(widths) if len(widths) > 1 else 0.0,
        }
    return result


def _compute_freq_stats(
    rows: list[dict[str, Any]], params: list[str]
) -> dict[Any, dict[str, dict]]:
    """Per-parameter statistics grouped by Frequency field."""
    by_freq: dict = defaultdict(list)
    for r in rows:
        by_freq[r.get("Frequency")].append(r)
    return {
        freq: _compute_param_stats(freq_rows, params)
        for freq, freq_rows in sorted(
            by_freq.items(), key=lambda kv: (kv[0] is None, kv[0])
        )
    }


# ---------------------------------------------------------------------------
# Professional PowerPoint
# ---------------------------------------------------------------------------

def build_ppt_professional(
    path: Path,
    rows: list[dict[str, Any]],
    chart_params: list[str],
    axis_config: dict | None = None,
    platform_infos: list[dict] | None = None,
    training_steps: list[dict] | None = None,
) -> None:
    """Build a professional RMT Summary PowerPoint with tables and charts."""
    if Presentation is None or CategoryChartData is None or XL_CHART_TYPE is None:
        raise RuntimeError("python-pptx is not installed. Run: pip install python-pptx")

    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    try:
        from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE as _MSAT
        _RECT = _MSAT.RECTANGLE
    except Exception:
        _RECT = 1  # integer fallback (MSO rectangle)

    # ── Colour palette ────────────────────────────────────────────────
    BLUE    = RGBColor(0x00, 0x71, 0xC5)
    NAVY    = RGBColor(0x1E, 0x29, 0x3B)
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
    LTGRAY  = RGBColor(0xF3, 0xF4, 0xF6)
    GRAY    = RGBColor(0x6B, 0x72, 0x80)
    PASS_FG = RGBColor(0x06, 0x5F, 0x46)
    PASS_BG = RGBColor(0xD1, 0xFA, 0xE5)
    WARN_FG = RGBColor(0x78, 0x35, 0x00)
    WARN_BG = RGBColor(0xFF, 0xF3, 0xCD)
    FAIL_FG = RGBColor(0x7F, 0x1D, 0x1D)
    FAIL_BG = RGBColor(0xFE, 0xE2, 0xE2)
    ALT     = RGBColor(0xEF, 0xF6, 0xFF)
    HDR     = RGBColor(0x1E, 0x40, 0xAF)
    ACCENT  = RGBColor(0xBF, 0xDB, 0xF7)
    S_BLUES = [
        RGBColor(0x00, 0x71, 0xC5), RGBColor(0x00, 0xA3, 0xE0),
        RGBColor(0x00, 0x5B, 0x99), RGBColor(0x4D, 0xB6, 0xFF),
        RGBColor(0x00, 0x30, 0x5E), RGBColor(0x66, 0xC5, 0xFF),
    ]

    # ── Helpers ───────────────────────────────────────────────────────
    def add_banner(slide, title_text: str, subtitle_text: str = "") -> None:
        sh = slide.shapes.add_shape(_RECT, 0, 0, prs.slide_width, Inches(1.12))
        sh.fill.solid(); sh.fill.fore_color.rgb = BLUE
        sh.line.fill.background()
        tf = sh.text_frame; tf.word_wrap = False
        tf.margin_left = Inches(0.35); tf.margin_top = Inches(0.14)
        p = tf.paragraphs[0]
        r = p.add_run(); r.text = title_text
        r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = WHITE
        if subtitle_text:
            p2 = tf.add_paragraph()
            r2 = p2.add_run(); r2.text = subtitle_text
            r2.font.size = Pt(10); r2.font.color.rgb = ACCENT

    def set_cell(cell, text, bg=None, fg=None, bold=False, sz=9,
                 align=PP_ALIGN.CENTER) -> None:
        from pptx.oxml.ns import qn as _qn
        tf = cell.text_frame; tf.word_wrap = False
        txb = tf._txBody
        for rn in txb.findall(".//" + _qn("a:r")):
            rn.getparent().remove(rn)
        for xp in txb.findall(_qn("a:p"))[1:]:
            txb.remove(xp)
        para = tf.paragraphs[0]; para.alignment = align
        run = para.add_run()
        run.text = str(text) if text is not None else ""
        run.font.size = Pt(sz); run.font.bold = bold
        run.font.color.rgb = fg if fg else NAVY
        if bg:
            cell.fill.solid(); cell.fill.fore_color.rgb = bg

    def set_hdr_row(table, headers: list[str]) -> None:
        for ci, h in enumerate(headers):
            set_cell(table.cell(0, ci), h, bg=HDR, fg=WHITE, bold=True, sz=8)

    def fmt(val, dec: int = 2) -> str:
        try:
            return f"{float(val):.{dec}f}"
        except (TypeError, ValueError):
            return "—"

    def status_style(status: str) -> tuple:
        return {"PASS": (PASS_BG, PASS_FG), "WARN": (WARN_BG, WARN_FG),
                "FAIL": (FAIL_BG, FAIL_FG)}.get(status, (LTGRAY, GRAY))

    # ── Data ──────────────────────────────────────────────────────────
    param_stats = _compute_param_stats(rows, chart_params)
    freq_stats  = _compute_freq_stats(rows, chart_params)
    freqs       = list(freq_stats.keys())
    thresholds  = {p: _get_threshold(p, axis_config) for p in chart_params}
    files       = sorted({str(r.get("SourceFile", "")) for r in rows})
    freq_disp   = ", ".join(str(f) for f in freqs if f is not None) or "N/A"
    from datetime import datetime as _dt
    generated = _dt.now().strftime("%Y-%m-%d %H:%M")

    # ── Presentation setup ────────────────────────────────────────────
    prs = Presentation()
    prs.slide_width  = int(13.33 * 914400)
    prs.slide_height = int(7.50  * 914400)
    blank = next(
        (l for l in prs.slide_layouts if "blank" in l.name.lower()),
        prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0],
    )

    # ==============================================================
    # Slide 1: Title
    # ==============================================================
    s1 = prs.slides.add_slide(blank)
    bg_sh = s1.shapes.add_shape(_RECT, 0, 0, prs.slide_width, prs.slide_height)
    bg_sh.fill.solid(); bg_sh.fill.fore_color.rgb = NAVY; bg_sh.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(0.9), Inches(1.5), Inches(11.5), Inches(1.9))
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = f"{TOOL_NAME} \u2014 RMT Margin Summary Report"
    r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = WHITE
    p_sub = tb.text_frame.add_paragraph()
    r = p_sub.add_run(); r.text = TOOL_SUBTITLE
    r.font.size = Pt(16); r.font.color.rgb = ACCENT

    acc = s1.shapes.add_shape(_RECT, Inches(0.9), Inches(3.5), Inches(4.8), Inches(0.06))
    acc.fill.solid(); acc.fill.fore_color.rgb = BLUE; acc.line.fill.background()

    meta_tb = s1.shapes.add_textbox(Inches(0.9), Inches(3.75), Inches(11.5), Inches(2.5))
    mlines = [
        f"Project     : {_ACTIVE_PROJECT_NAME or 'N/A'}",
        f"Generated   : {generated}",
        f"Source files: {len(files)}",
        f"Frequencies : {freq_disp}",
        f"Total rows  : {len(rows)}",
        f"Tool        : {TOOL_NAME} {TOOL_VERSION}",
    ]
    for i, line in enumerate(mlines):
        para = (meta_tb.text_frame.paragraphs[0]
                if i == 0 else meta_tb.text_frame.add_paragraph())
        rn = para.add_run(); rn.text = line
        rn.font.size = Pt(13); rn.font.color.rgb = ACCENT

    # ==============================================================
    # Slide 2: Test Configuration (Board / DIMM info)
    # ==============================================================
    if platform_infos:
        sp = prs.slides.add_slide(blank)
        add_banner(sp, "Test Configuration",
                   "Platform details and memory module population")

        boards = sorted({pi.get("board") for pi in platform_infos if pi.get("board")})
        cpus   = sorted({pi.get("cpu")   for pi in platform_infos if pi.get("cpu")})
        phys   = sorted({pi.get("phy_version") for pi in platform_infos
                         if pi.get("phy_version")})

        sys_items = [
            ("Board / RVP",      ", ".join(boards) if boards else "N/A"),
            ("CPU",              ", ".join(cpus)   if cpus   else "N/A"),
            ("PHY IP Version",   ", ".join(phys)   if phys   else "N/A"),
            ("Memory Technology", "DDR5"),
            ("Source Files",     str(len(platform_infos))),
        ]
        n_sys = len(sys_items) + 1
        st = sp.shapes.add_table(
            n_sys, 2, Inches(0.25), Inches(1.25),
            Inches(4.9), Inches(min(3.5, 0.42 * n_sys)),
        ).table
        for ri in range(n_sys):
            st.cell(ri, 0).width = Inches(1.9)
            st.cell(ri, 1).width = Inches(3.0)
        set_hdr_row(st, ["System Info", ""])
        for ri, (k, v) in enumerate(sys_items, 1):
            rbg = ALT if ri % 2 == 0 else None
            set_cell(st.cell(ri, 0), k, bg=rbg, fg=NAVY, bold=True,
                     sz=9, align=PP_ALIGN.LEFT)
            set_cell(st.cell(ri, 1), v, bg=rbg, fg=BLUE, sz=9,
                     align=PP_ALIGN.LEFT)

        # Collect unique populated DIMMs
        seen_pn: set = set()
        unique_dimms: list = []
        for pi in platform_infos:
            for dimm in pi.get("dimms", []):
                if dimm.get("populated"):
                    pn = dimm.get("part_number") or dimm.get("slot", "")
                    if pn not in seen_pn:
                        seen_pn.add(pn)
                        unique_dimms.append(dimm)
        if not unique_dimms:   # fallback: first file's dimms
            for pi in platform_infos[:1]:
                unique_dimms = [d for d in pi.get("dimms", []) if d.get("populated")]

        if unique_dimms:
            d_hdrs = ["Slot", "Type", "Manufacturer", "Part Number",
                      "Size", "Ranks", "Width", "Date"]
            n_d = len(unique_dimms) + 1
            dt = sp.shapes.add_table(
                n_d, 8, Inches(5.25), Inches(1.25),
                Inches(7.8), Inches(min(5.9, 0.42 * n_d)),
            ).table
            for ci, cw in enumerate([0.8, 1.4, 1.3, 2.2, 0.65, 0.6, 0.6, 0.75]):
                for ri in range(n_d):
                    dt.cell(ri, ci).width = Inches(cw)
            set_hdr_row(dt, d_hdrs)
            for ri, dimm in enumerate(unique_dimms, 1):
                rbg  = ALT if ri % 2 == 0 else None
                size = (f'{dimm.get("size_mb", 0) // 1024} GB'
                        if dimm.get("size_mb") else "?")
                width_str = (f'{dimm.get("sdram_width","?")}b'
                             if dimm.get("sdram_width") else "?")
                set_cell(dt.cell(ri, 0), dimm.get("slot", "?"),
                         bg=rbg, fg=NAVY, sz=8)
                set_cell(dt.cell(ri, 1), dimm.get("module_type", "DDR5") or "DDR5",
                         bg=rbg, fg=NAVY, sz=7)
                set_cell(dt.cell(ri, 2), dimm.get("manufacturer", "?") or "?",
                         bg=rbg, fg=NAVY, sz=8)
                set_cell(dt.cell(ri, 3),
                         dimm.get("part_number") or dimm.get("dram_id_code") or "N/A",
                         bg=rbg, fg=BLUE, bold=True, sz=8, align=PP_ALIGN.LEFT)
                set_cell(dt.cell(ri, 4), size,
                         bg=rbg, fg=NAVY, sz=8)
                set_cell(dt.cell(ri, 5), str(dimm.get("ranks", "?")),
                         bg=rbg, fg=NAVY, sz=8)
                set_cell(dt.cell(ri, 6), width_str,
                         bg=rbg, fg=NAVY, sz=8)
                set_cell(dt.cell(ri, 7),
                         dimm.get("date_display") or dimm.get("date_code") or "?",
                         bg=rbg, fg=GRAY, sz=8)

    # ==============================================================
    # Slide 3: Training Steps Summary (if available)
    # ==============================================================
    if training_steps:
        from collections import defaultdict as _dd

        total_ts  = len(training_steps)
        passed_ts = sum(1 for t in training_steps if t["status"] == "SUCCEEDED")
        failed_ts = sum(1 for t in training_steps if t["status"] == "FAILED")
        by_iter: dict[int, list] = _dd(list)
        for t in training_steps:
            by_iter[t["iteration"]].append(t)
        n_iters = len(by_iter)

        # Identify unique task names across all iterations for summary
        task_names: list[str] = []
        seen_names: set = set()
        for t in training_steps:
            if t["name"] not in seen_names:
                seen_names.add(t["name"])
                task_names.append(t["name"])

        st_slide = prs.slides.add_slide(blank)
        add_banner(st_slide, "MRC Training Steps Summary",
                   f"Total tasks: {total_ts}  |  "
                   f"Passed: {passed_ts}  |  "
                   f"Failed: {failed_ts}  |  "
                   f"Iterations: {n_iters}")

        # ── Left: overall summary table ──
        sum_rows = [
            ("Total Tasks",   str(total_ts)),
            ("Succeeded",     str(passed_ts)),
            ("Failed",        str(failed_ts)),
            ("Iterations",    str(n_iters)),
            ("Task Types",    f"{sum(1 for t in training_steps if 'GREEN' in t['task_type'])} GREEN  /  {sum(1 for t in training_steps if 'GREEN' not in t['task_type'])} BLUE"),
        ]
        failed_tasks = [(t["name"], t["iteration"], t.get("status_code") or "?")
                        for t in training_steps if t["status"] == "FAILED"]
        if failed_tasks:
            sum_rows.append(("⚠ Failed Tasks", ""))
            for fn, fi, fc in failed_tasks[:6]:
                sum_rows.append((f"  Iter {fi+1}: {fn}", f"Status={fc}"))

        n_sr = len(sum_rows) + 1
        sum_tbl = st_slide.shapes.add_table(
            n_sr, 2, Inches(0.25), Inches(1.25),
            Inches(5.0), Inches(min(5.8, 0.38 * n_sr)),
        ).table
        for ri in range(n_sr):
            sum_tbl.cell(ri, 0).width = Inches(2.6)
            sum_tbl.cell(ri, 1).width = Inches(2.4)
        set_hdr_row(sum_tbl, ["Category", "Value"])
        FAIL_BG = RGBColor(0xFF, 0xDD, 0xDD)
        for ri, (k, v) in enumerate(sum_rows, 1):
            is_fail_row = k.startswith("⚠") or (failed_tasks and ri > 5 and k.startswith("  Iter"))
            bg = FAIL_BG if is_fail_row else (ALT if ri % 2 == 0 else None)
            set_cell(sum_tbl.cell(ri, 0), k, bg=bg, fg=NAVY, bold=True,
                     sz=8, align=PP_ALIGN.LEFT)
            set_cell(sum_tbl.cell(ri, 1), v, bg=bg, fg=BLUE, sz=8,
                     align=PP_ALIGN.LEFT)

        # ── Right: per-iteration pass/fail table ──
        iter_hdrs = ["Iter", "Tasks", "✓ Pass", "✗ Fail", "Status"]
        n_it = len(by_iter) + 1
        it_tbl = st_slide.shapes.add_table(
            n_it, 5, Inches(5.35), Inches(1.25),
            Inches(7.7), Inches(min(5.8, 0.38 * n_it)),
        ).table
        for ci, cw in enumerate([0.7, 1.0, 1.0, 1.0, 4.0]):
            for ri in range(n_it):
                it_tbl.cell(ri, ci).width = Inches(cw)
        set_hdr_row(it_tbl, iter_hdrs)
        for ri, iter_num in enumerate(sorted(by_iter.keys()), 1):
            it = by_iter[iter_num]
            ip = sum(1 for x in it if x["status"] == "SUCCEEDED")
            if_ = sum(1 for x in it if x["status"] == "FAILED")
            fail_names = "; ".join(x["name"] for x in it if x["status"] == "FAILED")
            bg = FAIL_BG if if_ else (ALT if ri % 2 == 0 else None)
            set_cell(it_tbl.cell(ri, 0), str(iter_num), bg=bg, fg=NAVY, sz=9)
            set_cell(it_tbl.cell(ri, 1), str(len(it)),   bg=bg, fg=NAVY, sz=9)
            set_cell(it_tbl.cell(ri, 2), str(ip),         bg=bg, fg=RGBColor(0,128,0) if ip else NAVY, bold=(ip > 0), sz=9)
            set_cell(it_tbl.cell(ri, 3), str(if_),        bg=bg, fg=RGBColor(0xC0,0,0) if if_ else NAVY, bold=(if_ > 0), sz=9)
            status_str = ("FAILED: " + fail_names) if if_ else "PASSED"
            set_cell(it_tbl.cell(ri, 4), status_str,      bg=bg, fg=RGBColor(0xC0,0,0) if if_ else RGBColor(0,128,0), sz=8, align=PP_ALIGN.LEFT)

    # ==============================================================
    # Slide 4: Summary Statistics Table
    # ==============================================================
    s2 = prs.slides.add_slide(blank)
    add_banner(s2, "Summary Statistics",
               "Average margins across all ranks and frequencies")

    hdrs = ["Parameter", "Avg +", "Avg −", "Avg Width",
            "Min Width", "Max Width", "Std Dev", "Threshold", "Status"]
    nr = len(chart_params) + 1
    tbl = s2.shapes.add_table(
        nr, 9, Inches(0.25), Inches(1.25),
        Inches(12.8), Inches(min(5.9, 0.44 * nr)),
    ).table
    for ci, cw in enumerate([1.7, 0.9, 0.9, 1.15, 1.15, 1.15, 1.05, 1.1, 0.9]):
        for ri in range(nr):
            tbl.cell(ri, ci).width = Inches(cw)
    set_hdr_row(tbl, hdrs)

    for ri, param in enumerate(chart_params, 1):
        s   = param_stats.get(param, {})
        aw  = s.get("avg_width", 0.0)
        thr = thresholds[param]
        st  = _margin_status(aw, thr)
        sbg, sfg = status_style(st)
        rbg = ALT if ri % 2 == 0 else None
        set_cell(tbl.cell(ri, 0), param,              bg=rbg, fg=NAVY, bold=True, sz=9, align=PP_ALIGN.LEFT)
        set_cell(tbl.cell(ri, 1), fmt(s.get("avg_plus")),  bg=rbg, fg=PASS_FG, sz=9)
        set_cell(tbl.cell(ri, 2), fmt(s.get("avg_minus")), bg=rbg, fg=FAIL_FG, sz=9)
        set_cell(tbl.cell(ri, 3), fmt(aw),                 bg=rbg, fg=NAVY,    bold=True, sz=9)
        set_cell(tbl.cell(ri, 4), fmt(s.get("min_width")), bg=rbg, fg=GRAY,    sz=9)
        set_cell(tbl.cell(ri, 5), fmt(s.get("max_width")), bg=rbg, fg=GRAY,    sz=9)
        set_cell(tbl.cell(ri, 6), fmt(s.get("std_width")), bg=rbg, fg=GRAY,    sz=9)
        set_cell(tbl.cell(ri, 7), fmt(thr, 0),             bg=rbg, fg=GRAY,    sz=9)
        set_cell(tbl.cell(ri, 8), st,                      bg=sbg, fg=sfg,     bold=True, sz=9)

    # ==============================================================
    # Slide 3: Frequency Comparison Bar Chart
    # ==============================================================
    s3 = prs.slides.add_slide(blank)
    add_banner(s3, "Avg Window Width — Frequency Comparison",
               "One series per frequency  |  Higher = better margin")
    if param_stats:
        cd = CategoryChartData()
        cd.categories = [p for p in chart_params if p in param_stats]
        for freq, fst in freq_stats.items():
            lbl = f"{freq} MT/s" if freq is not None else "Unknown"
            cd.add_series(
                lbl,
                [round(fst.get(p, {}).get("avg_width", 0.0), 2)
                 for p in chart_params if p in param_stats],
            )
        cf = s3.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.35), Inches(1.25), Inches(12.6), Inches(5.9), cd,
        )
        ch = cf.chart
        ch.has_legend = True
        ch.value_axis.has_major_gridlines = True
        try:
            ch.value_axis.has_title = True
            ch.value_axis.axis_title.text_frame.text = "Avg Window Width"
        except Exception:
            pass
        for i, ser in enumerate(ch.series):
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = S_BLUES[i % len(S_BLUES)]

    # ==============================================================
    # Slide 4+: Per-Frequency Rank Detail Tables
    # ==============================================================
    for freq, fst in freq_stats.items():
        freq_label = f"{freq} MT/s" if freq is not None else "Unknown"
        sf = prs.slides.add_slide(blank)
        add_banner(sf, f"Frequency {freq_label} — Rank-Level Detail",
                   "Each cell: Avg + / Avg − / Avg Width.  Cell colour = Pass/Warn/Fail vs threshold.")
        freq_rows_f = [r for r in rows if r.get("Frequency") == freq]
        by_rank: dict = defaultdict(list)
        for r in freq_rows_f:
            by_rank[str(r.get("Params", ""))].append(r)
        ranks = sorted(by_rank.keys())
        show  = chart_params[:min(6, len(chart_params))]
        nr2 = len(ranks) + 1
        nc2 = len(show) + 1
        if nr2 < 2:
            continue
        th2 = min(Inches(5.9), Inches(0.42 * nr2))
        t2 = sf.shapes.add_table(nr2, nc2, Inches(0.25), Inches(1.25),
                                  Inches(12.8), th2).table
        cw0 = 1.6
        cw_r = round(11.2 / max(len(show), 1), 2)
        for ri in range(nr2):
            t2.cell(ri, 0).width = Inches(cw0)
        for ci in range(1, nc2):
            for ri in range(nr2):
                t2.cell(ri, ci).width = Inches(cw_r)
        set_hdr_row(t2, ["Rank"] + [f"{p}\n+/−/W" for p in show])

        for ri, rank in enumerate(ranks, 1):
            rrows = by_rank[rank]
            rbg   = ALT if ri % 2 == 0 else None
            set_cell(t2.cell(ri, 0), rank, bg=rbg, fg=NAVY, bold=True,
                     sz=8, align=PP_ALIGN.LEFT)
            for ci, param in enumerate(show, 1):
                pv = [float(r[f"{param}+"]) for r in rrows
                      if r.get(f"{param}+") is not None]
                mv = [float(r[f"{param}-"]) for r in rrows
                      if r.get(f"{param}-") is not None]
                wv = [w for w in (calc_window_width(r, param) for r in rrows)
                      if w is not None]
                ap = statistics.mean(pv) if pv else None
                am = statistics.mean(mv) if mv else None
                aw2 = statistics.mean(wv) if wv else None
                thr2 = thresholds[param]
                st2  = _margin_status(aw2, thr2) if aw2 is not None else "N/A"
                cbg, cfg2 = status_style(st2) if aw2 is not None else (rbg, GRAY)
                set_cell(t2.cell(ri, ci),
                         f"{fmt(ap,1)} / {fmt(am,1)} / {fmt(aw2,1)}",
                         bg=cbg, fg=cfg2, sz=7)

    # ==============================================================
    # Last slide: RunTemp Drift (only if multiple run temps)
    # ==============================================================
    run_temps_f = []
    for r in rows:
        rt = r.get("RunTemp")
        if rt is None:
            continue
        try:
            run_temps_f.append(float(rt))
        except (TypeError, ValueError):
            pass
    run_temps_f = sorted(set(run_temps_f))

    if len(run_temps_f) > 1 and chart_params:
        srt = prs.slides.add_slide(blank)
        add_banner(srt, "Run Temperature Drift Analysis",
                   "Avg window width per parameter vs run temperature")
        rtd = CategoryChartData()
        rtd.categories = [str(t) for t in run_temps_f]
        for param in chart_params:
            sv = []
            for rt in run_temps_f:
                wl = [calc_window_width(r, param) for r in rows
                      if r.get("RunTemp") is not None
                      and abs(float(r["RunTemp"]) - rt) < 0.5
                      and calc_window_width(r, param) is not None]
                sv.append(round(statistics.mean(wl), 2) if wl else None)
            rtd.add_series(param, sv)
        rtf = srt.shapes.add_chart(
            XL_CHART_TYPE.LINE_MARKERS,
            Inches(0.35), Inches(1.25), Inches(12.6), Inches(5.9), rtd,
        )
        rtc = rtf.chart
        rtc.has_legend = True
        rtc.value_axis.has_major_gridlines = True
        try:
            rtc.value_axis.has_title = True
            rtc.value_axis.axis_title.text_frame.text = "Avg Window Width"
            rtc.category_axis.has_title = True
            rtc.category_axis.axis_title.text_frame.text = "Run Temperature (°C)"
        except Exception:
            pass

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    print(f"Professional PPT saved ({len(prs.slides)} slides): {path}")


# ---------------------------------------------------------------------------
# Interactive HTML Report  (PowerBI-style)
# ---------------------------------------------------------------------------

def _build_training_steps_html(training_steps: list[dict] | None) -> list[str]:
    """Build HTML parts for the Training Steps tab.

    When the report aggregates more than one source log, a file dropdown is
    rendered so the user can pick which file's training steps to view; each
    file's tasks are shown in its own panel and the dropdown toggles them.
    """
    if not training_steps:
        return [
            '<div class="alert alert-info m-4">No training steps found '
            "(log did not contain MRC task markers).</div>"
        ]

    from collections import defaultdict as _dd

    # Category colours  (GREEN MRC = teal, MRC = blue)
    def _type_badge(tt: str) -> str:
        if "GREEN" in tt:
            return '<span class="badge" style="background:#198754">GREEN</span>'
        return '<span class="badge bg-primary">BLUE</span>'

    def _status_badge(s: str) -> str:
        if s == "SUCCEEDED":
            return '<span class="badge bg-success">&#10003; SUCCEEDED</span>'
        if s == "FAILED":
            return '<span class="badge bg-danger">&#10007; FAILED</span>'
        return '<span class="badge bg-warning text-dark">RUNNING</span>'

    def _render_source(steps: list[dict], sid: str) -> list[str]:
        """Render summary cards + per-iteration accordion for one file."""
        total   = len(steps)
        passed  = sum(1 for t in steps if t["status"] == "SUCCEEDED")
        failed  = sum(1 for t in steps if t["status"] == "FAILED")
        n_iters = (max(t["iteration"] for t in steps) + 1) if steps else 0
        by_iter: dict[int, list] = _dd(list)
        for t in steps:
            by_iter[t["iteration"]].append(t)

        out: list[str] = [
            '<div class="row g-3 mb-4">',
            f'<div class="col-6 col-md-3"><div class="border rounded p-3 text-center">'
            f'<div class="h4 fw-bold">{total}</div>'
            f'<div class="small text-muted">Total Tasks</div></div></div>',
            f'<div class="col-6 col-md-3"><div class="border rounded p-3 text-center bg-success bg-opacity-10">'
            f'<div class="h4 fw-bold text-success">{passed}</div>'
            f'<div class="small text-muted">Succeeded</div></div></div>',
            f'<div class="col-6 col-md-3"><div class="border rounded p-3 text-center bg-danger bg-opacity-10">'
            f'<div class="h4 fw-bold text-danger">{failed}</div>'
            f'<div class="small text-muted">Failed</div></div></div>',
            f'<div class="col-6 col-md-3"><div class="border rounded p-3 text-center bg-info bg-opacity-10">'
            f'<div class="h4 fw-bold text-info">{n_iters}</div>'
            f'<div class="small text-muted">Freq/RunTemp Sweeps</div></div></div>',
            '</div>',
            f'<div class="accordion" id="tsAcc{sid}">',
        ]
        for iter_num in sorted(by_iter.keys()):
            it_tasks  = by_iter[iter_num]
            it_fail   = sum(1 for t in it_tasks if t["status"] == "FAILED")
            it_pass   = sum(1 for t in it_tasks if t["status"] == "SUCCEEDED")
            hdr_class = "bg-danger text-white" if it_fail else "bg-success bg-opacity-10"
            badge_txt = f'<span class="badge bg-danger">{it_fail} FAILED</span>' \
                        if it_fail else '<span class="badge bg-success">All Passed</span>'
            label     = f"Iteration {iter_num + 1} &mdash; {it_pass} tasks {badge_txt}"
            acc_id    = f"tsIter{sid}_{iter_num}"
            show_cls  = "show" if iter_num == 0 else ""
            coll_cls  = "" if iter_num == 0 else "collapsed"
            out += [
                f'<div class="accordion-item mb-2 border-0 shadow-sm">',
                f'<h2 class="accordion-header"><button class="accordion-button {coll_cls} py-2 {hdr_class}" '
                f'type="button" data-bs-toggle="collapse" data-bs-target="#{acc_id}">'
                f'{label}</button></h2>',
                f'<div id="{acc_id}" class="accordion-collapse collapse {show_cls}" '
                f'data-bs-parent="#tsAcc{sid}">',
                '<div class="accordion-body p-0">',
                '<div class="table-responsive">',
                '<table class="table table-sm table-hover align-middle mb-0" style="font-size:.8rem">',
                '<thead class="table-dark sticky-top">'
                '<tr><th>#</th><th>Type</th><th>Task Name</th>'
                '<th>Status</th><th>Code</th>'
                '<th>Start Line</th><th>End Line</th><th>Span</th></tr></thead>',
                '<tbody>',
            ]
            for idx, t in enumerate(it_tasks, 1):
                row_cls = (
                    ' class="table-danger"'  if t["status"] == "FAILED"
                    else ' class="table-warning"' if t["status"] == "RUNNING"
                    else ""
                )
                span = (str(t["end_line"] - t["start_line"] + 1)
                        if t["end_line"] else "&#8212;")
                out.append(
                    f'<tr{row_cls}>'
                    f'<td class="text-muted small">{idx}</td>'
                    f'<td>{_type_badge(t["task_type"])}</td>'
                    f'<td class="fw-semibold">{t["name"]}</td>'
                    f'<td>{_status_badge(t["status"])}</td>'
                    f'<td><code class="small">{t["status_code"] or "&#8212;"}</code></td>'
                    f'<td class="text-muted small">{t["start_line"]}</td>'
                    f'<td class="text-muted small">{t["end_line"] or "&#8212;"}</td>'
                    f'<td class="text-muted small">{span}</td>'
                    f'</tr>'
                )
            out += ['</tbody></table></div>', '</div>', '</div>', '</div>']
        out.append('</div>')  # accordion
        return out

    # Group steps by source file
    by_src: dict[str, list] = _dd(list)
    for t in training_steps:
        by_src[str(t.get("source") or "Unknown")].append(t)
    sources = sorted(by_src.keys())

    parts: list[str] = []
    if len(sources) > 1:
        opts = "".join(f'<option value="tsSrc{i}">{s}</option>'
                       for i, s in enumerate(sources))
        parts += [
            '<div class="row mb-3"><div class="col-md-6">',
            '<label class="form-label small fw-semibold mb-1">Select source log:</label>',
            f'<select id="tsSrcSel" class="form-select form-select-sm" '
            f'onchange="tsShowSrc(this.value)">{opts}</select>',
            '</div></div>',
        ]
        for i, s in enumerate(sources):
            disp = "" if i == 0 else "display:none"
            parts.append(f'<div class="ts-src-panel" id="tsSrc{i}" style="{disp}">')
            parts += _render_source(by_src[s], str(i))
            parts.append('</div>')
        parts += [
            "<script>function tsShowSrc(id){"
            "document.querySelectorAll('.ts-src-panel').forEach(p=>p.style.display='none');"
            "var e=document.getElementById(id);if(e)e.style.display='';}</script>",
        ]
    else:
        parts += _render_source(by_src[sources[0]], "0")
    return parts



def _build_platform_html_section(platform_infos: list[dict] | None) -> list[str]:
    """Return a list of HTML strings for the Platform tab body."""
    if not platform_infos:
        return ['<div class="alert alert-info m-4">No platform info available '
                '(log files did not contain MRC board/DIMM output).</div>']

    parts: list[str] = []
    for pi in platform_infos:
        src   = pi.get("source_file", "Unknown")
        board = pi.get("board") or "N/A"
        cpu   = pi.get("cpu") or "N/A"
        phy   = pi.get("phy_version") or "N/A"
        bid   = pi.get("board_id") or "N/A"
        fab   = pi.get("fab_id") or "N/A"

        parts.append('<div class="section-card mb-4">')
        parts.append(f'<h6 class="fw-bold text-primary mb-3">&#x1F4C4; {src}</h6>')

        # System info grid
        # Collect distinct memory technologies from populated DIMMs
        _mem_types = sorted({d.get("module_type") or "DDR5"
                             for d in pi.get("dimms", []) if d.get("populated")})
        mem_tech = " / ".join(_mem_types) if _mem_types else "N/A"

        parts.append('<div class="row g-2 mb-3">')
        for label, value in [
            ("Board / RVP",       board),
            ("CPU",               cpu),
            ("Memory Technology", mem_tech),
            ("Board ID",          bid),
            ("FAB ID",            fab),
            ("PHY IP Version",    phy),
        ]:
            parts.append(
                f'<div class="col-md-4 col-lg-3"><div class="border rounded p-2 bg-light">'
                f'<div class="small text-muted">{label}</div>'
                f'<div class="fw-semibold text-dark">{value}</div>'
                f'</div></div>'
            )
        parts.append('</div>')

        # DIMM population table
        dimms = [d for d in pi.get("dimms", []) if d.get("populated")]
        empty = [d for d in pi.get("dimms", []) if not d.get("populated")]

        if dimms:
            parts.append(
                '<div class="table-responsive">'
                '<table class="table table-sm table-hover align-middle mb-2" style="font-size:.83rem">'
                '<thead class="table-dark">'
                '<tr><th>Slot</th><th>Type</th><th>Manufacturer</th>'
                '<th>Part Number</th><th>Size</th><th>Ranks</th>'
                '<th>SDRAM Width</th><th>Bus Width</th>'
                '<th>Banks×Groups</th><th>ECC</th>'
                '<th>PMIC</th><th>Die Density</th><th>Mfg Date</th></tr>'
                '</thead><tbody>'
            )
            for d in dimms:
                sz   = f'{d.get("size_mb",0)//1024} GB' if d.get("size_mb") else "?"
                ecc  = ("Yes" if d.get("ecc") else "No") if d.get("ecc") is not None else "?"
                bg_x = ("banks", "bank_groups")
                bnk  = (f'{d.get("banks")}×{d.get("bank_groups")}'
                        if d.get("banks") else "?")
                parts.append(
                    f'<tr>'
                    f'<td class="fw-semibold">{d.get("slot","?")}</td>'
                    f'<td><span class="badge rounded-pill" style="background:#0071C5;font-size:.78rem">'
                    f'{d.get("module_type","DDR5") or "DDR5"}</span></td>'
                    f'<td class="text-primary fw-semibold">{d.get("manufacturer","?") or "?"}</td>'
                    f'<td class="fw-bold text-dark">{d.get("part_number","N/A") or "N/A"}</td>'
                    f'<td>{sz}</td>'
                    f'<td>{d.get("ranks","?")}</td>'
                    f'<td>{d.get("sdram_width","?") or "?"}b</td>'
                    f'<td>{d.get("bus_width","?") or "?"}b</td>'
                    f'<td>{bnk}</td>'
                    f'<td>{"<span class=badge bg-success>Yes</span>" if d.get("ecc") else "<span class=badge bg-secondary>No</span>"}</td>'
                    f'<td>{d.get("pmic_type","?") or "?"}</td>'
                    f'<td>{d.get("density","?") or "?"}</td>'
                    f'<td>{d.get("date_display","?") or d.get("date_code","?") or "?"}</td>'
                    f'</tr>'
                )
            parts.append('</tbody></table></div>')

        if empty:
            empty_slots = ", ".join(d.get("slot", "?") for d in empty)
            parts.append(
                f'<div class="text-muted small">Empty slots: {empty_slots}</div>'
            )
        elif not dimms:
            parts.append('<div class="alert alert-warning py-2">No populated DIMMs found.</div>')

        parts.append('</div>')  # section-card

    return parts


def _embed_jmp_charts_html(jmp_charts_dir: Path, ref_note: str = "per-parameter &plusmn;Ref") -> list[str]:
    """Return HTML parts for the JMP Charts tab with base64-embedded PNGs.

    Layout: a comparison workspace of chart panels. Every panel has a
    dropdown listing all JMP charts, so any two (or more) charts can be put
    side-by-side. Panels start pre-populated with the default one-chart-per-
    panel layout (never blank); a column selector, "Add panel" and "Reset"
    are provided. Each PNG is embedded exactly once (base64), keeping the
    HTML self-contained.
    """
    import base64

    def _b64src(p: Path) -> str:
        return "data:image/png;base64," + base64.b64encode(p.read_bytes()).decode("ascii")

    charts: list[tuple[str, Path]] = [
        (param, jmp_charts_dir / f"{param}.png")
        for param in PARAMS
        if (jmp_charts_dir / f"{param}.png").exists()
    ]
    known = {p.name for _, p in charts}
    charts += [(p.stem, p) for p in sorted(jmp_charts_dir.glob("*.png")) if p.name not in known]

    if not charts:
        return ['<div class="alert alert-info m-4">No JMP chart images found in output folder.</div>']

    names = [_html.escape(n) for n, _ in charts]
    options = "".join(f'<option value="{n}">{n}</option>' for n in names)

    parts: list[str] = [
        '<div class="d-flex align-items-center gap-3 mb-3 px-1 flex-wrap">',
        f'<span class="badge rounded-pill text-bg-primary fs-6 px-3 py-2">'
        f'&#x1F4CA; {len(charts)} JMP Charts</span>',
        '<span class="text-muted small">Generated by JMP Graph Builder &mdash; '
        'X: Rank (Params) &nbsp;|&nbsp; Group: Frequency (or Boot/Run temperature for DTR) '
        f'&nbsp;|&nbsp; Overlay: Gear / Source &nbsp;|&nbsp; Ref lines: {ref_note}</span>',
        '</div>',
        '<div class="section-card">',
        '<div class="d-flex align-items-center flex-wrap gap-2 mb-2">',
        '<h5 class="fw-bold mb-0 me-auto">&#x1F50D; Per-Parameter Margin Charts &mdash; Comparison</h5>',
        '<label class="small fw-semibold" for="jmpCols">Panels per row</label>',
        '<select id="jmpCols" class="form-select form-select-sm w-auto" onchange="jmpSetCols(this.value)">'
        '<option value="1">1</option><option value="2" selected>2</option>'
        '<option value="3">3</option></select>',
        '<button class="btn btn-sm btn-outline-primary" onclick="jmpAddPanel()">&#x2795; Add panel</button>',
        '<button class="btn btn-sm btn-outline-secondary" onclick="jmpReset()">&#x21BA; Reset to default</button>',
        '</div>',
        '<p class="text-muted small mb-3">Pick any chart in each panel&rsquo;s dropdown to compare '
        'parameters side-by-side. Each chart shows positive (+) and negative (&minus;) '
        'training margins per DIMM rank; reference lines mark the minimum acceptable margin.</p>',
        '<div class="row g-3" id="jmpGrid">',
    ]
    for idx, ((name, png_path), esc) in enumerate(zip(charts, names)):
        sel_opts = options.replace(f'value="{esc}"', f'value="{esc}" selected', 1)
        parts += [
            f'<div class="col-xl-6 jmp-slot" data-default="{esc}">',
            '<div class="border rounded p-2 bg-light h-100">',
            '<div class="d-flex align-items-center gap-2 mb-1">',
            f'<select class="form-select form-select-sm jmp-sel" aria-label="Chart for panel {idx + 1}" '
            f'onchange="jmpShow(this)">{sel_opts}</select>',
            '<button class="btn btn-sm btn-outline-danger" title="Remove panel" '
            'onclick="jmpRemove(this)">&times;</button>',
            '</div>',
            f'<img src="{_b64src(png_path)}" data-chart="{esc}" class="img-fluid rounded jmp-img" '
            f'style="width:100%;border:1px solid #dee2e6;" alt="JMP Chart: {esc}">',
            '</div></div>',
        ]
    parts += ['</div>', '</div>']
    return parts


def _mr_odt_snapshots(platform_infos: list[dict] | None) -> list[tuple[str, dict, dict]]:
    """Flatten ``platform_infos[*]["mr_odt"]["snapshots"]`` into
    ``(label, file_info, snapshot)`` tuples with integer MR keys."""
    out: list[tuple[str, dict, dict]] = []
    for pi in platform_infos or []:
        mo = pi.get("mr_odt") or {}
        snaps = mo.get("snapshots") or []
        for s in snaps:
            s = dict(s)
            for key in ("final_mrs", "init_mrs"):
                s[key] = {
                    rank: {int(mr): int(v) for mr, v in (mrs or {}).items()}
                    for rank, mrs in (s.get(key) or {}).items()
                }
            label = str(pi.get("source_file") or mo.get("source_file") or "?")
            if len(snaps) > 1:
                fg = ""
                if s.get("frequency"):
                    fg = f" {s['frequency']}" + (f" G{s['gear']}" if s.get("gear") else "")
                label += f" \u2014 training pass {s.get('index')}{fg}"
            out.append((label, mo, s))
    return out


def _rank_sort_key(rank: str) -> list[int]:
    return [int(x) for x in re.findall(r"\d+", rank)]


def _snap_selector(prefix: str, labels: list[str]) -> str:
    opts = "".join(
        f'<option value="{i}">{_html.escape(lbl)}</option>' for i, lbl in enumerate(labels))
    return (
        f'<label class="small fw-semibold" for="{prefix}Sel">Log:</label>'
        f'<select id="{prefix}Sel" class="form-select form-select-sm w-auto" style="max-width:520px" '
        f'onchange="showSnap(\'{prefix}\', this.value)">{opts}</select>'
    )


def _build_mr_html_section(platform_infos: list[dict] | None) -> list[str]:
    """HTML for the Mode Registers tab (final per-rank MR values)."""
    snaps = _mr_odt_snapshots(platform_infos)
    if not snaps:
        return ['<div class="alert alert-info m-4">No Mode Register data found. The logs '
                'need the MRC "SAGV Finalization" per-rank MR table or the "InitMrwDdr5" '
                'JEDEC-reset writes.</div>']

    parts: list[str] = [
        '<div class="section-card">',
        '<h5 class="fw-bold mb-1">&#x1F9EE; DDR5 Mode Registers &mdash; final trained values per rank</h5>',
        '<p class="text-muted small mb-2">Final values come from the per-rank MR table MRC prints '
        'during <b>SAGV Finalization</b>; JEDEC-reset values come from <b>InitMrwDdr5</b>. '
        '<span class="mr-chg px-1">Amber</span> cells were changed by training (hover a cell for the '
        'reset value and decode); <b>bold</b> rows differ between ranks. Decodes follow JEDEC '
        'JESD79-5 &mdash; MR5 / MR34 / MR35 / MR37&ndash;39 are cross-checked against MRC&rsquo;s own '
        'DIMM ODT summary and ODT timing prints.</p>',
        '<div class="d-flex align-items-center flex-wrap gap-3 mb-3">',
        _snap_selector("mr", [s[0] for s in snaps]),
        '<div class="form-check form-switch mb-0"><input class="form-check-input" type="checkbox" '
        'id="mrChgOnly" onchange="filterMr()"><label class="form-check-label small" for="mrChgOnly">'
        'Only MRs changed by training</label></div>',
        '<div class="form-check form-switch mb-0"><input class="form-check-input" type="checkbox" '
        'id="mrVarOnly" onchange="filterMr()"><label class="form-check-label small" for="mrVarOnly">'
        'Only MRs that differ between ranks</label></div>',
        '<div class="form-check form-switch mb-0"><input class="form-check-input" type="checkbox" '
        'id="mrDecOnly" onchange="filterMr()"><label class="form-check-label small" for="mrDecOnly">'
        'Only decoded MRs (CL / Ron / Vref / ODT)</label></div>',
        '</div>',
    ]

    for k, (label, _mo, s) in enumerate(snaps):
        final, init = s["final_mrs"], s["init_mrs"]
        use_final = bool(final)
        table = final if use_final else init
        ranks = sorted(table, key=_rank_sort_key)
        mrs = sorted({mr for r in ranks for mr in table[r]})
        parts.append(f'<div class="mr-snap" id="mr-snap-{k}" style="display:{"block" if k == 0 else "none"}">')
        if not use_final:
            parts.append('<div class="alert alert-warning py-2 small">This log has no final '
                         '(SAGV Finalization) MR table &mdash; showing JEDEC-reset '
                         '(InitMrwDdr5) values.</div>')
        parts += [
            '<div class="table-responsive" style="max-height:620px">',
            '<table class="table table-sm table-bordered align-middle mr-table mb-0" style="font-size:.8rem">',
            '<thead class="table-dark" style="position:sticky;top:0"><tr><th>MR</th><th>Function</th>'
            '<th>Decoded</th>',
            *[f'<th class="text-center">{_html.escape(r)}</th>' for r in ranks],
            '</tr></thead><tbody>',
        ]
        for mr in mrs:
            vals = [table[r].get(mr) for r in ranks]
            distinct = {v for v in vals if v is not None}
            varies = len(distinct) > 1
            decoded_any = False
            changed_any = False
            cells: list[str] = []
            for r, v in zip(ranks, vals):
                if v is None:
                    cells.append('<td class="text-center text-muted">&mdash;</td>')
                    continue
                dec = decode_ddr5_mr(mr, v)
                decoded_any = decoded_any or bool(dec)
                iv = init.get(r, {}).get(mr) if use_final else None
                changed = use_final and iv is not None and iv != v
                changed_any = changed_any or changed
                tip = f"{r} MR{mr} = 0x{v:02X}"
                if iv is not None and use_final:
                    tip += f" | JEDEC reset 0x{iv:02X}"
                if dec:
                    tip += f" | {dec}"
                cls = ' class="text-center mr-chg"' if changed else ' class="text-center"'
                cells.append(f'<td{cls} title="{_html.escape(tip)}"><code>0x{v:02X}</code></td>')
            if len(distinct) == 1:
                decoded = decode_ddr5_mr(mr, next(iter(distinct)))
            elif decoded_any:
                decoded = "varies by rank (hover cells)"
            else:
                decoded = ""
            row_style = ' style="font-weight:700"' if varies else ""
            parts.append(
                f'<tr data-chg="{int(changed_any)}" data-var="{int(varies)}" data-dec="{int(decoded_any)}"{row_style}>'
                f'<td class="fw-semibold">MR{mr}</td><td class="small">{_html.escape(_mr_description(mr))}</td>'
                f'<td class="small" style="min-width:180px">{_html.escape(decoded)}</td>'
                + "".join(cells)
                + '</tr>'
            )
        parts += ['</tbody></table></div>', '</div>']
    parts.append('</div>')

    # ── Cross-log comparison: MRs whose final value differs between logs ──
    if len(snaps) > 1:
        per_snap: list[dict[int, str]] = []
        for _label, _mo, s in snaps:
            table = s["final_mrs"] or s["init_mrs"]
            vals: dict[int, set[int]] = defaultdict(set)
            for r in table:
                for mr, v in table[r].items():
                    vals[mr].add(v)
            per_snap.append({mr: "/".join(f"0x{v:02X}" for v in sorted(vs)) for mr, vs in vals.items()})
        all_mrs = sorted({mr for d in per_snap for mr in d})
        diff_mrs = [mr for mr in all_mrs if len({d.get(mr) for d in per_snap}) > 1]
        parts += [
            '<div class="section-card">',
            '<h5 class="fw-bold mb-1">&#x1F50E; Cross-Log MR Differences</h5>',
            '<p class="text-muted small mb-3">MRs whose trained value differs between the loaded logs '
            '(values across ranks shown as a / separated set). Useful to spot training decisions that '
            'change with frequency, gear or temperature.</p>',
        ]
        if not diff_mrs:
            parts.append(f'<div class="alert alert-success py-2 mb-0">All {len(snaps)} logs have '
                         'identical Mode Register values.</div>')
        else:
            parts += [
                '<div class="table-responsive" style="max-height:520px">',
                '<table class="table table-sm table-bordered table-hover mb-0" style="font-size:.78rem;white-space:nowrap">',
                '<thead class="table-dark" style="position:sticky;top:0"><tr><th>Log</th>',
                *[f'<th>MR{mr}<div class="small fw-normal">{_html.escape(_mr_description(mr))}</div></th>'
                  for mr in diff_mrs],
                '</tr></thead><tbody>',
            ]
            for (label, _mo, _s), d in zip(snaps, per_snap):
                parts.append(f'<tr><td class="fw-semibold">{_html.escape(label)}</td>'
                             + "".join(f'<td><code>{d.get(mr, "&mdash;")}</code></td>' for mr in diff_mrs)
                             + '</tr>')
            parts += ['</tbody></table></div>']
        parts.append('</div>')
    return parts


def _rank_odt_from_mrs(mrs: dict[int, int]) -> dict[str, str]:
    """Per-rank DRAM ODT / drive-strength values decoded from final MRs."""
    out: dict[str, str] = {}
    if 34 in mrs:
        out["RTT_WR"] = _ohm(_RTT_OHMS, (mrs[34] >> 3) & 0x7)
        out["RTT_PARK"] = _ohm(_RTT_OHMS, mrs[34] & 0x7)
    if 35 in mrs:
        out["RTT_NOM_WR"] = _ohm(_RTT_OHMS, mrs[35] & 0x7)
        out["RTT_NOM_RD"] = _ohm(_RTT_OHMS, (mrs[35] >> 3) & 0x7)
    if 36 in mrs:
        out["RTT_LOOPBACK"] = _ohm(_RTT_OHMS, mrs[36] & 0x7)
    if 33 in mrs:
        out["DQS_RTT_PARK"] = _ohm(_RTT_OHMS, (mrs[33] >> 3) & 0x7)
        out["CA_ODT"] = _ohm(_CA_ODT_OHMS, mrs[33] & 0x7)
    if 32 in mrs:
        out["CK_ODT"] = _ohm(_CA_ODT_OHMS, mrs[32] & 0x7)
        out["CS_ODT"] = _ohm(_CA_ODT_OHMS, (mrs[32] >> 3) & 0x7)
    if 5 in mrs:
        out["Ron PU"] = _ohm(_RON_OHMS, (mrs[5] >> 1) & 0x3)
        out["Ron PD"] = _ohm(_RON_OHMS, (mrs[5] >> 6) & 0x3)
    return out


def _simple_table(headers: list[str], rows: list[list[Any]], small: bool = True) -> str:
    th = "".join(f"<th>{_html.escape(str(h))}</th>" for h in headers)
    first = ' class="fw-semibold"'
    body = "".join(
        "<tr>" + "".join(
            f'<td{first if j == 0 else ""}>{_html.escape(str(c))}</td>'
            for j, c in enumerate(r)) + "</tr>"
        for r in rows
    )
    fs = "font-size:.8rem;" if small else ""
    return ('<div class="table-responsive"><table class="table table-sm table-bordered table-hover '
            f'align-middle mb-3" style="{fs}white-space:nowrap"><thead class="table-dark"><tr>{th}</tr>'
            f'</thead><tbody>{body}</tbody></table></div>')


def _build_odt_html_section(platform_infos: list[dict] | None) -> list[str]:
    """HTML for the ODT tab: DIMM ODT summary, CPU read ODT, per-rank DRAM
    ODT decoded from final MRs, ODT latency timing and BIOS ODT inputs."""
    snaps = _mr_odt_snapshots(platform_infos)
    if not snaps:
        return ['<div class="alert alert-info m-4">No ODT data found. The logs need the MRC '
                '"MRC Data Summary" DIMM ODT summary, "DDR5 ODT Timing Config" or a final MR table.</div>']

    parts: list[str] = [
        '<div class="section-card">',
        '<h5 class="fw-bold mb-1">&#x1F50C; On-Die Termination &amp; Drive Strength</h5>',
        '<p class="text-muted small mb-2">All ODT values MRC reports for the selected log: DIMM ODT '
        'summary (RTT_WR, RTT_NOM, RTT_PARK, CA/CS groups, Ron) and CPU read ODT from '
        '<b>MRC Data Summary</b>, per-rank DRAM ODT decoded from the final MRs, ODT latency offsets '
        'from <b>DDR5 ODT Timing Config</b>, and the BIOS ODT input overrides. Values in &Omega;.</p>',
        '<div class="d-flex align-items-center flex-wrap gap-3 mb-3">',
        _snap_selector("odt", [s[0] for s in snaps]),
        '</div>',
    ]
    for k, (label, mo, s) in enumerate(snaps):
        parts.append(f'<div class="odt-snap" id="odt-snap-{k}" style="display:{"block" if k == 0 else "none"}">')
        any_data = False
        if s.get("odt_summary"):
            any_data = True
            fields = s.get("odt_fields") or sorted({f for d in s["odt_summary"].values() for f in d})
            rows = [[dimm] + [d.get(f, "") for f in fields]
                    for dimm, d in sorted(s["odt_summary"].items(), key=lambda kv: _rank_sort_key(kv[0]))]
            parts += ['<h6 class="fw-bold text-primary">DIMM ODT Summary (&Omega;)</h6>',
                      _simple_table(["DIMM"] + fields, rows)]
        if s.get("cpu_read_odt"):
            any_data = True
            rows = [[ch, v] for ch, v in sorted(s["cpu_read_odt"].items(), key=lambda kv: _rank_sort_key(kv[0]))]
            parts += ['<h6 class="fw-bold text-primary">CPU Read ODT (&Omega;)</h6>',
                      _simple_table(["MC.Channel", "Read ODT"], rows)]
        final = s.get("final_mrs") or {}
        rank_odt = {r: _rank_odt_from_mrs(m) for r, m in final.items()}
        rank_odt = {r: d for r, d in rank_odt.items() if d}
        if rank_odt:
            any_data = True
            cols = list(dict.fromkeys(c for d in rank_odt.values() for c in d))
            rows = [[r] + [rank_odt[r].get(c, "") for c in cols]
                    for r in sorted(rank_odt, key=_rank_sort_key)]
            parts += ['<h6 class="fw-bold text-primary">Per-Rank DRAM ODT / Ron (decoded from final MRs)</h6>',
                      _simple_table(["Rank"] + cols, rows)]
        if s.get("odtl"):
            any_data = True
            cols = list(dict.fromkeys(c for d in s["odtl"].values() for c in d))
            rows = [[r] + [s["odtl"][r].get(c, "") for c in cols]
                    for r in sorted(s["odtl"], key=_rank_sort_key)]
            parts += ['<h6 class="fw-bold text-primary">ODT Latency Offsets (tCK, DDR5 ODT Timing Config)</h6>',
                      _simple_table(["Rank"] + cols, rows)]
        inputs = mo.get("odt_inputs") or {}
        facts = []
        if mo.get("rcomp_rd_odt") is not None:
            facts.append(("CPU RcompTarget[RdOdt]", f"{mo['rcomp_rd_odt']} \u03a9"))
        if mo.get("ddrio_odt_mode"):
            facts.append(("DDRIO ODT Mode", mo["ddrio_odt_mode"]))
        if inputs or facts:
            any_data = True
            parts.append('<h6 class="fw-bold text-primary">BIOS / CPU ODT Inputs</h6>')
            if facts:
                parts.append('<div class="row g-2 mb-2">' + "".join(
                    f'<div class="col-md-4 col-lg-3"><div class="border rounded p-2 bg-light">'
                    f'<div class="small text-muted">{_html.escape(a)}</div>'
                    f'<div class="fw-semibold">{_html.escape(str(b))}</div></div></div>'
                    for a, b in facts) + '</div>')
            if inputs:
                fields = list(dict.fromkeys(f for d in inputs.values() for f in d))
                rows = [[dimm] + [inputs[dimm].get(f, "") for f in fields] for dimm in sorted(inputs)]
                parts += ['<div class="small text-muted mb-1">BIOS ODT overrides '
                          '(&ldquo;Disabled&rdquo; = MRC auto / trained value is used).</div>',
                          _simple_table(["DIMM"] + fields, rows)]
        if not any_data:
            parts.append('<div class="alert alert-warning py-2">No ODT prints found for this log.</div>')
        parts.append('</div>')
    parts.append('</div>')

    # ── Cross-log ODT matrix ─────────────────────────────────────────────
    with_sum = [(lbl, s) for lbl, _mo, s in snaps if s.get("odt_summary") or s.get("cpu_read_odt")]
    if len(with_sum) > 1:
        fields = list(dict.fromkeys(
            f for _l, s in with_sum for f in (s.get("odt_fields") or [])))
        matrix: list[list[str]] = []
        for lbl, s in with_sum:
            row = [lbl]
            for f in fields:
                vs = sorted({str(d.get(f)) for d in s.get("odt_summary", {}).values() if d.get(f) is not None},
                            key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else 0, x))
                row.append("/".join(vs) or "\u2014")
            cpu = sorted({str(v) for v in s.get("cpu_read_odt", {}).values()})
            row.append("/".join(cpu) or "\u2014")
            matrix.append(row)
        headers = ["Log"] + fields + ["CPU RdODT"]
        varies = [len({r[i] for r in matrix}) > 1 for i in range(len(headers))]
        var_cls = ' class="odt-var"'
        th = "".join(
            f'<th{var_cls if varies[i] and i else ""}>{_html.escape(h)}</th>'
            for i, h in enumerate(headers))
        body = "".join(
            "<tr>" + "".join(
                f'<td class="{"fw-semibold" if i == 0 else ("odt-var-cell" if varies[i] else "")}">'
                f'{_html.escape(c)}</td>' for i, c in enumerate(r)) + "</tr>"
            for r in matrix)
        parts += [
            '<div class="section-card">',
            '<h5 class="fw-bold mb-1">&#x1F50E; Cross-Log ODT Comparison</h5>',
            '<p class="text-muted small mb-3">Trained ODT per log (values across DIMMs shown as a / '
            'separated set). <span class="odt-var-cell px-1">Highlighted</span> columns differ between logs.</p>',
            '<div class="table-responsive"><table class="table table-sm table-bordered table-hover mb-0" '
            f'style="font-size:.8rem;white-space:nowrap"><thead class="table-dark"><tr>{th}</tr></thead>'
            f'<tbody>{body}</tbody></table></div>',
            '</div>',
        ]
    return parts


def generate_html_report(
    path: Path,
    rows: list[dict[str, Any]],
    chart_params: list[str],
    axis_config: dict | None = None,
    platform_infos: list[dict] | None = None,
    training_steps: list[dict] | None = None,
    jmp_charts_dir: Path | None = None,
    project_name: str | None = None,
) -> None:
    """Generate a self-contained interactive HTML report (PowerBI-style)."""
    import json as _json
    import math as _math
    from datetime import datetime as _dt

    project_name = project_name or _ACTIVE_PROJECT_NAME

    def _sf(v):
        """Return a JSON-safe float, or None for NaN/Inf/None."""
        if v is None:
            return None
        try:
            f = float(v)
            return None if (_math.isnan(f) or _math.isinf(f)) else round(f, 3)
        except (TypeError, ValueError):
            return None

    # ── Compute statistics ────────────────────────────────────────────
    param_stats = _compute_param_stats(rows, chart_params)
    freq_stats  = _compute_freq_stats(rows, chart_params)
    freqs       = sorted(freq_stats.keys(), key=lambda x: (x is None, x))
    thresholds  = {p: _get_threshold(p, axis_config) for p in chart_params}
    files       = sorted({str(r.get("SourceFile", "")) for r in rows})
    generated   = _dt.now().strftime("%Y-%m-%d %H:%M:%S")
    freq_labels = [str(f) if f is not None else "Unknown" for f in freqs]
    # Per-parameter +Ref / -Ref reference-line thresholds for raw-data
    # highlighting (cells whose margin is weaker than its ref line).
    thresholds_full = {
        p: {"plus": _get_threshold(p, axis_config),
            "minus": _get_threshold_minus(p, axis_config)}
        for p in chart_params
    }
    _ref_pairs = {(thresholds_full[p]["plus"], thresholds_full[p]["minus"]) for p in chart_params}
    if len(_ref_pairs) == 1:
        _rp, _rm = next(iter(_ref_pairs))
        ref_note = f"+{_fmt_num(_rp)} / {_fmt_num(_rm)}"
    else:
        ref_note = "per-parameter &plusmn;Ref (Tab 2 &middot; Parameters &amp; Axis)"

    n_pass = sum(1 for p in chart_params
                 if _margin_status(
                     param_stats.get(p, {}).get("avg_width", 0),
                     thresholds[p]) == "PASS")
    n_warn = sum(1 for p in chart_params
                 if _margin_status(
                     param_stats.get(p, {}).get("avg_width", 0),
                     thresholds[p]) == "WARN")
    n_fail = len(chart_params) - n_pass - n_warn

    # ── Embedded JS data ──────────────────────────────────────────────
    def _cpk(p: str) -> float | None:
        """One-sided lower-spec process-capability index.

        Cpk = (mean_width - LSL) / (3*sigma), where LSL is the margin
        threshold. This is the standard semiconductor margin-qualification
        metric: Cpk >= 1.33 is the industry-accepted "capable" bar, 1.0-1.33
        is marginal, < 1.0 is a fail risk. Returns None for zero-variance
        (constant) data, which is shown as the unbounded case.
        """
        st = param_stats.get(p, {})
        mu = st.get("avg_width")
        sd = st.get("std_width")
        if mu is None:
            return None
        if not sd or sd <= 0:
            return None
        return (mu - thresholds[p]) / (3.0 * sd)

    def _guardband(p: str) -> float | None:
        """Worst-case window minus threshold; negative => spec violation."""
        mn = param_stats.get(p, {}).get("min_width")
        return (mn - thresholds[p]) if mn is not None else None

    summary_data = [
        {
            "param":     p,
            "avg_plus":  _sf(param_stats.get(p, {}).get("avg_plus")),
            "avg_minus": _sf(param_stats.get(p, {}).get("avg_minus")),
            "avg_width": _sf(param_stats.get(p, {}).get("avg_width")),
            "min_width": _sf(param_stats.get(p, {}).get("min_width")),
            "max_width": _sf(param_stats.get(p, {}).get("max_width")),
            "median":    _sf(param_stats.get(p, {}).get("median")),
            "std_width": _sf(param_stats.get(p, {}).get("std_width")),
            "n":         int(param_stats.get(p, {}).get("count", 0)),
            "cpk":       _sf(_cpk(p)),
            # Worst-case guardband: how far the weakest measured window sits
            # above (or below) the spec threshold. Negative => spec violation.
            "guardband": _sf(_guardband(p)),
            "threshold": _sf(thresholds[p]),
            "status":    _margin_status(
                param_stats.get(p, {}).get("avg_width", 0), thresholds[p]),
        }
        for p in chart_params
    ]

    freq_series = {
        p: [_sf(freq_stats.get(fr, {}).get(p, {}).get("avg_width")) for fr in freqs]
        for p in chart_params
    }
    freq_plus_series = {
        p: [_sf(freq_stats.get(fr, {}).get(p, {}).get("avg_plus")) for fr in freqs]
        for p in chart_params
    }
    freq_minus_series = {
        p: [_sf(freq_stats.get(fr, {}).get(p, {}).get("avg_minus")) for fr in freqs]
        for p in chart_params
    }

    # RunTemp drift
    run_temps_all: list[float] = []
    for r in rows:
        rt = r.get("RunTemp")
        if rt is not None:
            try:
                run_temps_all.append(float(rt))
            except (TypeError, ValueError):
                pass
    run_temps = sorted(set(run_temps_all))

    drift_series: dict = {}
    if len(run_temps) > 1:
        for p in chart_params:
            drift_series[p] = []
            for rt in run_temps:
                wl = [calc_window_width(r, p) for r in rows
                      if r.get("RunTemp") is not None
                      and abs(float(r["RunTemp"]) - rt) < 0.5
                      and calc_window_width(r, p) is not None]
                drift_series[p].append(_sf(statistics.mean(wl)) if wl else None)

    # Per-rank averaged data for scatter/bar by rank
    by_rank: dict = defaultdict(list)
    for r in rows:
        by_rank[str(r.get("Params", ""))].append(r)

    rank_data: dict = {}
    for p in chart_params:
        ranks_sorted = sorted(by_rank.keys())
        plus_avg, minus_avg = [], []
        plus_worst, minus_worst = [], []
        for rank in ranks_sorted:
            rr = by_rank[rank]
            pv = [float(r[f"{p}+"]) for r in rr if r.get(f"{p}+") is not None]
            mv = [float(r[f"{p}-"]) for r in rr if r.get(f"{p}-") is not None]
            plus_avg.append(_sf(statistics.mean(pv)) if pv else None)
            minus_avg.append(_sf(statistics.mean(mv)) if mv else None)
            # Worst-case (qualification-relevant) margin per rank:
            #   weakest positive  = smallest +margin (min of pv)
            #   weakest negative  = -margin closest to 0 (max of mv, since mv<0)
            plus_worst.append(_sf(min(pv)) if pv else None)
            minus_worst.append(_sf(max(mv)) if mv else None)
        rank_data[p] = {
            "ranks": ranks_sorted,
            "plus": plus_avg, "minus": minus_avg,
            "plusWorst": plus_worst, "minusWorst": minus_worst,
        }

    # Raw table rows (all fields)
    raw_table_rows = []
    _is_dtr = any(r.get("RunTemp") not in (None, "") for r in rows)
    _file_blocks: dict[str, set] = defaultdict(set)
    for r in rows:
        _file_blocks[str(r.get("SourceFile", ""))].add(str(r.get("BlockIndex", "")))

    def _blk(r) -> int | None:
        try:
            return int(float(r.get("BlockIndex")))
        except (TypeError, ValueError):
            return None

    def _phase(r) -> str:
        b = _blk(r)
        if _is_dtr:
            return {1: "Boot RMT", 2: "Run RMT"}.get(b, f"Blk {b}" if b is not None else "RMT")
        if len(_file_blocks[str(r.get("SourceFile", ""))]) > 1 and b is not None:
            return f"Blk {b}"
        return "RMT"

    # Short, unique per-file labels (e.g. "BCRH 3200 G2") — full name on hover.
    _file_label: dict[str, str] = {}
    _file_group: dict[str, str] = {}
    for r in rows:
        f = str(r.get("SourceFile", ""))
        if f in _file_label:
            continue
        stem = Path(f).stem
        grp = re.split(r"[_\-\s]+", stem)[0] or stem
        fq, gr = r.get("Frequency"), r.get("Gear")
        if fq not in (None, "") and str(fq) in stem:
            lbl = f"{grp} {fq}" + (f" G{gr}" if gr not in (None, "") else "")
        else:
            lbl = stem if len(stem) <= 24 else stem[:11] + "\u2026" + stem[-11:]
        _file_label[f] = lbl
        _file_group[f] = grp
    _seen_lbl: dict[str, int] = {}
    for f in list(_file_label):
        lbl = _file_label[f]
        if list(_file_label.values()).count(lbl) > 1:
            _seen_lbl[lbl] = _seen_lbl.get(lbl, 0) + 1
            _file_label[f] = f"{lbl} ({_seen_lbl[lbl]})"

    for r in rows:
        f = str(r.get("SourceFile", ""))
        rd: dict = {
            "file": f,
            "lbl":  _file_label.get(f, f),
            "grp":  _file_group.get(f, f),
            "phase": _phase(r),
            "freq": str(r.get("Frequency", "") if r.get("Frequency") is not None else ""),
            "gear": str(r.get("Gear", "") if r.get("Gear") is not None else ""),
            "rank": str(r.get("Params", "")),
            "bt":   _sf(r.get("BootTemp")),
            "rt":   _sf(r.get("RunTemp")),
        }
        for p in chart_params:
            rd[p + "+"] = _sf(r.get(p + "+"))
            rd[p + "-"] = _sf(r.get(p + "-"))
        raw_table_rows.append(rd)

    # Parameters whose margins never change (e.g. RecEnDelay ±32) are hidden
    # by default in the Raw Data tab — they add columns but no information.
    raw_constant_params = [
        p for p in chart_params
        if len({(rd.get(p + "+"), rd.get(p + "-")) for rd in raw_table_rows}) <= 1
    ]

    # ── Margin degradation summary (values weaker than configured ±Ref) ──
    # For each parameter, count samples whose + margin falls below +Ref, or
    # whose − margin rises above −Ref (i.e. closer to 0 than the spec limit).
    degradation_rows = []
    for p in chart_params:
        thr = thresholds_full[p]
        plus_thr, minus_thr = thr["plus"], thr["minus"]
        plus_vals  = [float(r[f"{p}+"]) for r in rows if r.get(f"{p}+") is not None]
        minus_vals = [float(r[f"{p}-"]) for r in rows if r.get(f"{p}-") is not None]
        plus_fail  = [v for v in plus_vals  if v < plus_thr]
        minus_fail = [v for v in minus_vals if v > minus_thr]
        total = len(plus_vals) + len(minus_vals)
        n_fail = len(plus_fail) + len(minus_fail)
        degradation_rows.append({
            "param":       p,
            "plus_fail":   len(plus_fail),
            "minus_fail":  len(minus_fail),
            "total_fail":  n_fail,
            "total":       total,
            "pct":         (n_fail / total * 100.0) if total else 0.0,
            "worst_plus":  _sf(min(plus_vals)) if plus_vals else None,
            "worst_minus": _sf(max(minus_vals)) if minus_vals else None,
            "plus_thr":    _sf(plus_thr),
            "minus_thr":   _sf(minus_thr),
        })
    degradation_rows.sort(key=lambda d: d["total_fail"], reverse=True)
    total_degraded = sum(d["total_fail"] for d in degradation_rows)

    embedded_data = _json.dumps({
        "params":          chart_params,
        "freqLabels":      freq_labels,
        "summary":         summary_data,
        "freqSeries":      freq_series,
        "freqPlusSeries":  freq_plus_series,
        "freqMinusSeries": freq_minus_series,
        "drift":           drift_series,
        "runTemps":        run_temps,
        "rankData":        rank_data,
        "thresholds":      thresholds_full,
    }, separators=(",", ":"))
    raw_data_json = _json.dumps(raw_table_rows, separators=(",", ":"))
    raw_meta_json = _json.dumps({
        "dtr": _is_dtr,
        "constant": raw_constant_params,
    }, separators=(",", ":"))

    # ── Dynamic HTML snippets ─────────────────────────────────────────
    stats_table_rows = ""
    for s in summary_data:
        badge = {"PASS": "badge bg-success", "WARN": "badge bg-warning text-dark",
                 "FAIL": "badge bg-danger"}.get(s["status"], "badge bg-secondary")
        def _v(x):
            return str(x) if x is not None else "—"
        # Cpk colouring against the 1.33 / 1.00 industry capability bars.
        cpk = s["cpk"]
        if cpk is None:
            cpk_txt, cpk_cls = ("&infin;" if s["std_width"] in (0, None) else "—"), "text-success"
        else:
            cpk_txt = f"{cpk:.2f}"
            cpk_cls = ("text-success fw-semibold" if cpk >= 1.33
                       else "text-warning fw-semibold" if cpk >= 1.0
                       else "text-danger fw-semibold")
        # Guardband: negative = worst-case window below threshold.
        gb = s["guardband"]
        gb_cls = "text-danger fw-semibold" if (gb is not None and gb < 0) else "text-muted"
        stats_table_rows += (
            f'<tr>'
            f'<td class="fw-semibold">{s["param"]}</td>'
            f'<td class="text-success fw-semibold">{_v(s["avg_plus"])}</td>'
            f'<td class="text-danger fw-semibold">{_v(s["avg_minus"])}</td>'
            f'<td class="fw-bold">{_v(s["avg_width"])}</td>'
            f'<td>{_v(s["min_width"])}</td>'
            f'<td>{_v(s["max_width"])}</td>'
            f'<td class="text-muted">{_v(s["median"])}</td>'
            f'<td class="text-muted">{_v(s["std_width"])}</td>'
            f'<td class="{cpk_cls}">{cpk_txt}</td>'
            f'<td class="{gb_cls}">{_v(gb)}</td>'
            f'<td class="text-muted">{_v(s["threshold"])}</td>'
            f'<td class="text-muted">{s["n"]}</td>'
            f'<td><span class="{badge}">{s["status"]}</span></td>'
            f'</tr>\n'
        )

    # ── Margin degradation summary HTML (Overview tab) ────────────────
    # Only rendered when at least one sample falls outside the configured
    # ±Ref limits; if everything is within spec, the section is omitted
    # entirely (no green "all clear" banner is shown).
    def _dv(x):
        return str(x) if x is not None else "—"
    if total_degraded == 0:
        degradation_html = ""
    else:
        _deg_rows_html = ""
        for d in degradation_rows:
            if d["total_fail"] == 0:
                continue
            pct_cls = ("text-danger fw-bold" if d["pct"] >= 5
                       else "text-warning fw-semibold" if d["pct"] > 0 else "text-muted")
            _deg_rows_html += (
                f'<tr>'
                f'<td class="fw-semibold">{d["param"]}</td>'
                f'<td class="text-danger">{d["plus_fail"]}</td>'
                f'<td class="text-danger">{d["minus_fail"]}</td>'
                f'<td class="fw-bold text-danger">{d["total_fail"]} / {d["total"]}</td>'
                f'<td class="{pct_cls}">{d["pct"]:.1f}%</td>'
                f'<td class="text-muted">{_dv(d["worst_plus"])} (Ref {_dv(d["plus_thr"])})</td>'
                f'<td class="text-muted">{_dv(d["worst_minus"])} (Ref {_dv(d["minus_thr"])})</td>'
                f'</tr>\n'
            )
        degradation_html = (
            '<div class="alert alert-danger py-2 mb-3">'
            f'&#9888; <b>{total_degraded}</b> sample(s) across '
            f'<b>{sum(1 for d in degradation_rows if d["total_fail"])}</b> parameter(s) '
            'fall below the configured &plusmn;Ref margin limits.</div>'
            '<div class="table-responsive">'
            '<table class="table table-sm table-hover align-middle mb-0">'
            '<thead class="table-dark"><tr>'
            '<th>Parameter</th><th>Below +Ref</th><th>Above &minus;Ref</th>'
            '<th>Total Violations / N</th><th>% of Samples</th>'
            '<th>Worst +</th><th>Worst &minus;</th>'
            '</tr></thead><tbody>' + _deg_rows_html + '</tbody></table></div>'
        )

    param_opts = "".join(
        f'<option value="{p}">{p}</option>' for p in chart_params
    )
    param_btns = "".join(
        '<button class="btn btn-{cls} btn-sm param-btn me-1 mb-1"'
        ' onclick="showRankChart(\'{p}\')">{p}</button>'.format(
            cls="primary" if i == 0 else "outline-primary", p=p
        )
        for i, p in enumerate(chart_params)
    )
    freq_sel_opts = "".join(
        '<option value="{fl}">{fl} MT/s</option>'.format(fl=fl)
        for fl in freq_labels
    )
    drift_tab_body = (
        '<div class="section-card"><div class="chart-lg">'
        '<canvas id="driftChart"></canvas></div></div>'
        if len(run_temps) > 1 else
        '<div class="alert alert-info m-4">Only one run temperature found'
        ' — drift analysis not available.</div>'
    )
    drift_init_js = ""
    if len(run_temps) > 1:
        drift_init_js = (
            "new Chart(document.getElementById('driftChart'),{"
            "type:'line',"
            "data:{labels:D.runTemps.map(t=>t+'°C'),"
            "datasets:D.params.map((p,i)=>({label:p,"
            "data:D.drift[p]||[],"
            "borderColor:PAL[i%PAL.length],"
            "backgroundColor:PAL[i%PAL.length]+'22',"
            "tension:.3,fill:false,pointRadius:5"
            "}))},"
            "options:{responsive:true,maintainAspectRatio:false,"
            "plugins:{legend:{position:'right'}},"
            "scales:{y:{title:{display:true,text:'Avg Window Width'},"
            "grid:{color:'#e5e7eb'}},"
            "x:{title:{display:true,text:'Run Temperature (°C)'},"
            "grid:{display:false}}}}});"
        )

    # ── Full HTML ─────────────────────────────────────────────────────
    html_parts = [
        "<!DOCTYPE html>",
        '<html lang="en">',
        "<head>",
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        f"<title>{TOOL_NAME} \u2014 RMT Margin Analysis Report"
        + (f" \u2014 {_html.escape(project_name)}" if project_name else "") + "</title>",
        '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css">',
        "<style>",
        "body{background:#f0f4f8;font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;}",
        ".rmt-hdr{background:linear-gradient(135deg,#0071C5,#004E8C);color:#fff;padding:26px 40px 20px;}",
        ".rmt-hdr h1{font-size:1.9rem;font-weight:700;margin:0 0 4px;}",
        ".rmt-hdr p{margin:0;opacity:.85;font-size:.93rem;}",
        ".kpi-card{background:#fff;border-radius:14px;box-shadow:0 2px 12px rgba(0,71,197,.1);padding:18px 22px;text-align:center;}",
        ".kpi-num{font-size:2.1rem;font-weight:700;line-height:1;}",
        ".kpi-lbl{font-size:.8rem;color:#6b7280;margin-top:4px;text-transform:uppercase;letter-spacing:.04em;}",
        ".section-card{background:#fff;border-radius:12px;box-shadow:0 1px 8px rgba(0,0,0,.07);padding:22px;margin-bottom:22px;}",
        ".chart-md{position:relative;height:340px;}",
        ".chart-lg{position:relative;height:460px;}",
        ".nav-tabs .nav-link{font-weight:500;color:#374151;}",
        ".nav-tabs .nav-link.active{color:#0071C5;border-color:#0071C5 #0071C5 #fff;}",
        "thead.table-dark th{background:#1e40af!important;border-color:#1e3a8a;}",
        ".footer-bar{background:#1e293b;color:#94a3b8;text-align:center;padding:12px;font-size:.82rem;margin-top:28px;}",
        ".badge.bg-success{background:#059669!important;}",
        ".badge.bg-danger{background:#dc2626!important;}",
        "td.weak-cell{background:#fee2e2!important;color:#b91c1c;font-weight:700;}",
        ".rmt-hdr .brand-sub{font-size:1rem;opacity:.92;margin:0 0 6px;font-weight:500;}",
        ".rmt-hdr .proj-badge{background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.35);"
        "border-radius:999px;padding:3px 12px;font-size:.85rem;font-weight:600;margin-left:10px;vertical-align:middle;"
        "display:inline-block;white-space:nowrap;}",
        ".mr-chg{background:#fef3c7!important;}",
        ".odt-var{background:#b45309!important;}",
        ".odt-var-cell{background:#fef3c7;}",
        ".raw-wrap{max-height:70vh;overflow:auto;border:1px solid #dee2e6;border-radius:6px;position:relative;}",
        ".raw-table{border-collapse:separate;border-spacing:0;font-size:.78rem;white-space:nowrap;width:max-content;min-width:100%;}",
        ".raw-table th,.raw-table td{padding:3px 8px;border-right:1px solid #e5e7eb;border-bottom:1px solid #e5e7eb;background:#fff;}",
        ".raw-table thead th{position:sticky;top:0;background:#1e40af;color:#fff;z-index:3;text-align:center;font-weight:600;}",
        ".raw-table thead tr.flt th{background:#1e3a8a;padding:2px 3px;}",
        ".raw-table thead tr.flt input,.raw-table thead tr.flt select{font-size:.72rem;padding:1px 4px;min-width:56px;width:100%;}",
        ".raw-table thead tr.flt input.bad{border-color:#dc2626;background:#fee2e2;}",
        ".raw-table .stk{position:sticky;z-index:2;}",
        ".raw-wrap.nostk .raw-table td.stk{position:static;}",
        ".raw-wrap.nostk .raw-table thead th.stk{left:auto!important;}",
        ".raw-wrap.nostk .raw-table .stk-last{box-shadow:none;}",
        ".raw-table thead .stk{z-index:4;}",
        ".raw-table td.stk{background:#f8fafc;}",
        ".raw-table td.stk-last,.raw-table th.stk-last{box-shadow:inset -2px 0 0 #94a3b8;}",
        ".raw-table td.num{text-align:right;font-variant-numeric:tabular-nums;}",
        ".raw-table th.sortable{cursor:pointer;user-select:none;}",
        ".raw-table th .sa{opacity:.75;font-size:.7rem;margin-left:3px;}",
        ".raw-table tbody tr:hover td{box-shadow:inset 0 0 0 999px rgba(30,64,175,.06);}",
        ".st-fail,td.st-fail{background:#fee2e2!important;color:#b91c1c;font-weight:700;}",
        ".st-warn,td.st-warn{background:#fef3c7!important;color:#92400e;font-weight:600;}",
        ".rs{display:inline-block;min-width:44px;text-align:center;border-radius:999px;padding:0 6px;font-size:.7rem;font-weight:700;color:#fff;}",
        ".rs-PASS{background:#059669;}.rs-WARN{background:#d97706;}.rs-FAIL{background:#dc2626;}",
        ".df-neg{background:#fee2e2;}.df-pos{background:#dcfce7;}",
        ".raw-colmenu{position:relative;}",
        ".raw-colmenu>summary{list-style:none;}.raw-colmenu>summary::-webkit-details-marker{display:none;}",
        ".raw-colmenu-body{position:absolute;left:0;z-index:20;max-width:90vw;background:#fff;border:1px solid #cbd5e1;"
        "border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.12);padding:8px 12px;min-width:220px;font-size:.82rem;}",
        "</style>",
        "</head>",
        "<body>",

        # Header
        '<div class="rmt-hdr">',
        f'<h1>&#x1F4CA; {TOOL_NAME} &mdash; RMT Margin Analysis Report'
        + (f'<span class="proj-badge">{_html.escape(project_name)}</span>' if project_name else '')
        + '</h1>',
        f'<p class="brand-sub">{TOOL_SUBTITLE}</p>',
        f'<p>Generated: {generated} &nbsp;|&nbsp; Source files: {len(files)}'
        f' &nbsp;|&nbsp; Frequencies: {", ".join(freq_labels)}</p>',
        "</div>",

        # KPI cards
        '<div class="container-fluid px-4 py-4">',
        '<div class="row g-3 mb-4">',
        f'<div class="col-6 col-md-3"><div class="kpi-card">'
        f'<div class="kpi-num text-primary">{len(rows)}</div>'
        f'<div class="kpi-lbl">Total RMT Rows</div></div></div>',
        f'<div class="col-6 col-md-3"><div class="kpi-card">'
        f'<div class="kpi-num text-secondary">{len(files)}</div>'
        f'<div class="kpi-lbl">Source Files</div></div></div>',
        f'<div class="col-6 col-md-3"><div class="kpi-card">'
        f'<div class="kpi-num text-info">{len(freqs)}</div>'
        f'<div class="kpi-lbl">Frequencies</div></div></div>',
        '<div class="col-6 col-md-3"><div class="kpi-card">',
        f'<div class="kpi-num">'
        f'<span class="text-success">{n_pass}</span> / '
        f'<span class="text-warning">{n_warn}</span> / '
        f'<span class="text-danger">{n_fail}</span></div>',
        '<div class="kpi-lbl">Pass / Warn / Fail</div></div></div>',
        "</div>",

        # Tabs nav
        '<ul class="nav nav-tabs mb-0" id="rmtTabs">',
        '<li class="nav-item"><a class="nav-link active" data-bs-toggle="tab" href="#ovTab">&#x1F4CB; Overview</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#freqTab">&#x1F4C8; Frequency</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#paramTab">&#x1F52C; Parameters</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#driftTab">&#x1F321; RunTemp</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#tsTab">&#x1F4CB; Training Steps</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#platTab">&#x1F4BE; Platform</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#mrTab">&#x1F9EE; Mode Registers</a></li>',
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#odtTab">&#x1F50C; ODT</a></li>',
        *(
            ['<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#jmpTab">&#x1F4CA; JMP Charts</a></li>']
            if jmp_charts_dir and jmp_charts_dir.is_dir() and any(jmp_charts_dir.glob("*.png"))
            else []
        ),
        '<li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#dataTab">&#x1F5C4; Raw Data</a></li>',
        "</ul>",

        # Tab content
        '<div class="tab-content pt-4">',

        # ── Overview tab ──
        '<div class="tab-pane fade show active" id="ovTab">',
        # ── Margin degradation summary (only when violations exist) — first row ──
        (
            '<div class="section-card">'
            '<h5 class="fw-bold mb-1">&#9888; Margin Degradation &mdash; Below &plusmn;Ref</h5>'
            '<p class="text-muted small mb-3">Samples whose training margin is weaker than the '
            'configured per-parameter reference line (Tab&nbsp;2&nbsp;&middot;&nbsp;Parameters &amp; Axis). '
            'These are the same cells highlighted in the Raw Data tab.</p>'
            + degradation_html +
            '</div>'
        ) if total_degraded > 0 else '',
        '<div class="row g-4 mb-2">',
        '<div class="col-12"><div class="section-card">',
        '<h5 class="fw-bold mb-3">Average Window Width by Parameter</h5>',
        '<div class="chart-md"><canvas id="summaryChart"></canvas></div>',
        '</div></div>',
        '</div>',
        '<div class="section-card">',
        '<h5 class="fw-bold mb-1">Parameter Statistics</h5>',
        '<p class="text-muted small mb-3">'
        'Margin window = (+margin) &minus; (&minus;margin) in UI/mV. '
        '<b>Cpk</b> = (mean&nbsp;&minus;&nbsp;threshold)&nbsp;/&nbsp;3&sigma; '
        '(one-sided process capability &mdash; &ge;1.33 capable, 1.0&ndash;1.33 marginal, '
        '&lt;1.0 at risk). '
        '<b>Guardband</b> = worst-case (Min) window minus threshold; negative = spec violation. '
        '<b>N</b> = sample count.</p>',
        '<div class="table-responsive">',
        '<table class="table table-sm table-hover align-middle mb-0">',
        '<thead class="table-dark"><tr>',
        '<th>Parameter</th><th>Avg+</th><th>Avg&minus;</th><th>Avg Width</th>',
        '<th>Min</th><th>Max</th><th>Median</th><th>&sigma;</th>',
        '<th>Cpk</th><th>Guardband</th><th>Thr</th><th>N</th><th>Status</th>',
        '</tr></thead>',
        '<tbody>' + stats_table_rows + '</tbody>',
        '</table></div></div>',
        # ── Interactive RMT Dashboard (replaces removed JMP dashboard PNG) ──
        '<div class="section-card">',
        '<h5 class="fw-bold mb-1">&#x1F4CA; RMT Dashboard &mdash; Margin Overview</h5>',
        '<p class="text-muted small mb-2">Average margin window width (UI/mV) for every '
        'parameter. Each data source is shown in a distinct color &mdash; use the '
        'checkboxes to enable/disable sources for side-by-side comparison.</p>',
        '<div id="dashSrcBox" class="d-flex flex-wrap gap-3 mb-3 small"></div>',
        '<div class="chart-lg"><canvas id="dashChart"></canvas></div>',
        '</div>',
        '</div>',

        # ── Frequency tab ──
        '<div class="tab-pane fade" id="freqTab">',
        '<div class="section-card">',
        '<h5 class="fw-bold mb-3">Avg Window Width per Parameter by Frequency</h5>',
        '<div class="chart-lg"><canvas id="freqChart"></canvas></div>',
        '</div>',
        '<div class="section-card">',
        '<h5 class="fw-bold mb-3">Plus / Minus Margins by Frequency</h5>',
        '<div class="mb-3"><label class="me-2 fw-semibold">Parameter:</label>',
        f'<select class="form-select form-select-sm d-inline-block w-auto" id="pmSel">{param_opts}</select>',
        '</div>',
        '<div class="chart-md"><canvas id="pmChart"></canvas></div>',
        '</div></div>',

        # ── Parameter detail tab ──
        '<div class="tab-pane fade" id="paramTab">',
        '<div class="section-card">',
        '<h5 class="fw-bold mb-1">Margin by Rank &mdash; select parameter:</h5>',
        '<p class="text-muted small mb-2">Per-rank +/&minus; margins. The dashed orange '
        'lines are the &plusmn;threshold spec limits; bars must clear them on both sides. '
        'Enable <b>worst-case</b> to overlay the weakest single measurement per rank '
        '(the qualification-relevant value) instead of just the average.</p>',
        f'<div class="mb-2">{param_btns}</div>',
        '<div class="form-check form-switch mb-3">',
        '<input class="form-check-input" type="checkbox" id="rankWorst" onchange="showRankChart(window._rankParam)">',
        '<label class="form-check-label small" for="rankWorst">Overlay worst-case (Min) margin per rank</label>',
        '</div>',
        '<div class="chart-lg"><canvas id="rankChart"></canvas></div>',
        '</div></div>',

        # ── RunTemp tab ──
        '<div class="tab-pane fade" id="driftTab">',
        drift_tab_body,
        '</div>',

        # ── Training Steps tab ──
        '<div class="tab-pane fade" id="tsTab">',
        '<div class="section-card">',
        '<h5 class="fw-bold mb-3">&#x1F4CB; MRC Training Steps</h5>',
        *_build_training_steps_html(training_steps),
        '</div>',
        '</div>',

        # ── Platform tab ──
        '<div class="tab-pane fade" id="platTab">',
        *_build_platform_html_section(platform_infos),
        '</div>',

        # ── Mode Registers tab ──
        '<div class="tab-pane fade" id="mrTab">',
        *_build_mr_html_section(platform_infos),
        '</div>',

        # ── ODT tab ──
        '<div class="tab-pane fade" id="odtTab">',
        *_build_odt_html_section(platform_infos),
        '</div>',

        # ── JMP Charts tab (populated only when PNG images are present) ──
        *(
            [
                '<div class="tab-pane fade" id="jmpTab">',
                '<div class="container-fluid px-3 py-3">',
                *_embed_jmp_charts_html(jmp_charts_dir, ref_note),
                '</div>',
                '</div>',
            ]
            if jmp_charts_dir and jmp_charts_dir.is_dir() and any(jmp_charts_dir.glob("*.png"))
            else []
        ),

        # ── Raw data tab ──
        '<div class="tab-pane fade" id="dataTab">',
        '<div class="section-card">',
        '<div class="d-flex flex-wrap align-items-center gap-2 mb-2">',
        '<h5 class="fw-bold mb-0 me-auto">Raw RMT Data</h5>',
        '<div class="btn-group btn-group-sm" role="group" aria-label="Raw data view">',
        '<input type="radio" class="btn-check" name="rawView" id="rawViewM" value="margin" checked>',
        '<label class="btn btn-outline-primary" for="rawViewM">Margins</label>',
        '<input type="radio" class="btn-check" name="rawView" id="rawViewW" value="width">',
        '<label class="btn btn-outline-primary" for="rawViewW">Window width</label>',
        '<input type="radio" class="btn-check" name="rawView" id="rawViewS" value="slack">',
        '<label class="btn btn-outline-primary" for="rawViewS">Slack vs Ref</label>',
        '</div>',
        '<details class="raw-colmenu"><summary class="btn btn-sm btn-outline-secondary">Columns &#9662;</summary>'
        '<div class="raw-colmenu-body" id="rawColMenu"></div></details>',
        '<button class="btn btn-sm btn-outline-success" id="rawCopy" title="Copy the filtered view as '
        'tab-separated text (paste into Excel)">&#x1F4CB; Copy for Excel</button>',
        '<button class="btn btn-sm btn-outline-success" id="rawCsv" title="Download the filtered view as CSV">'
        '&#x2B07; CSV</button>',
        '</div>',
        '<div class="d-flex flex-wrap align-items-center gap-3 mb-2 small">',
        '<label>Show <select id="rawOnly" class="form-select form-select-sm d-inline-block w-auto">'
        '<option value="all">All rows</option>'
        '<option value="warn">Near or below &plusmn;Ref (WARN + FAIL)</option>'
        '<option value="fail">Below &plusmn;Ref (FAIL)</option></select></label>',
        '<label title="A margin is WARN when |margin| is below this multiple of |Ref| (2&times; = the Overview '
        'PASS rule)">Warn band <select id="rawWarn" class="form-select form-select-sm d-inline-block w-auto">'
        '<option value="1.25">1.25&times; Ref</option><option value="1.5">1.5&times; Ref</option>'
        '<option value="2" selected>2&times; Ref (Overview rule)</option></select></label>',
        '<div class="form-check form-switch mb-0"><input class="form-check-input" type="checkbox" id="rawHeat" checked>'
        '<label class="form-check-label" for="rawHeat">Heatmap</label></div>',
        '<input type="search" id="rawSearch" class="form-control form-control-sm w-auto" '
        'placeholder="Search file / rank / phase&hellip;">',
        '<label>Rows <select id="rawPageSize" class="form-select form-select-sm d-inline-block w-auto">'
        '<option>25</option><option selected>50</option><option>100</option><option value="0">All</option>'
        '</select></label>',
        '<button class="btn btn-sm btn-outline-secondary" id="rawReset">&#x21BA; Reset filters</button>',
        '</div>',
        '<p class="text-muted small mb-2">'
        '<span class="st-fail px-1">Red</span> = weaker than &plusmn;Ref, '
        '<span class="st-warn px-1">amber</span> = inside the warn band. Heatmap shades each column from '
        'weakest (red) to strongest (green). Numeric filters accept <code>&lt;15</code>, <code>&gt;=20</code>, '
        '<code>10..20</code> or a value; on margin columns they compare |margin|. Click a header to sort. '
        'Status / Worst / Min slack use the visible parameters.</p>',
        '<div id="rawCount" class="small fw-semibold mb-1"></div>',
        '<div class="raw-wrap" id="rawWrap"><table id="rawTable" class="raw-table">'
        '<thead></thead><tbody></tbody></table></div>',
        '<div class="d-flex align-items-center gap-2 mt-2 small" id="rawPager"></div>',
        '</div>',

        # ── Drift comparison (A vs B) ──
        '<div class="section-card" id="driftCard">',
        '<div class="d-flex flex-wrap align-items-center gap-2 mb-2">',
        '<h5 class="fw-bold mb-0 me-auto">&#x0394; Drift Comparison (A &rarr; B)</h5>',
        '<label class="small">Compare by <select id="dfDim" class="form-select form-select-sm d-inline-block w-auto">'
        '</select></label>',
        '<label class="small">A <select id="dfA" class="form-select form-select-sm d-inline-block w-auto"></select></label>',
        '<label class="small">B <select id="dfB" class="form-select form-select-sm d-inline-block w-auto"></select></label>',
        '<label class="small">Metric <select id="dfMetric" class="form-select form-select-sm d-inline-block w-auto">'
        '<option value="side">Per side (&minus; / +)</option><option value="width">Window width</option>'
        '</select></label>',
        '<button class="btn btn-sm btn-outline-success" id="dfCsv">&#x2B07; &#x0394; CSV</button>',
        '</div>',
        '<p class="text-muted small mb-2">&#x0394; = |margin B| &minus; |margin A| for the same rank (matched on '
        'the other keys and averaged when several rows match). <span class="df-neg px-1">Red</span> = margin '
        'lost, <span class="df-pos px-1">green</span> = margin gained. Default: Boot RMT &rarr; Run RMT for '
        'thermal (DTR) data. Uses the parameters visible in the table above.</p>',
        '<div id="dfSummary" class="mb-2"></div>',
        '<div class="raw-wrap" style="max-height:55vh"><table id="dfTable" class="raw-table">'
        '<thead></thead><tbody></tbody></table></div>',
        '</div>',
        '</div>',

        "</div>",  # tab-content
        "</div>",  # container

        f'<div class="footer-bar">{TOOL_NAME} {TOOL_VERSION} &mdash; {TOOL_SUBTITLE}'
        + (f' &mdash; {_html.escape(project_name)}' if project_name else '')
        + ' &mdash; ' + generated + "</div>",

        # CDN scripts
        '<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>',

        # Embedded data + JS
        "<script>",
        "const D=" + embedded_data + ";",
        "const RAW=" + raw_data_json + ";",
        "const RAWMETA=" + raw_meta_json + ";",
        r"""
const PAL=["#0071C5","#00A3E0","#00305E","#4DB6FF","#059669","#d97706","#7C3AED","#dc2626"];

/* ── Summary chart ── */
const freqDS=D.freqLabels.map((fl,i)=>({
  label:fl+' MT/s',
  data:D.params.map(p=>D.freqSeries[p]?D.freqSeries[p][i]:0),
  backgroundColor:PAL[i%PAL.length]+'CC',
  borderColor:PAL[i%PAL.length],borderWidth:1
}));
new Chart(document.getElementById('summaryChart'),{
  type:'bar',
  data:{labels:D.params,datasets:freqDS},
  options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'top'}},
    scales:{y:{title:{display:true,text:'Avg Window Width'},grid:{color:'#e5e7eb'}},
            x:{grid:{display:false}}}}
});

/* ── Frequency comparison chart ── */
new Chart(document.getElementById('freqChart'),{
  type:'bar',
  data:{labels:D.params,datasets:freqDS},
  options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'right'}},
    scales:{y:{title:{display:true,text:'Avg Window Width'},beginAtZero:true,grid:{color:'#e5e7eb'}},
            x:{grid:{display:false}}}}
});

/* ── Plus/Minus by frequency ── */
let pmC=null;
function buildPmChart(param){
  if(pmC)pmC.destroy();
  const plusD=D.freqPlusSeries[param]||[];
  const minusD=D.freqMinusSeries[param]||[];
  pmC=new Chart(document.getElementById('pmChart'),{
    type:'bar',
    data:{labels:D.freqLabels,datasets:[
      {label:param+'+',data:plusD,backgroundColor:'#05966988',borderColor:'#059669',borderWidth:1},
      {label:param+'-',data:minusD,backgroundColor:'#dc262688',borderColor:'#dc2626',borderWidth:1}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{title:{display:true,text:'Margin'},grid:{color:'#e5e7eb'}},
              x:{grid:{display:false}}}}
  });
}
document.getElementById('pmSel').addEventListener('change',e=>buildPmChart(e.target.value));
buildPmChart(D.params[0]);

/* ── Rank detail chart ── */
let rkC=null;
function showRankChart(param){
  window._rankParam=param;
  document.querySelectorAll('.param-btn').forEach(b=>{
    const active=b.textContent.trim()===param;
    b.className=b.className.replace(/btn-(outline-)?primary/,'btn-'+(active?'':'outline-')+'primary');
  });
  if(rkC)rkC.destroy();
  const rd=D.rankData[param];if(!rd)return;
  const st=D.summary.find(s=>s.param===param);
  const thr=(st&&st.threshold!=null)?st.threshold:10;
  const n=rd.ranks.length;
  const worstEl=document.getElementById('rankWorst');
  const worst=worstEl&&worstEl.checked;
  const ds=[
    {type:'bar',label:param+' +',data:rd.plus,backgroundColor:'#05966988',borderColor:'#059669',borderWidth:1,order:3},
    {type:'bar',label:param+' −',data:rd.minus,backgroundColor:'#dc262688',borderColor:'#dc2626',borderWidth:1,order:3},
    {type:'line',label:'+Threshold ('+thr+')',data:Array(n).fill(thr),borderColor:'#f59e0b',borderWidth:2,borderDash:[6,3],pointRadius:0,fill:false,order:1},
    {type:'line',label:'−Threshold ('+(-thr)+')',data:Array(n).fill(-thr),borderColor:'#f59e0b',borderWidth:2,borderDash:[6,3],pointRadius:0,fill:false,order:1}
  ];
  if(worst){
    ds.push({type:'line',label:param+' + worst',data:rd.plusWorst||[],borderColor:'#065f46',backgroundColor:'#065f46',borderWidth:2,pointRadius:4,pointStyle:'rectRot',fill:false,order:0});
    ds.push({type:'line',label:param+' − worst',data:rd.minusWorst||[],borderColor:'#7f1d1d',backgroundColor:'#7f1d1d',borderWidth:2,pointRadius:4,pointStyle:'rectRot',fill:false,order:0});
  }
  rkC=new Chart(document.getElementById('rankChart'),{
    data:{labels:rd.ranks,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'}},
      scales:{y:{title:{display:true,text:'Margin Value (UI/mV)'},grid:{color:'#e5e7eb'}},
              x:{grid:{display:false},ticks:{maxRotation:45}}}}
  });
}
if(D.params.length>0)showRankChart(D.params[0]);

/* ── Multi-source interactive infrastructure ── */
const SOURCES=[...new Set(RAW.map(r=>r.file))];
const HAS_TEMP=RAW.some(r=>r.rt!=null&&Number(r.rt)!==0);
function srcLabel(s){return s||'Dataset';}
function srcColor(i){return PAL[i%PAL.length];}
function _mean(a){const v=a.filter(x=>x!=null&&!isNaN(Number(x)));return v.length?v.reduce((s,x)=>s+Number(x),0)/v.length:null;}
function buildSrcChecks(boxId,visState,rebuild){
  const box=document.getElementById(boxId);if(!box)return;
  box.innerHTML='';
  SOURCES.forEach((src,si)=>{
    const id=boxId+'_'+si;
    const w=document.createElement('div');w.className='form-check form-check-inline m-0';
    w.innerHTML='<input class="form-check-input" type="checkbox" id="'+id+'" '+
      (visState[src]?'checked':'')+'>'+
      '<label class="form-check-label" for="'+id+'" style="color:'+srcColor(si)+';font-weight:600">'+
      srcLabel(src)+'</label>';
    w.querySelector('input').addEventListener('change',e=>{visState[src]=e.target.checked;rebuild();});
    box.appendChild(w);
  });
}

/* ── RMT Dashboard (Overview) ── */
let dashC=null;
const dashVisible={};SOURCES.forEach(s=>dashVisible[s]=true);
function buildDashboard(){
  const cv=document.getElementById('dashChart');if(!cv)return;
  if(dashC)dashC.destroy();
  const datasets=SOURCES.map((src,si)=>({
    label:srcLabel(src),
    data:D.params.map(p=>_mean(RAW.filter(r=>r.file===src&&r[p+'+']!=null&&r[p+'-']!=null)
        .map(r=>Number(r[p+'+'])-Number(r[p+'-'])))),
    backgroundColor:srcColor(si)+'CC',borderColor:srcColor(si),borderWidth:1,
    hidden:!dashVisible[src]
  }));
  dashC=new Chart(cv,{type:'bar',data:{labels:D.params,datasets},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top'}},
      scales:{y:{title:{display:true,text:'Avg Window Width (UI/mV)'},beginAtZero:true,grid:{color:'#e5e7eb'}},
              x:{grid:{display:false}}}}
  });
}
buildSrcChecks('dashSrcBox',dashVisible,buildDashboard);
buildDashboard();
""",
        drift_init_js,
        r"""
/* ── Raw RMT Data explorer (sticky columns, views, status, heatmap, filters, export) ── */
(function(){
const PARAMS=D.params||[], THR=D.thresholds||{}, META=(typeof RAWMETA!=='undefined'&&RAWMETA)||{};
const MINUS='\u2212';
const S={view:'margin',sortK:'',dir:1,page:0,flt:{},only:'all',warn:2,heat:true,q:'',size:50,
         hidden:new Set(META.constant||[]),
         showTemp:RAW.some(r=>r.bt!=null||r.rt!=null)};
const $id=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const num=v=>(v==null||v===''||Number.isNaN(Number(v)))?null:Number(v);
const fmt=v=>{const n=num(v);if(n==null)return '';return Number.isInteger(n)?String(n):String(+n.toFixed(2));};
const refP=p=>Math.abs(num((THR[p]||{}).plus)??10), refM=p=>Math.abs(num((THR[p]||{}).minus)??10);
const visParams=()=>PARAMS.filter(p=>!S.hidden.has(p));
const natCmp=(a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true});

function sideSt(mag,ref){if(mag==null)return '';if(mag<ref)return 'FAIL';if(mag<ref*S.warn)return 'WARN';return 'PASS';}
function slackSt(s,ref){if(s==null)return '';if(s<0)return 'FAIL';if(s<ref*(S.warn-1))return 'WARN';return 'PASS';}
function widthOf(r,p){const a=num(r[p+'+']),b=num(r[p+'-']);return (a==null||b==null)?null:a-b;}
function widthSt(w,p){if(w==null)return '';const t=refP(p);if(w<t)return 'FAIL';if(w<t*S.warn)return 'WARN';return 'PASS';}
function slackOf(r,p,plus){const v=num(r[p+(plus?'+':'-')]);return v==null?null:Math.abs(v)-(plus?refP(p):refM(p));}

/* Row status / worst parameter / min slack over the visible parameters. */
function derive(r){
  let worst='',slack=null,st='PASS',any=false;
  const rank={PASS:0,WARN:1,FAIL:2};
  visParams().forEach(p=>{[false,true].forEach(plus=>{
    const s=slackOf(r,p,plus);if(s==null)return;any=true;
    const cs=slackSt(s,plus?refP(p):refM(p));if(rank[cs]>rank[st])st=cs;
    if(slack==null||s<slack){slack=s;worst=p+' '+(plus?'+':MINUS);}
  });});
  r._d={st:any?st:'',worst,slack};
}

function columns(){
  const c=[
    {k:'lbl',h:'File',cat:true,get:r=>r.lbl,tip:r=>r.file,stk:true},
    {k:'phase',h:'Phase',cat:true,get:r=>r.phase,stk:true},
    {k:'rank',h:'Rank',cat:true,get:r=>r.rank,stk:true},
    {k:'freq',h:'Freq',cat:true,numSort:true,get:r=>r.freq},
    {k:'gear',h:'Gear',cat:true,numSort:true,get:r=>r.gear},
    {k:'_st',h:'Status',cat:true,get:r=>r._d.st,badge:true},
    {k:'_worst',h:'Worst',cat:true,get:r=>r._d.worst},
    {k:'_slack',h:'Min slack',get:r=>r._d.slack,st:v=>v==null?'':(v<0?'FAIL':'')},
  ];
  if(S.showTemp){
    c.push({k:'bt',h:'BootTemp',cat:true,numSort:true,get:r=>r.bt==null?'':fmt(r.bt)});
    c.push({k:'rt',h:'RunTemp',cat:true,numSort:true,get:r=>r.rt==null?'':fmt(r.rt)});
  }
  visParams().forEach(p=>{
    if(S.view==='width'){
      c.push({k:p+'|W',grp:p,h:'W',get:r=>widthOf(r,p),st:v=>widthSt(v,p),heat:true});
    }else if(S.view==='slack'){
      c.push({k:p+'|s-',grp:p,h:MINUS,get:r=>slackOf(r,p,false),st:v=>slackSt(v,refM(p)),heat:true});
      c.push({k:p+'|s+',grp:p,h:'+',get:r=>slackOf(r,p,true),st:v=>slackSt(v,refP(p)),heat:true});
    }else{
      c.push({k:p+'|-',grp:p,h:MINUS,mag:true,get:r=>num(r[p+'-']),st:v=>v==null?'':sideSt(Math.abs(v),refM(p)),heat:true});
      c.push({k:p+'|+',grp:p,h:'+',mag:true,get:r=>num(r[p+'+']),st:v=>v==null?'':sideSt(Math.abs(v),refP(p)),heat:true});
    }
  });
  return c;
}
/* Value used for sorting / heat / numeric filters: |margin| on margin columns. */
const strength=(c,v)=>v==null?null:(c.mag?Math.abs(v):v);

function mkPred(s){
  s=(s||'').trim();if(!s)return null;
  const N='(-?\\d*\\.?\\d+)';let m;
  if((m=s.match(new RegExp('^'+N+'\\s*\\.\\.\\s*'+N+'$')))){const a=+m[1],b=+m[2];return v=>v>=Math.min(a,b)&&v<=Math.max(a,b);}
  if((m=s.match(new RegExp('^(<=|>=|!=|<|>|=)\\s*'+N+'$')))){const n=+m[2];
    return {'<':v=>v<n,'<=':v=>v<=n,'>':v=>v>n,'>=':v=>v>=n,'=':v=>v===n,'!=':v=>v!==n}[m[1]];}
  if((m=s.match(new RegExp('^'+N+'$')))){const n=+m[1];return v=>v===n;}
  return false;
}

let COLS=[];
function filtered(){
  RAW.forEach(derive);
  const q=S.q.trim().toLowerCase();
  const preds=COLS.map(c=>{const f=S.flt[c.k];if(f==null||f==='')return null;
    if(c.cat)return r=>String(c.get(r))===f;
    const p=mkPred(f);if(!p)return null;return r=>{const v=strength(c,num(c.get(r)));return v!=null&&p(v);};});
  let rows=RAW.filter(r=>{
    if(S.only==='fail'&&r._d.st!=='FAIL')return false;
    if(S.only==='warn'&&r._d.st!=='FAIL'&&r._d.st!=='WARN')return false;
    if(q&&![r.lbl,r.file,r.phase,r.rank,r.freq,r.gear].some(v=>String(v).toLowerCase().includes(q)))return false;
    return preds.every(p=>!p||p(r));
  });
  if(S.sortK){
    const c=COLS.find(x=>x.k===S.sortK);
    if(c){rows=rows.slice().sort((a,b)=>{
      let va=c.get(a),vb=c.get(b);
      if(!c.cat){va=strength(c,num(va));vb=strength(c,num(vb));
        if(va==null&&vb==null)return 0;if(va==null)return 1;if(vb==null)return -1;return (va-vb)*S.dir;}
      if(c.badge){const o={FAIL:0,WARN:1,PASS:2,'':3};return (o[va]-o[vb])*S.dir;}
      return natCmp(va,vb)*S.dir;});}
  }
  return rows;
}

function heatRange(){
  const hr={};
  COLS.forEach(c=>{if(!c.heat)return;let lo=null,hi=null;
    RAW.forEach(r=>{const v=strength(c,num(c.get(r)));if(v==null)return;lo=lo==null?v:Math.min(lo,v);hi=hi==null?v:Math.max(hi,v);});
    hr[c.k]=[lo,hi];});
  return hr;
}
function heatBg(c,v,hr){
  const [lo,hi]=hr[c.k]||[];if(v==null||lo==null||hi==null||hi===lo)return '';
  const t=(strength(c,v)-lo)/(hi-lo);return 'background:hsl('+Math.round(120*t)+',70%,90%)';
}

function renderHead(){
  const th=$id('rawTable').tHead;
  const hasGrp=COLS.some(c=>c.grp);
  let r1='',r2='',r3='';
  const sortMark=k=>S.sortK===k?'<span class="sa">'+(S.dir>0?'\u25B2':'\u25BC')+'</span>':'';
  let i=0;
  while(i<COLS.length){
    const c=COLS[i];
    if(!c.grp){
      r1+='<th class="sortable'+(c.stk?' stk':'')+'" data-k="'+esc(c.k)+'" data-ci="'+i+'"'+(hasGrp?' rowspan="2"':'')+'>'+esc(c.h)+sortMark(c.k)+'</th>';
      i++;continue;
    }
    let j=i;while(j<COLS.length&&COLS[j].grp===c.grp)j++;
    const rp=refP(c.grp),rm=refM(c.grp);
    r1+='<th colspan="'+(j-i)+'" title="Ref: +'+fmt(rp)+' / '+MINUS+fmt(rm)+'">'+esc(c.grp)+'</th>';
    for(let k=i;k<j;k++)r2+='<th class="sortable" data-k="'+esc(COLS[k].k)+'">'+esc(COLS[k].h)+sortMark(COLS[k].k)+'</th>';
    i=j;
  }
  COLS.forEach((c,ci)=>{
    let inner;
    if(c.cat){
      const vals=[...new Set(RAW.map(r=>{derive(r);return String(c.get(r));}))].filter(v=>v!=='').sort(c.numSort?(a,b)=>num(a)-num(b):natCmp);
      inner='<select data-k="'+esc(c.k)+'"><option value="">All</option>'+vals.map(v=>'<option'+(S.flt[c.k]===v?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select>';
    }else{
      inner='<input type="text" data-k="'+esc(c.k)+'" placeholder="&lt;15, 10..20" value="'+esc(S.flt[c.k]||'')+'">';
    }
    r3+='<th class="'+(c.stk?'stk':'')+'" data-ci="'+ci+'">'+inner+'</th>';
  });
  th.innerHTML='<tr class="hr1">'+r1+'</tr>'+(hasGrp?'<tr class="hr2">'+r2+'</tr>':'')+'<tr class="flt">'+r3+'</tr>';
  th.querySelectorAll('th.sortable').forEach(el=>el.addEventListener('click',()=>{
    const k=el.dataset.k;if(S.sortK===k)S.dir=-S.dir;else{S.sortK=k;S.dir=1;}S.page=0;renderHead();renderBody();}));
  th.querySelectorAll('tr.flt select, tr.flt input').forEach(el=>{
    const ev=el.tagName==='SELECT'?'change':'input';
    el.addEventListener(ev,()=>{
      const v=el.value;S.flt[el.dataset.k]=v;
      if(el.tagName==='INPUT'){const ok=mkPred(v)!==false;el.classList.toggle('bad',!ok);}
      S.page=0;renderBody();});
  });
}

function layoutSticky(){
  const tbl=$id('rawTable'), th=tbl.tHead;
  const h1=th.querySelector('tr.hr1'), h2=th.querySelector('tr.hr2'), f=th.querySelector('tr.flt');
  const t1=h1?h1.getBoundingClientRect().height:0, t2=h2?h2.getBoundingClientRect().height:0;
  if(h2)h2.querySelectorAll('th').forEach(x=>x.style.top=t1+'px');
  if(f)f.querySelectorAll('th').forEach(x=>x.style.top=(t1+t2)+'px');
  let left=0;const css=[];const stkIdx=COLS.map((c,i)=>c.stk?i:-1).filter(i=>i>=0);
  const h1cells=[...h1.querySelectorAll('th.stk')];
  stkIdx.forEach((ci,n)=>{css.push('#rawTable .stk-'+ci+'{left:'+left+'px}');left+=h1cells[n]?h1cells[n].getBoundingClientRect().width:0;});
  let st=$id('rawStkCss');if(!st){st=document.createElement('style');st.id='rawStkCss';document.head.appendChild(st);}
  st.textContent=css.join('\n');
  /* Narrow window: frozen columns would hide the data, so let them scroll. */
  const wrap=$id('rawWrap');wrap.classList.toggle('nostk',left>wrap.clientWidth*0.55);
  h1cells.forEach((el,n)=>{el.classList.add('stk-'+stkIdx[n]);if(n===h1cells.length-1)el.classList.add('stk-last');});
  if(f)f.querySelectorAll('th.stk').forEach(el=>{el.classList.add('stk-'+el.dataset.ci);if(+el.dataset.ci===stkIdx[stkIdx.length-1])el.classList.add('stk-last');});
}

let LAST=[];
function renderBody(){
  const rows=filtered();LAST=rows;
  const hr=S.heat?heatRange():{};
  const size=S.size>0?S.size:rows.length||1;
  const pages=Math.max(1,Math.ceil(rows.length/size));if(S.page>=pages)S.page=pages-1;
  const start=S.page*size, page=rows.slice(start,start+size);
  const lastStk=Math.max(...COLS.map((c,i)=>c.stk?i:-1));
  let html='';
  page.forEach(r=>{
    html+='<tr>';
    COLS.forEach((c,ci)=>{
      const raw=c.get(r);let cls=c.stk?'stk stk-'+ci+(ci===lastStk?' stk-last':''):'';let style='';let txt;
      if(c.badge){txt=raw?'<span class="rs rs-'+raw+'">'+raw+'</span>':'';}
      else if(c.cat){txt=esc(raw);}
      else{
        const v=num(raw);txt=fmt(v);cls+=' num';
        const st=c.st?c.st(v):'';
        if(st==='FAIL')cls+=' st-fail';else if(st==='WARN')cls+=' st-warn';else if(S.heat&&c.heat)style=heatBg(c,v,hr);
      }
      const tip=c.tip?' title="'+esc(c.tip(r))+'"':'';
      html+='<td class="'+cls.trim()+'"'+(style?' style="'+style+'"':'')+tip+'>'+txt+'</td>';
    });
    html+='</tr>';
  });
  $id('rawTable').tBodies[0].innerHTML=html||'<tr><td colspan="'+COLS.length+'" class="text-muted">No rows match the filters.</td></tr>';
  const cnt={PASS:0,WARN:0,FAIL:0};rows.forEach(r=>{if(cnt[r._d.st]!=null)cnt[r._d.st]++;});
  $id('rawCount').innerHTML='Showing '+(rows.length?start+1:0)+'&ndash;'+Math.min(start+size,rows.length)+' of '+rows.length+
    ' filtered rows ('+RAW.length+' total) &nbsp;&middot;&nbsp; <span class="rs rs-FAIL">'+cnt.FAIL+'</span> FAIL '+
    '<span class="rs rs-WARN">'+cnt.WARN+'</span> WARN <span class="rs rs-PASS">'+cnt.PASS+'</span> PASS';
  $id('rawPager').innerHTML=pages>1?
    '<button class="btn btn-sm btn-outline-secondary" id="rawPrev"'+(S.page===0?' disabled':'')+'>&laquo; Prev</button>'+
    '<span>Page '+(S.page+1)+' of '+pages+'</span>'+
    '<button class="btn btn-sm btn-outline-secondary" id="rawNext"'+(S.page>=pages-1?' disabled':'')+'>Next &raquo;</button>':'';
  if(pages>1){$id('rawPrev').onclick=()=>{S.page--;renderBody();};$id('rawNext').onclick=()=>{S.page++;renderBody();};}
  requestAnimationFrame(layoutSticky);
  renderDrift();
}

function exportRows(){
  const hdr=['Source file',...COLS.map(c=>c.grp?c.grp+' '+c.h:c.h)];
  const body=LAST.map(r=>[r.file,...COLS.map(c=>{const v=c.get(r);return c.cat||c.badge?String(v??''):fmt(v);})]);
  return [hdr,...body];
}
function toDelim(tab,d){return tab.map(row=>row.map(v=>{v=String(v);return (d===','&&/[",\n]/.test(v))?'"'+v.replace(/"/g,'""')+'"':v;}).join(d)).join('\r\n');}
function download(name,text){
  const blob=new Blob(['\ufeff'+text],{type:'text/csv;charset=utf-8'});const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},500);
}
function copyText(text,btn){
  const done=()=>{const o=btn.innerHTML;btn.innerHTML='&#x2714; Copied';setTimeout(()=>btn.innerHTML=o,1500);};
  if(navigator.clipboard&&window.isSecureContext){navigator.clipboard.writeText(text).then(done,()=>fallback());}else fallback();
  function fallback(){const ta=document.createElement('textarea');ta.value=text;document.body.appendChild(ta);ta.select();
    try{document.execCommand('copy');done();}catch(e){alert('Copy failed: '+e);}ta.remove();}
}

function buildColMenu(){
  const box=$id('rawColMenu');
  const cst=new Set(META.constant||[]);
  box.innerHTML='<div class="fw-semibold mb-1">Parameters</div>'+PARAMS.map(p=>
    '<div class="form-check"><input class="form-check-input" type="checkbox" id="rawCol_'+p+'" data-p="'+esc(p)+'"'+(S.hidden.has(p)?'':' checked')+'>'+
    '<label class="form-check-label" for="rawCol_'+p+'">'+esc(p)+(cst.has(p)?' <span class="text-muted">(constant)</span>':'')+'</label></div>').join('')+
    '<hr class="my-1"><div class="form-check"><input class="form-check-input" type="checkbox" id="rawColTemp"'+(S.showTemp?' checked':'')+'>'+
    '<label class="form-check-label" for="rawColTemp">BootTemp / RunTemp</label></div>';
  box.querySelectorAll('input[data-p]').forEach(el=>el.addEventListener('change',()=>{
    if(el.checked)S.hidden.delete(el.dataset.p);else S.hidden.add(el.dataset.p);rebuild();}));
  $id('rawColTemp').addEventListener('change',e=>{S.showTemp=e.target.checked;rebuild();});
}

function rebuild(){COLS=columns();Object.keys(S.flt).forEach(k=>{if(!COLS.some(c=>c.k===k))delete S.flt[k];});
  if(S.sortK&&!COLS.some(c=>c.k===S.sortK))S.sortK='';renderHead();renderBody();}

/* ── Drift comparison A -> B ── */
const DIMS=[{k:'phase',h:'Phase'},{k:'grp',h:'File group'},{k:'freq',h:'Frequency'},{k:'gear',h:'Gear'},{k:'lbl',h:'File'}];
const distinct=k=>[...new Set(RAW.map(r=>String(r[k])))].filter(v=>v!=='').sort(natCmp);
const DF={dim:'',a:'',b:'',metric:'side',rows:[],cols:[]};
function keyDims(d){
  if(d==='lbl')return ['phase','rank'];
  if(d==='phase')return ['lbl','rank'];
  return ['grp','freq','gear','phase','rank'].filter(x=>x!==d);
}
function initDrift(){
  const avail=DIMS.filter(d=>distinct(d.k).length>=2);
  const sel=$id('dfDim');
  if(!avail.length){$id('driftCard').querySelector('#dfSummary').innerHTML=
    '<div class="alert alert-info py-2 mb-0">Only one data set is loaded &mdash; nothing to compare.</div>';
    ['dfDim','dfA','dfB','dfMetric','dfCsv'].forEach(id=>$id(id).disabled=true);return false;}
  sel.innerHTML=avail.map(d=>'<option value="'+d.k+'">'+d.h+'</option>').join('');
  const phases=distinct('phase');
  DF.dim=(phases.includes('Boot RMT')&&phases.includes('Run RMT'))?'phase':avail[0].k;
  sel.value=DF.dim;fillAB(true);
  sel.addEventListener('change',()=>{DF.dim=sel.value;fillAB(true);renderDrift();});
  $id('dfA').addEventListener('change',e=>{DF.a=e.target.value;renderDrift();});
  $id('dfB').addEventListener('change',e=>{DF.b=e.target.value;renderDrift();});
  $id('dfMetric').addEventListener('change',e=>{DF.metric=e.target.value;renderDrift();});
  $id('dfCsv').addEventListener('click',()=>{
    const hdr=DF.cols.map(c=>c.h),body=DF.rows.map(r=>DF.cols.map(c=>c.txt(r)));
    download('RMT_drift_'+DF.dim+'.csv',toDelim([hdr,...body],','));});
  return true;
}
function fillAB(reset){
  const vals=distinct(DF.dim);
  if(reset){
    if(DF.dim==='phase'&&vals.includes('Boot RMT')&&vals.includes('Run RMT')){DF.a='Boot RMT';DF.b='Run RMT';}
    else{DF.a=vals[0];DF.b=vals[1]||vals[0];}
  }
  const opts=v=>vals.map(x=>'<option'+(x===v?' selected':'')+'>'+esc(x)+'</option>').join('');
  $id('dfA').innerHTML=opts(DF.a);$id('dfB').innerHTML=opts(DF.b);
}
function renderDrift(){
  if(!DF.dim)return;
  const kd=keyDims(DF.dim), ps=visParams(), side=DF.metric==='side';
  const agg=(val)=>{const m=new Map();
    RAW.filter(r=>String(r[DF.dim])===val).forEach(r=>{
      const key=kd.map(k=>r[k]).join('\u0001');
      let e=m.get(key);if(!e){e={keys:kd.map(k=>r[k]),n:{},s:{}};m.set(key,e);}
      ps.forEach(p=>{
        const vals=side?[[p+'|-',num(r[p+'-'])],[p+'|+',num(r[p+'+'])]]:[[p+'|W',widthOf(r,p)]];
        vals.forEach(([k,v])=>{if(v==null)return;v=side?Math.abs(v):v;e.s[k]=(e.s[k]||0)+v;e.n[k]=(e.n[k]||0)+1;});
      });});
    return m;};
  const A=agg(DF.a),B=agg(DF.b);
  const cellKeys=[];ps.forEach(p=>{if(side){cellKeys.push([p,p+'|-',MINUS],[p,p+'|+','+']);}else cellKeys.push([p,p+'|W','W']);});
  const rows=[];
  B.forEach((eb,key)=>{const ea=A.get(key);if(!ea)return;
    const d={keys:eb.keys,v:{}};let worst=null,wk='';
    cellKeys.forEach(([p,k,h])=>{if(!ea.n[k]||!eb.n[k])return;const dv=eb.s[k]/eb.n[k]-ea.s[k]/ea.n[k];d.v[k]=dv;
      if(worst==null||dv<worst){worst=dv;wk=p+' '+h;}});
    d.worst=worst;d.wk=wk;rows.push(d);});
  rows.sort((x,y)=>(x.worst??0)-(y.worst??0));
  const dimH=k=>(DIMS.find(d=>d.k===k)||{h:k==='rank'?'Rank':k}).h;
  const sgn=v=>v==null?'':(v>0?'+':'')+fmt(v);
  DF.cols=[...kd.map((k,i)=>({h:dimH(k),txt:r=>String(r.keys[i])})),
    {h:'Worst \u0394',txt:r=>sgn(r.worst)},{h:'Worst at',txt:r=>r.wk},
    ...cellKeys.map(([p,k,h])=>({h:p+' '+h,k,p,sub:h,txt:r=>sgn(r.v[k])}))];
  DF.rows=rows;
  let mx=0;rows.forEach(r=>Object.values(r.v).forEach(v=>{mx=Math.max(mx,Math.abs(v));}));
  const bg=v=>{if(v==null||v===0||!mx)return '';const a=(0.12+0.55*Math.abs(v)/mx).toFixed(2);
    return v<0?'background:rgba(239,68,68,'+a+')':'background:rgba(34,197,94,'+a+')';};
  const th=$id('dfTable').tHead;
  let h1='',h2='';
  kd.forEach(k=>{h1+='<th rowspan="2">'+esc(dimH(k))+'</th>';});
  h1+='<th rowspan="2">Worst &#x0394;</th><th rowspan="2">Worst at</th>';
  ps.forEach(p=>{h1+='<th colspan="'+(side?2:1)+'">'+esc(p)+'</th>';h2+=side?'<th>'+MINUS+'</th><th>+</th>':'<th>W</th>';});
  th.innerHTML='<tr class="hr1">'+h1+'</tr><tr class="hr2">'+h2+'</tr>';
  const r1h=th.querySelector('tr.hr1').getBoundingClientRect().height;th.querySelectorAll('tr.hr2 th').forEach(x=>x.style.top=r1h+'px');
  $id('dfTable').tBodies[0].innerHTML=rows.length?rows.map(r=>'<tr>'+
    r.keys.map(v=>'<td>'+esc(v)+'</td>').join('')+
    '<td class="num fw-bold" style="'+bg(r.worst)+'">'+sgn(r.worst)+'</td><td>'+esc(r.wk)+'</td>'+
    cellKeys.map(([p,k])=>'<td class="num" style="'+bg(r.v[k])+'">'+sgn(r.v[k])+'</td>').join('')+'</tr>').join('')
    :'<tr><td colspan="'+(kd.length+2+cellKeys.length)+'" class="text-muted">No matching rows between A and B '+
     '(A and B must share the other keys, e.g. the same rank).</td></tr>';
  /* summary: per parameter mean / worst + top-10 worst cells */
  const cells=[];rows.forEach(r=>cellKeys.forEach(([p,k,h])=>{if(r.v[k]!=null)cells.push({p,h,v:r.v[k],r});}));
  const byP=ps.map(p=>{const cs=cells.filter(c=>c.p===p);if(!cs.length)return '';
    const mean=cs.reduce((s,c)=>s+c.v,0)/cs.length;const w=cs.reduce((a,c)=>c.v<a.v?c:a);
    return '<tr><td class="fw-semibold">'+esc(p)+'</td><td class="num" style="'+bg(mean)+'">'+sgn(mean)+'</td>'+
      '<td class="num" style="'+bg(w.v)+'">'+sgn(w.v)+'</td><td>'+esc(w.r.keys.join(' \u00b7 '))+' ('+w.h+')</td></tr>';}).join('');
  const top=cells.slice().sort((a,b)=>a.v-b.v).slice(0,10).filter(c=>c.v<0);
  $id('dfSummary').innerHTML=rows.length?
    '<div class="row g-3"><div class="col-lg-7"><table class="raw-table"><thead><tr><th>Parameter</th><th>Mean &#x0394;</th>'+
    '<th>Worst &#x0394;</th><th>Worst at</th></tr></thead><tbody>'+byP+'</tbody></table></div>'+
    '<div class="col-lg-5"><div class="small fw-semibold mb-1">Top '+top.length+' margin losses ('+esc(DF.a)+' &rarr; '+esc(DF.b)+')</div>'+
    (top.length?'<ol class="small mb-0 ps-3">'+top.map(c=>'<li><b>'+sgn(c.v)+'</b> '+esc(c.p)+' '+c.h+' &mdash; '+
      esc(c.r.keys.join(' \u00b7 '))+'</li>').join('')+'</ol>':'<div class="small text-success">No margin lost.</div>')+
    '</div></div>':'';
}

/* ── wire controls ── */
document.querySelectorAll('input[name="rawView"]').forEach(el=>el.addEventListener('change',()=>{S.view=el.value;S.page=0;rebuild();}));
$id('rawOnly').addEventListener('change',e=>{S.only=e.target.value;S.page=0;renderBody();});
$id('rawWarn').addEventListener('change',e=>{S.warn=+e.target.value;renderHead();renderBody();});
$id('rawHeat').addEventListener('change',e=>{S.heat=e.target.checked;renderBody();});
$id('rawSearch').addEventListener('input',e=>{S.q=e.target.value;S.page=0;renderBody();});
$id('rawPageSize').addEventListener('change',e=>{S.size=+e.target.value;S.page=0;renderBody();});
$id('rawReset').addEventListener('click',()=>{S.flt={};S.q='';S.only='all';S.sortK='';S.page=0;
  $id('rawSearch').value='';$id('rawOnly').value='all';renderHead();renderBody();});
$id('rawCopy').addEventListener('click',e=>copyText(toDelim(exportRows(),'\t'),e.currentTarget));
$id('rawCsv').addEventListener('click',()=>download('RMT_raw_'+S.view+'.csv',toDelim(exportRows(),',')));
document.addEventListener('click',e=>{const d=document.querySelector('.raw-colmenu');if(d&&d.open&&!d.contains(e.target))d.open=false;});
/* Keep the Columns menu inside the window whichever side the toolbar wraps to. */
document.querySelector('.raw-colmenu').addEventListener('toggle',e=>{
  const d=e.currentTarget,b=$id('rawColMenu');if(!d.open)return;
  b.style.left='0';b.style.right='auto';
  const r=b.getBoundingClientRect();
  if(r.right>window.innerWidth-8){b.style.left='auto';b.style.right='0';}
  if(b.getBoundingClientRect().left<8){b.style.right='auto';b.style.left=(8-d.getBoundingClientRect().left)+'px';}
});
window.addEventListener('resize',()=>requestAnimationFrame(layoutSticky));
const tabLink=document.querySelector('a[href="#dataTab"]');
if(tabLink)tabLink.addEventListener('shown.bs.tab',()=>{layoutSticky();renderDrift();});

buildColMenu();initDrift();rebuild();
})();
""",
        "</script>",
        "<script>",
        r"""
function showSnap(prefix, idx){
  document.querySelectorAll('.'+prefix+'-snap').forEach(d=>{d.style.display='none';});
  const el=document.getElementById(prefix+'-snap-'+idx);
  if(el){el.style.display='block';}
  if(prefix==='mr'){filterMr();}
}
function filterMr(){
  const chg=document.getElementById('mrChgOnly'), vr=document.getElementById('mrVarOnly'),
        dec=document.getElementById('mrDecOnly');
  document.querySelectorAll('.mr-table tbody tr').forEach(tr=>{
    let show=true;
    if(chg&&chg.checked&&tr.dataset.chg!=='1')show=false;
    if(vr&&vr.checked&&tr.dataset.var!=='1')show=false;
    if(dec&&dec.checked&&tr.dataset.dec!=='1')show=false;
    tr.style.display=show?'':'none';
  });
}
const JMPIMG={};
document.querySelectorAll('.jmp-img').forEach(img=>{JMPIMG[img.dataset.chart]=img.src;});
function jmpShow(sel){
  const img=sel.closest('.jmp-slot').querySelector('.jmp-img');
  if(JMPIMG[sel.value]){img.src=JMPIMG[sel.value];img.dataset.chart=sel.value;img.alt='JMP Chart: '+sel.value;}
}
function jmpColClass(n){return n==='1'?'col-12':(n==='3'?'col-xl-4':'col-xl-6');}
function jmpSetCols(n){
  document.querySelectorAll('#jmpGrid .jmp-slot').forEach(s=>{
    s.classList.remove('col-12','col-xl-4','col-xl-6');s.classList.add(jmpColClass(n));
  });
}
function jmpAddPanel(){
  const grid=document.getElementById('jmpGrid');const first=grid&&grid.querySelector('.jmp-slot');
  if(!first){return;}
  const clone=first.cloneNode(true);clone.dataset.extra='1';clone.style.display='';
  grid.appendChild(clone);jmpShow(clone.querySelector('.jmp-sel'));
}
function jmpRemove(btn){
  const grid=document.getElementById('jmpGrid');
  const visible=[...grid.querySelectorAll('.jmp-slot')].filter(s=>s.style.display!=='none');
  if(visible.length<=1){return;}
  const slot=btn.closest('.jmp-slot');
  if(slot.dataset.extra==='1'){slot.remove();}else{slot.style.display='none';}
}
function jmpReset(){
  const grid=document.getElementById('jmpGrid');
  grid.querySelectorAll('.jmp-slot[data-extra="1"]').forEach(s=>s.remove());
  grid.querySelectorAll('.jmp-slot').forEach(s=>{
    s.style.display='';
    const sel=s.querySelector('.jmp-sel');sel.value=s.dataset.default;jmpShow(sel);
  });
  const cols=document.getElementById('jmpCols');if(cols){cols.value='2';jmpSetCols('2');}
}
""",
        "</script>",
        "</body>",
        "</html>",
    ]

    html = "\n".join(html_parts)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
    print(f"HTML report saved: {path}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description=f"{TOOL_NAME} {TOOL_VERSION} - {TOOL_SUBTITLE}: extract RMT data from MRC logs")
    parser.add_argument(
        "--version", action="version", version=f"{TOOL_NAME} {TOOL_VERSION}")
    parser.add_argument(
        "--project",
        default=None,
        help=(
            "Project / platform the logs come from (key, code or name from "
            "projects.json, e.g. NVL or WCL). Labels the reports and selects "
            "project-specific START_RMT header aliases. Defaults to the "
            "registry default."
        ),
    )
    parser.add_argument(
        "--input",
        nargs="+",
        required=False,
        default=None,
        help="Input TXT/LOG files or folders. Not required when --jmp-from-csv or --ppt-from-charts is used.",
    )
    parser.add_argument(
        "--pattern",
        default="*.txt",
        help="Glob pattern when input is folder (default: *.txt)",
    )
    parser.add_argument(
        "--dtr",
        action=argparse.BooleanOptionalAction,
        default=False,
        help=(
            "Parse logs as DTR (Dynamic Thermal Range) thermal experiments. "
            "Each log yields two RMT data sets: a Boot-temperature RMT (under "
            "the 'Temp Drift RMT Test' task, before the Hit-Enter pause) and a "
            "Run-temperature RMT (under the 'Rank Margin Tool' task, after it). "
            "Boot/Run temperatures are read from 'PHY Temperature1' and snapped "
            "to the nearest setpoint (0 or 90). Use for Thermal Experiment "
            "profile logs (BCRH / BHRC)."
        ),
    )
    parser.add_argument(
        "--outdir",
        required=True,
        help="Output directory for CSV/XLSX/PPT",
    )
    parser.add_argument(
        "--excel-name",
        default="RMT_Extraction.xlsx",
        help="Excel output filename",
    )
    parser.add_argument(
        "--ppt-name",
        default="RMT_Summary.pptx",
        help="PPT output filename",
    )
    parser.add_argument(
        "--chart-fields",
        default=None,
        help=(
            "Comma-separated chart field names (e.g. RxDqVrefByte,TxVref). "
            "If omitted, defaults to all fields and asks for approval before PPT chart generation."
        ),
    )
    parser.add_argument(
        "--ask-chart-approval",
        action=argparse.BooleanOptionalAction,
        default=True,
        help=(
            "Ask for user approval of chart fields before generating PPT chart slides. "
            "Use --no-ask-chart-approval for non-interactive automation."
        ),
    )
    parser.add_argument(
        "--jmp-exe",
        default=None,
        help=(
            "Path to jmp.exe for optional JMP chart generation. "
            "Example: C:/Program Files/SAS/JMPPRO/17/jmp.exe"
        ),
    )
    parser.add_argument(
        "--generate-jmp-charts",
        action=argparse.BooleanOptionalAction,
        default=False,
        help=(
            "Generate JMP charts from the extended CSV using JSL. "
            "Requires --jmp-exe."
        ),
    )
    parser.add_argument(
        "--jmp-jsl-only",
        action=argparse.BooleanOptionalAction,
        default=False,
        help=(
            "Generate JSL script only and do not launch JMP. "
            "Useful for batch systems or offline review."
        ),
    )
    parser.add_argument(
        "--ppt-template",
        default=None,
        help=(
            "Path to a .pptx template file. When JMP charts are generated, "
            "a second PPT is created from those PNG charts using this template "
            "for styling. Example: C:/Reports/MyTemplate.pptx"
        ),
    )
    parser.add_argument(
        "--jmp-axis-config",
        default=None,
        help=(
            "Path to a JSON file with per-parameter JMP chart axis settings "
            "(Min, Max, Inc, Minor Ticks, Reference Lines). "
            "Edit jmp_axis_settings.json to customise chart scales. "
            "Defaults to no per-parameter overrides when omitted."
        ),
    )
    parser.add_argument(
        "--jmp-from-csv",
        default=None,
        nargs="+",
        metavar="CSV_PATH",
        help=(
            "Skip log extraction. Load one or more existing CSV files (must have "
            "the Extended column layout) and generate JMP charts directly from "
            "them. When several files are given their rows are concatenated and "
            "tagged with a SourceFile column so the merged data sets can be told "
            "apart (colour overlay). Requires --jmp-exe and --outdir."
        ),
    )
    parser.add_argument(
        "--ppt-from-charts",
        default=None,
        metavar="CHARTS_DIR",
        help=(
            "Skip everything. Build a PPT slide-deck from the PNG files already "
            "present in CHARTS_DIR (e.g. a previously generated jmp_charts/ folder). "
            "Requires --outdir. Optionally uses --ppt-template and --ppt-name."
        ),
    )
    parser.add_argument(
        "--jmp-from-excel",
        default=None,
        metavar="EXCEL_PATH",
        help=(
            "Read data from this .xlsx file (All_RMT sheet) and generate "
            "JMP charts, PPT, and HTML report. "
            "Requires --jmp-exe and --outdir."
        ),
    )
    args = parser.parse_args()

    _registry = load_projects()
    _project_key = resolve_project_key(args.project, _registry)
    if _project_key is None:
        print(f"Unknown --project '{args.project}'. Known projects: "
              f"{', '.join(_registry['projects'])} (see projects.json).")
        return 2
    set_param_aliases(project_param_aliases(_project_key, _registry))
    project_name = project_label(_project_key, _registry)
    global _ACTIVE_PROJECT_NAME
    _ACTIVE_PROJECT_NAME = project_name
    print(f"{TOOL_NAME} {TOOL_VERSION} | Project: {project_name}")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------ #
    # Short-circuit 1: build PPT from an existing jmp_charts/ folder only #
    # ------------------------------------------------------------------ #
    if args.ppt_from_charts:
        charts_dir = Path(args.ppt_from_charts)
        if not charts_dir.is_dir():
            print(f"Charts directory not found: {charts_dir}")
            return 2
        pngs = list(charts_dir.glob("*.png"))
        if not pngs:
            print(f"No PNG files found in: {charts_dir}")
            return 2
        template_path = Path(args.ppt_template) if args.ppt_template else None
        ppt_path = outdir / (Path(args.ppt_name).stem + "_JMP_Charts.pptx")
        build_ppt_from_jmp_charts(ppt_path, charts_dir, template_path)
        print(f"PPT created from {len(pngs)} charts: {ppt_path}")
        return 0

    # ------------------------------------------------------------------ #
    # Short-circuit 2: generate JMP charts from an existing CSV (no logs) #
    # ------------------------------------------------------------------ #
    if args.jmp_from_csv:
        # --jmp-from-csv may be a single path or several (nargs="+").
        _csv_args = args.jmp_from_csv if isinstance(args.jmp_from_csv, list) else [args.jmp_from_csv]
        csv_paths = [Path(p) for p in _csv_args]
        _missing = [p for p in csv_paths if not p.is_file()]
        if _missing:
            for p in _missing:
                print(f"CSV file not found: {p}")
            return 2
        if not args.jmp_exe:
            print("--jmp-from-csv requires --jmp-exe.")
            return 4
        # Primary path used for default naming / fallback
        csv_path = csv_paths[0]
        multi_csv = len(csv_paths) > 1
        chart_fields = parse_chart_fields(args.chart_fields)
        axis_cfg = load_axis_config(Path(args.jmp_axis_config) if args.jmp_axis_config else None)

        # ── Load CSV rows so we can produce Excel + HTML ──────────────
        # When several CSVs are supplied their rows are concatenated and each
        # row is tagged with a SourceFile value (the file stem) so the merged
        # data sets can be differentiated by colour overlay in the JMP charts.
        csv_rows: list[dict[str, Any]] = []
        for _src_csv in csv_paths:
            _src_tag = _src_csv.stem
            try:
                with _src_csv.open("r", newline="", encoding="utf-8") as _cf:
                    _reader = csv.DictReader(_cf)
                    # Strip leading/trailing whitespace from all column names so that
                    # external CSVs with headers like "Gear " (trailing space) are
                    # handled correctly — row.get("Gear") would otherwise return None.
                    _reader.fieldnames = [f.strip() for f in _reader.fieldnames]
                    _file_rows = 0
                    for row in _reader:
                        for _col in ("Frequency", "Gear", "BlockIndex"):
                            _v = row.get(_col, "")
                            try:
                                row[_col] = int(_v) if _v not in ("", None) else None
                            except (ValueError, TypeError):
                                row[_col] = None
                        for _col in ("BootTemp", "RunTemp"):
                            _v = row.get(_col, "")
                            try:
                                row[_col] = float(_v) if _v not in ("", None) else None
                            except (ValueError, TypeError):
                                row[_col] = None
                        for _p in PARAMS:
                            for _s in ("+", "-"):
                                _k = f"{_p}{_s}"
                                _v = row.get(_k, "")
                                try:
                                    row[_k] = float(_v) if _v not in ("", None) else None
                                except (ValueError, TypeError):
                                    row[_k] = None
                        # Tag the data set so concatenated files can be told apart.
                        # Preserve any existing SourceFile only for single-file loads.
                        if multi_csv or not row.get("SourceFile"):
                            row["SourceFile"] = _src_tag
                        csv_rows.append(row)
                        _file_rows += 1
                print(f"Loaded {_file_rows} rows from: {_src_csv.name}")
            except Exception as _csv_err:
                print(f"Warning: could not load CSV rows from {_src_csv.name} ({_csv_err}).")
        if multi_csv:
            print(f"Concatenated {len(csv_rows)} rows from {len(csv_paths)} CSV files "
                  f"(differentiated by SourceFile).")

        # ── Write normalized CSV (stripped headers) for JMP ───────────
        # JMP JSL references column names exactly; writing a clean copy ensures
        # "Gear", "Frequency", etc. match what the JSL script expects.
        _clean_csv = outdir / "RMT_Combined_Extended.csv"
        if csv_rows:
            try:
                write_csv(_clean_csv, csv_rows, EXTENDED_COLUMNS)
            except Exception as _wcsv_err:
                print(f"Warning: could not write normalized CSV ({_wcsv_err}); using original.")
                _clean_csv = csv_path

        # ── Excel ────────────────────────────────────────────────────
        if csv_rows:
            try:
                write_excel(outdir / args.excel_name, csv_rows)
                print(f"Excel saved: {outdir / args.excel_name}")
            except Exception as _xl_err:
                print(f"Warning: Excel generation failed: {_xl_err}")

        # ── HTML (initial, without JMP charts) ───────────────────────
        _active_params = chart_fields if chart_fields else PARAMS

        # Reload training steps and platform info saved by a prior --input run.
        _csv_meta_training: list[dict] | None = None
        _csv_meta_platform: list[dict] | None = None
        _meta_path = outdir / "rmt_metadata.json"
        if _meta_path.is_file():
            try:
                import json as _jm
                _saved = _jm.loads(_meta_path.read_text(encoding="utf-8"))
                _csv_meta_training = _saved.get("training_steps") or None
                _csv_meta_platform = _saved.get("platform_infos") or None
                if _csv_meta_training:
                    print(f"Loaded {len(_csv_meta_training)} training steps from metadata sidecar.")
                if _csv_meta_platform:
                    print(f"Loaded {len(_csv_meta_platform)} platform info records from metadata sidecar.")
            except Exception as _ml_err:
                print(f"  [warn] could not load metadata sidecar: {_ml_err}")

        if csv_rows:
            try:
                generate_html_report(
                    outdir / "RMT_Report.html",
                    csv_rows,
                    _active_params,
                    axis_config=axis_cfg,
                    platform_infos=_csv_meta_platform,
                    training_steps=_csv_meta_training,
                )
                print(f"HTML report saved: {outdir / 'RMT_Report.html'}")
            except Exception as _html_err:
                print(f"Warning: HTML report generation failed: {_html_err}")

        # ── JMP JSL + charts ─────────────────────────────────────────
        # Use the normalized CSV (stripped headers) so JMP sees "Gear" not "Gear "
        _jmp_csv_input = _clean_csv if csv_rows else csv_path
        jsl_script = generate_jmp_jsl(_jmp_csv_input, outdir, _active_params, axis_cfg, rows=csv_rows)
        print(f"Generated JMP JSL script: {jsl_script}")
        if not args.jmp_jsl_only:
            run_jmp_script(Path(args.jmp_exe), jsl_script)
            print("JMP script executed successfully.")
            jmp_charts_dir = outdir / "jmp_charts"
            if jmp_charts_dir.exists() and any(jmp_charts_dir.glob("*.png")):
                template_path = Path(args.ppt_template) if args.ppt_template else None
                jmp_ppt_path = outdir / (Path(args.ppt_name).stem + "_JMP_Charts.pptx")
                try:
                    build_ppt_from_jmp_charts(jmp_ppt_path, jmp_charts_dir, template_path)
                except Exception as ppt_err:
                    print(f"Warning: could not build JMP chart PPT: {ppt_err}")
                # Regenerate HTML with embedded JMP charts
                if csv_rows:
                    try:
                        generate_html_report(
                            outdir / "RMT_Report.html",
                            csv_rows,
                            _active_params,
                            axis_config=axis_cfg,
                            jmp_charts_dir=jmp_charts_dir,
                            platform_infos=_csv_meta_platform,
                            training_steps=_csv_meta_training,
                        )
                        print("HTML report updated with embedded JMP charts.")
                    except Exception as _html_jmp_err:
                        print(f"Warning: could not embed JMP charts in HTML: {_html_jmp_err}")
                    except Exception as _html_jmp_err:
                        print(f"Warning: could not embed JMP charts in HTML: {_html_jmp_err}")
            else:
                print("No JMP PNG charts found — skipping JMP PPT/HTML update.")
        else:
            print("JSL-only mode. JMP was not launched.")
        return 0

    # ------------------------------------------------------------------ #
    # Short-circuit 3: generate from an existing Excel file (.xlsx)        #
    # ------------------------------------------------------------------ #
    if args.jmp_from_excel:
        _excel_path = Path(args.jmp_from_excel)
        if not _excel_path.is_file():
            print(f"Excel file not found: {_excel_path}")
            return 2
        if not args.jmp_exe:
            print("--jmp-from-excel requires --jmp-exe.")
            return 4
        chart_fields = parse_chart_fields(args.chart_fields)
        axis_cfg = load_axis_config(Path(args.jmp_axis_config) if args.jmp_axis_config else None)
        _active_params = chart_fields if chart_fields else PARAMS

        # ── Read Excel into rows ────────────────────────────────────
        xl_rows: list[dict[str, Any]] = []
        try:
            from openpyxl import load_workbook as _lw
            _wb = _lw(str(_excel_path), read_only=True, data_only=True)
            _ws = _wb["All_RMT"] if "All_RMT" in _wb.sheetnames else _wb.active
            _hdrs = [
                str(c.value) if c.value is not None else ""
                for c in next(_ws.iter_rows(min_row=1, max_row=1))
            ]
            for _rv in _ws.iter_rows(min_row=2, values_only=True):
                _rec: dict[str, Any] = dict(zip(_hdrs, _rv))
                for _col in ("Frequency", "Gear", "BlockIndex"):
                    _v = _rec.get(_col)
                    try:
                        _rec[_col] = int(_v) if _v not in ("", None) else None
                    except (ValueError, TypeError):
                        _rec[_col] = None
                for _col in ("BootTemp", "RunTemp"):
                    _v = _rec.get(_col)
                    try:
                        _rec[_col] = float(_v) if _v not in ("", None) else None
                    except (ValueError, TypeError):
                        _rec[_col] = None
                for _p in PARAMS:
                    for _s in ("+", "-"):
                        _k = f"{_p}{_s}"
                        _v = _rec.get(_k)
                        try:
                            _rec[_k] = float(_v) if _v not in ("", None) else None
                        except (ValueError, TypeError):
                            _rec[_k] = None
                xl_rows.append(_rec)
            _wb.close()
            print(f"Loaded {len(xl_rows)} rows from: {_excel_path.name}")
        except Exception as _xl_read_err:
            print(f"Failed to read Excel: {_xl_read_err}")
            return 2

        # ── Write CSV for JMP ───────────────────────────────────────
        _csv_for_jmp = outdir / "RMT_Combined_Extended.csv"
        write_csv(_csv_for_jmp, xl_rows, EXTENDED_COLUMNS)
        print(f"CSV written: {_csv_for_jmp}")

        # ── HTML (initial, without JMP charts) ──────────────────────
        try:
            generate_html_report(
                outdir / "RMT_Report.html",
                xl_rows,
                _active_params,
                axis_config=axis_cfg,
            )
            print(f"HTML report saved: {outdir / 'RMT_Report.html'}")
        except Exception as _html_err:
            print(f"Warning: HTML report generation failed: {_html_err}")

        # ── JMP charts ──────────────────────────────────────────────
        jsl_script = generate_jmp_jsl(_csv_for_jmp, outdir, _active_params, axis_cfg, rows=xl_rows)
        print(f"Generated JMP JSL script: {jsl_script}")
        if not args.jmp_jsl_only:
            run_jmp_script(Path(args.jmp_exe), jsl_script)
            print("JMP script executed successfully.")
            jmp_charts_dir = outdir / "jmp_charts"
            if jmp_charts_dir.exists() and any(jmp_charts_dir.glob("*.png")):
                template_path = Path(args.ppt_template) if args.ppt_template else None
                jmp_ppt_path = outdir / (Path(args.ppt_name).stem + "_JMP_Charts.pptx")
                try:
                    build_ppt_from_jmp_charts(jmp_ppt_path, jmp_charts_dir, template_path)
                except Exception as ppt_err:
                    print(f"Warning: could not build JMP chart PPT: {ppt_err}")
                try:
                    generate_html_report(
                        outdir / "RMT_Report.html",
                        xl_rows,
                        _active_params,
                        axis_config=axis_cfg,
                        jmp_charts_dir=jmp_charts_dir,
                    )
                    print("HTML report updated with embedded JMP charts.")
                except Exception as _html_jmp_err:
                    print(f"Warning: could not embed JMP charts in HTML: {_html_jmp_err}")
            else:
                print("No JMP PNG charts found — skipping JMP PPT/HTML update.")
        else:
            print("JSL-only mode. JMP was not launched.")
        return 0

    # ------------------------------------------------------------------ #
    # Normal path: extract from log files                                  #
    # ------------------------------------------------------------------ #
    if not args.input:
        print("--input is required unless --jmp-from-csv, --jmp-from-excel, or --ppt-from-charts is used.")
        return 2

    input_files = collect_input_files(args.input, args.pattern)

    if not input_files:
        print("No input files found. Check --input and --pattern.")
        return 1

    all_rows: list[dict[str, Any]] = []
    platform_infos: list[dict] = []
    all_training_steps: list[dict] = []
    for file_path in input_files:
        try:
            text = read_text_file(file_path)
        except FileNotFoundError:
            print(f"Skipping missing file: {file_path}")
            continue
        except PermissionError:
            print(f"Skipping file due to permission error: {file_path}")
            continue

        rows = (parse_dtr_rmt_from_text(text, file_path.name)
                if args.dtr else parse_rmt_from_text(text, file_path.name))
        all_rows.extend(rows)
        try:
            pinfo = parse_platform_info(text, file_path.name)
        except Exception as _pi_exc:
            pinfo = None
            print(f"  [warn] platform info skipped for {file_path.name}: {_pi_exc}")
        if pinfo is not None:
            try:
                pinfo["mr_odt"] = parse_mr_odt_info(text, file_path.name)
            except Exception as _mr_exc:
                print(f"  [warn] MR/ODT info skipped for {file_path.name}: {_mr_exc}")
            platform_infos.append(pinfo)
        try:
            ts = parse_training_steps(text, file_path.name)
            all_training_steps.extend(ts)
        except Exception as _ts_exc:
            print(f"  [warn] training steps skipped for {file_path.name}: {_ts_exc}")

    if not all_rows:
        print("No START_RMT blocks found in input files.")
        return 2

    # Combined outputs
    similar_csv = outdir / "RMT_Combined_Similar.csv"
    extended_csv = outdir / "RMT_Combined_Extended.csv"
    write_csv(similar_csv, all_rows, CSV_COLUMNS)
    write_csv(extended_csv, all_rows, EXTENDED_COLUMNS)

    # ── Save metadata sidecar so subsequent --jmp-from-csv runs can
    #    re-attach training steps and platform info to the final HTML ──
    import json as _json_meta
    _meta = {
        "tool": f"{TOOL_NAME} {TOOL_VERSION}",
        "project": _project_key,
        "training_steps": all_training_steps if all_training_steps else [],
        "platform_infos": platform_infos   if platform_infos   else [],
    }
    _meta_path = outdir / "rmt_metadata.json"
    try:
        _meta_path.write_text(_json_meta.dumps(_meta, separators=(",", ":")),
                              encoding="utf-8")
    except Exception as _mw_err:
        print(f"  [warn] could not write metadata sidecar: {_mw_err}")

    # Per-file CSV outputs (similar format)
    per_file_dir = outdir / "csv_by_file"
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in all_rows:
        grouped[row["SourceFile"]].append(row)

    for source_file, rows in grouped.items():
        write_csv(per_file_dir / f"{Path(source_file).stem}_RMT.csv", rows, CSV_COLUMNS)

    write_excel(outdir / args.excel_name, all_rows)

    chart_fields = parse_chart_fields(args.chart_fields)
    if args.ask_chart_approval:
        if not sys.stdin.isatty():
            print(
                "Interactive chart approval requested but no TTY is available. "
                "Use --no-ask-chart-approval for automation."
            )
            chart_fields = []
        else:
            approved, chart_fields = ask_user_chart_approval(chart_fields)
            if not approved:
                print("PPT generation stopped: chart fields were not approved by user.")
                return 3

    _axis_cfg_main = load_axis_config(
        Path(args.jmp_axis_config) if args.jmp_axis_config else None
    )

    try:
        build_ppt_professional(
            outdir / args.ppt_name,
            all_rows,
            chart_fields,
            axis_config=_axis_cfg_main,
            platform_infos=platform_infos if platform_infos else None,
            training_steps=all_training_steps if all_training_steps else None,
        )
    except RuntimeError as err:
        print(str(err))
        print("CSV and Excel were created. Install python-pptx and rerun to generate PPT.")

    try:
        generate_html_report(
            outdir / "RMT_Report.html",
            all_rows,
            chart_fields,
            axis_config=_axis_cfg_main,
            platform_infos=platform_infos if platform_infos else None,
            training_steps=all_training_steps if all_training_steps else None,
        )
    except Exception as html_err:
        print(f"Warning: HTML report generation failed: {html_err}")

    if args.generate_jmp_charts:
        if not args.jmp_exe:
            print("JMP chart generation requested but --jmp-exe was not provided.")
            return 4

        try:
            axis_cfg = load_axis_config(Path(args.jmp_axis_config) if args.jmp_axis_config else None)
            jsl_script = generate_jmp_jsl(extended_csv, outdir, chart_fields if chart_fields else PARAMS, axis_cfg, rows=all_rows)
            print(f"Generated JMP JSL script: {jsl_script}")
            if not args.jmp_jsl_only:
                run_jmp_script(Path(args.jmp_exe), jsl_script)
                print("JMP script executed successfully.")
                # Build a PPT from the generated JMP chart PNGs
                jmp_charts_dir = outdir / "jmp_charts"
                if jmp_charts_dir.exists() and any(jmp_charts_dir.glob("*.png")):
                    template_path = Path(args.ppt_template) if args.ppt_template else None
                    jmp_ppt_path  = outdir / (Path(args.ppt_name).stem + "_JMP_Charts.pptx")
                    try:
                        build_ppt_from_jmp_charts(jmp_ppt_path, jmp_charts_dir, template_path)
                    except Exception as ppt_err:
                        print(f"Warning: could not build JMP chart PPT: {ppt_err}")
                    # Re-generate HTML report with JMP charts embedded as base64 images
                    try:
                        generate_html_report(
                            outdir / "RMT_Report.html",
                            all_rows,
                            chart_fields,
                            axis_config=_axis_cfg_main,
                            platform_infos=platform_infos if platform_infos else None,
                            training_steps=all_training_steps if all_training_steps else None,
                            jmp_charts_dir=jmp_charts_dir,
                        )
                        print("HTML report updated with embedded JMP charts.")
                    except Exception as _html_jmp_err:
                        print(f"Warning: could not embed JMP charts in HTML: {_html_jmp_err}")
                else:
                    print("No JMP PNG charts found - skipping JMP PPT creation.")
            else:
                print("JSL-only mode enabled. JMP was not launched.")
        except (FileNotFoundError, PermissionError, OSError, subprocess.CalledProcessError) as err:
            print(f"Failed to generate or run JMP charts: {err}")
            return 5

    print(f"Done. Parsed {len(all_rows)} RMT rows from {len(input_files)} files.")
    print(f"Output folder: {outdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
