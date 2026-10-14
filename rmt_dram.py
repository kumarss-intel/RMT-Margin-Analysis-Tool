"""DRAM die / process-node decoding for the MarginIQ Platform report.

Two independent sources are combined and kept separate in the output:

* **Log facts** - decoded from the raw DDR5 SPD dump in the MRC log
  (JEDEC JESD400-5 byte map): DRAM stepping, die density/package, voltages,
  fastest speed bin, CL-tRCD-tRP, tRFC, module organisation.
* **Knowledge base** - ``dram_nodes.json`` (repo root): manufacturer part
  number nomenclature and per die-revision process-node data. Every value
  there carries a ``basis`` label (Published / Industry consensus /
  Engineering Estimate / Not publicly disclosed) and source URLs, so the
  report never shows an unlabelled guess.

Only the Python standard library is used.
"""
from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
DRAM_NODES_FILE = HERE / "dram_nodes.json"

# ---------------------------------------------------------------------------
# JEDEC JESD400-5 (DDR5 SPD) decode tables
# ---------------------------------------------------------------------------
_DEVICE_TYPE = {0x12: "DDR5", 0x13: "LPDDR5", 0x14: "DDR5 NVDIMM-P", 0x15: "LPDDR5X"}
_MODULE_TYPE = {
    0x1: "RDIMM", 0x2: "UDIMM", 0x3: "SODIMM", 0x4: "LRDIMM", 0x5: "CUDIMM",
    0x6: "CSODIMM", 0x7: "MRDIMM", 0x8: "CAMM2", 0xA: "DDIMM", 0xB: "Solder-down",
}
_DIE_DENSITY_GB = {1: 4, 2: 8, 3: 12, 4: 16, 5: 24, 6: 32, 7: 48, 8: 64}
_IO_WIDTH = {0: 4, 1: 8, 2: 16, 3: 32}
_VDD_NOMINAL = {0: 1.1}          # bits 7:4 of bytes 16 / 17
_VPP_NOMINAL = {0: 1.8}          # bits 7:4 of byte 18
_JEDEC_BINS = (3200, 3600, 4000, 4400, 4800, 5200, 5600, 6000, 6400,
               6800, 7200, 7600, 8000, 8400, 8800, 9200, 9600)


def _u16(spd: dict[int, int], lo: int) -> int | None:
    if lo in spd and lo + 1 in spd:
        return spd[lo] | (spd[lo + 1] << 8)
    return None


def decode_spd_ddr5(spd: dict[int, int]) -> dict[str, Any]:
    """Decode the DDR5 SPD bytes relevant to die / speed identification.

    ``spd`` maps byte offset -> value (as parsed from the MRC SPD dump).
    Returns only the fields that are present; never raises on short dumps.
    """
    out: dict[str, Any] = {}
    if not spd:
        return out

    if 2 in spd:
        out["memory_type"] = _DEVICE_TYPE.get(spd[2], f"0x{spd[2]:02X}")
    if 3 in spd:
        out["spd_module_type"] = _MODULE_TYPE.get(spd[3] & 0x0F, f"0x{spd[3] & 0x0F:X}")

    if 4 in spd:
        dens = _DIE_DENSITY_GB.get(spd[4] & 0x1F)
        if dens:
            out["die_density_gb"] = dens
        dpp_code = (spd[4] >> 5) & 0x7
        out["package"] = ("Monolithic (1 die / package)" if dpp_code == 0
                          else f"Multi-die package (SPD byte 4 code {dpp_code:03b}b)")
        out["monolithic"] = dpp_code == 0
    if 6 in spd:
        w = _IO_WIDTH.get((spd[6] >> 5) & 0x7)
        if w:
            out["io_width"] = w
    if 5 in spd:
        out["row_bits"] = 16 + (spd[5] & 0x1F)
        out["col_bits"] = 10 + ((spd[5] >> 5) & 0x7)
    if out.get("die_density_gb") and out.get("io_width"):
        depth_g = out["die_density_gb"] // out["io_width"]
        out["die_organization"] = (
            f'{out["die_density_gb"]}Gb ({depth_g}G x{out["io_width"]})'
            if depth_g else f'{out["die_density_gb"]}Gb (x{out["io_width"]})'
        )

    for key, off, table in (("vdd_v", 16, _VDD_NOMINAL), ("vddq_v", 17, _VDD_NOMINAL),
                            ("vpp_v", 18, _VPP_NOMINAL)):
        if off in spd:
            v = table.get((spd[off] >> 4) & 0xF)
            if v is not None:
                out[key] = v

    tck = _u16(spd, 20)
    if tck:
        out["tck_min_ps"] = tck
        raw_rate = 2_000_000 / tck
        out["max_data_rate"] = min(_JEDEC_BINS, key=lambda b: abs(b - raw_rate))
        tck_ns = 2000.0 / out["max_data_rate"]   # exact JEDEC tCK for the bin
        taa, trcd, trp = _u16(spd, 30), _u16(spd, 32), _u16(spd, 34)
        if taa and trcd and trp:
            # JESD79-5 rounding: nCK = ceil(t / tCK - 0.01 guard-band); CL is even.
            def _nck(ps: int) -> int:
                return math.ceil(ps / 1000.0 / tck_ns - 0.01)
            cl = _nck(taa)
            cl += cl % 2
            out["taa_ns"] = taa / 1000.0
            out["trcd_ns"] = trcd / 1000.0
            out["trp_ns"] = trp / 1000.0
            out["cl_trcd_trp"] = f"{cl}-{_nck(trcd)}-{_nck(trp)}"
        if all(b in spd for b in range(24, 29)):
            cls = [20 + 2 * i for i in range(40) if spd[24 + i // 8] >> (i % 8) & 1]
            if cls:
                out["cas_latencies"] = cls

    for key, off in (("trfc1_ns", 42), ("trfc2_ns", 44), ("trfcsb_ns", 46)):
        v = _u16(spd, off)
        if v:
            out[key] = v

    if 234 in spd:
        out["package_ranks"] = ((spd[234] >> 3) & 0x7) + 1
    if 235 in spd:
        out["subchannels"] = ((spd[235] >> 5) & 0x3) + 1
        out["subchannel_width"] = {0: 8, 1: 16, 2: 32, 3: 64}.get(spd[235] & 0x7)
    if out.get("monolithic") and all(out.get(k) for k in (
            "subchannels", "subchannel_width", "io_width", "package_ranks", "die_density_gb")):
        dies_per_rank = out["subchannels"] * out["subchannel_width"] // out["io_width"]
        out["dies_total"] = dies_per_rank * out["package_ranks"]
        out["module_capacity_gb"] = out["dies_total"] * out["die_density_gb"] // 8

    if 551 in spd:
        out["module_revision"] = f"0x{spd[551]:02X}"
    if 552 in spd and 553 in spd:
        out["dram_mfr_id"] = f"0x{spd[552]:02X}{spd[553]:02X}"
    if 554 in spd:
        b = spd[554]
        out["dram_stepping_raw"] = f"0x{b:02X}"
        if 0x41 <= b <= 0x5A:              # printable A-Z
            out["dram_stepping"] = chr(b)
    return out


# ---------------------------------------------------------------------------
# Knowledge base (dram_nodes.json)
# ---------------------------------------------------------------------------
_KB_CACHE: dict[str, Any] | None = None


def load_dram_nodes(path: Path | None = None) -> dict[str, Any]:
    """Load ``dram_nodes.json``; returns an empty registry when unavailable."""
    global _KB_CACHE
    if _KB_CACHE is not None and path is None:
        return _KB_CACHE
    src = path or DRAM_NODES_FILE
    try:
        data = json.loads(src.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("top level must be an object")
    except FileNotFoundError:
        print(f"Warning: {src.name} not found; DRAM process-node info will be omitted.")
        data = {}
    except (PermissionError, ValueError, json.JSONDecodeError) as exc:
        print(f"Warning: could not read {src.name} ({exc}); "
              "DRAM process-node info will be omitted.")
        data = {}
    if path is None:
        _KB_CACHE = data
    return data


def _row(field: str, value: Any, basis: str, sources: list[str] | None = None,
         note: str = "") -> dict[str, Any]:
    return {"field": field, "value": str(value), "basis": basis,
            "sources": list(sources or []), "note": note}


def _vendor_entries(kb: dict[str, Any], manufacturer: str | None) -> list[tuple[str, dict]]:
    vendors = kb.get("vendors") or {}
    if manufacturer:
        needle = manufacturer.strip().lower()
        hits = [(k, v) for k, v in vendors.items()
                if needle in {a.lower() for a in [k, *(v.get("aliases") or [])]}]
        if hits:
            return hits
    return list(vendors.items())


def decode_part_number(part_number: str | None, manufacturer: str | None = None,
                       kb: dict[str, Any] | None = None) -> dict[str, Any] | None:
    """Decode a module / component MPN with the nomenclature in ``dram_nodes.json``.

    Returns ``{"vendor", "kind", "groups", "fields": [...], "sources"}`` or
    ``None`` when no decoder matches.
    """
    if not part_number:
        return None
    kb = kb if kb is not None else load_dram_nodes()
    pn = part_number.strip().upper()
    for vendor, ventry in _vendor_entries(kb, manufacturer):
        for dec in ventry.get("decoders") or []:
            try:
                m = re.match(dec["regex"], pn)
            except (KeyError, re.error):
                continue
            if not m:
                continue
            groups = {k: v for k, v in m.groupdict().items() if v is not None}
            fields = []
            for key, spec in (dec.get("fields") or {}).items():
                if key not in groups:
                    continue
                code = groups[key]
                codes = spec.get("codes") or {}
                meaning = codes.get(code)
                if meaning is None and "*" in codes:
                    meaning = codes["*"].replace("{code}", code)
                fields.append({
                    "key": key, "label": spec.get("label", key), "code": code,
                    "meaning": meaning if meaning is not None else
                    spec.get("unknown", "Code not in knowledge base"),
                    "known": meaning is not None,
                    "basis": spec.get("basis", "Engineering Estimate") if meaning is not None
                    else "Not publicly disclosed",
                })
            return {"vendor": vendor, "kind": dec.get("kind", ""), "groups": groups,
                    "fields": fields, "sources": dec.get("sources") or []}
    return None


def _find_die(kb: dict[str, Any], vendor: str, mem_type: str | None,
              density: int | None, die_rev: str | None) -> dict[str, Any] | None:
    for die in kb.get("dies") or []:
        if die.get("vendor") != vendor or die.get("die_rev") != die_rev:
            continue
        if mem_type and die.get("memory_type") not in (None, mem_type):
            continue
        if density and die.get("die_density_gb") not in (None, density):
            continue
        return die
    return None


def _first_spd_ext(dimms: list[dict]) -> dict[str, Any]:
    for d in dimms:
        if d.get("spd_ext"):
            return d["spd_ext"]
    return {}


def _module_year_month(date_code: str | None) -> tuple[int, int] | None:
    """SPD date code '0xYYWW' (BCD) -> (year, approx. month)."""
    try:
        c = int(str(date_code), 16)
    except (TypeError, ValueError):
        return None
    yy, ww = (c >> 8) & 0xFF, c & 0xFF
    year = 2000 + (yy >> 4) * 10 + (yy & 0xF)
    week = (ww >> 4) * 10 + (ww & 0xF)
    if not 1 <= week <= 53:
        return None
    return year, min(12, max(1, math.ceil(week * 12 / 52)))


def build_die_report(part_number: str, dimms: list[dict],
                     kb: dict[str, Any] | None = None) -> dict[str, Any]:
    """Build the delta "DRAM die & process node" report for one part number."""
    kb = kb if kb is not None else load_dram_nodes()
    first = dimms[0] if dimms else {}
    manufacturer = first.get("manufacturer")
    spd = _first_spd_ext(dimms)
    dec = decode_part_number(part_number, manufacturer, kb)
    vendor = dec["vendor"] if dec else (manufacturer or "")
    log = "Log (SPD)"

    mpn_rev = (dec or {}).get("groups", {}).get("die_rev")
    spd_rev = spd.get("dram_stepping")
    die_rev = spd_rev or mpn_rev
    density = spd.get("die_density_gb")
    mem_type = spd.get("memory_type")
    die = _find_die(kb, vendor, mem_type, density, die_rev) if die_rev else None
    family = (kb.get("node_families") or {}).get(die.get("family")) if die else None

    # ── Step 1: decode & node identification ─────────────────────────────
    step1: list[dict] = []
    if spd.get("dram_mfr_id"):
        step1.append(_row("DRAM manufacturer", f'{manufacturer or "?"} '
                          f'(DRAM JEDEC ID {spd["dram_mfr_id"]}, SPD bytes 552-553)', log))
    if mem_type:
        step1.append(_row("Memory type", f'{mem_type} {spd.get("spd_module_type", "")}'.strip()
                          + " (SPD bytes 2-3)", log))
    if spd.get("die_organization"):
        step1.append(_row("Die density / organisation",
                          f'{spd["die_organization"]}, {spd.get("package", "")} (SPD bytes 4-6)', log))
    if die_rev:
        parts = []
        if spd_rev:
            parts.append(f'SPD byte 554 (DRAM stepping) = {spd["dram_stepping_raw"]} '
                         f'(ASCII "{spd_rev}")')
        if mpn_rev:
            parts.append(f'part-number die-revision field = "{mpn_rev}"')
        agree = ""
        if spd_rev and mpn_rev:
            agree = " - consistent" if spd_rev == mpn_rev else " - MISMATCH, check the module"
        step1.append(_row("Die revision (stepping)",
                          f'{die_rev}-die: ' + "; ".join(parts) + agree,
                          log if spd_rev else "Part number"))
    elif spd.get("dram_stepping_raw"):
        step1.append(_row("Die revision (stepping)",
                          f'SPD byte 554 = {spd["dram_stepping_raw"]} (vendor-defined encoding)', log))
    if die:
        node = die.get("node") or {}
        step1.append(_row("Process node generation", node.get("value", "?"),
                          node.get("basis", "Engineering Estimate"), node.get("sources"),
                          node.get("note", "")))
        ym = _module_year_month(first.get("date_code"))
        cands = die.get("candidates") or []
        if ym and cands:
            in_mp = [c["name"] for c in cands
                     if str(c.get("mass_production", "9999-99")) <= f"{ym[0]:04d}-{ym[1]:02d}"]
            step1.append(_row(
                "Nodes in mass production by module date",
                f'{first.get("date_display") or first.get("date_code")}: '
                + (", ".join(in_mp) if in_mp else "none of the listed candidates")
                + " (date code from SPD bytes 515-516; vendor mass-production start dates)",
                "Engineering Estimate", node.get("sources"),
                "Narrows the candidates only; does not identify the node."))
    else:
        step1.append(_row("Process node generation",
                          "Not in knowledge base (add the die to dram_nodes.json)",
                          "Not available"))
    if spd.get("max_data_rate"):
        txt = f'{mem_type or "DDR"}-{spd["max_data_rate"]}'
        if spd.get("cl_trcd_trp"):
            txt += f' {spd["cl_trcd_trp"]}'
        txt += f' (tCKAVGmin {spd["tck_min_ps"]} ps'
        if spd.get("taa_ns") is not None:
            txt += f', tAA/tRCD/tRP {spd["taa_ns"]:g}/{spd["trcd_ns"]:g}/{spd["trp_ns"]:g} ns'
        step1.append(_row("Nominal speed bin", txt + ")", log))
    if spd.get("vdd_v"):
        step1.append(_row("Voltage", f'VDD {spd["vdd_v"]} V / VDDQ {spd.get("vddq_v", "?")} V / '
                          f'VPP {spd.get("vpp_v", "?")} V (SPD bytes 16-18, nominal)', log))
    if spd.get("module_capacity_gb"):
        step1.append(_row("Module organisation",
                          f'{spd["module_capacity_gb"]} GB module = {spd["dies_total"]} x '
                          f'{density}Gb x{spd["io_width"]} dies ({spd["subchannels"]} sub-channels '
                          f'x {spd["subchannel_width"]}b x {spd["package_ranks"]} ranks)', log))

    # ── Step 2: node metric reconstruction ───────────────────────────────
    sections: list[dict] = []
    for sec in (family or {}).get("sections") or []:
        sections.append({"key": sec.get("key", ""), "title": sec.get("title", ""),
                         "rows": [_row(r.get("field", ""), r.get("value", ""),
                                       r.get("basis", "Engineering Estimate"),
                                       r.get("sources"), r.get("note", ""))
                                  for r in sec.get("rows") or []]})

    jedec = kb.get("jedec", {}).get(mem_type or "", {})
    retention = [_row(r["field"], r["value"], r.get("basis", "JEDEC"), r.get("sources"),
                      r.get("note", "")) for r in jedec.get("refresh") or []]
    if spd.get("trfc1_ns"):
        retention.append(_row("tRFC1 / tRFC2 / tRFCsb (this die)",
                              f'{spd["trfc1_ns"]} / {spd.get("trfc2_ns", "?")} / '
                              f'{spd.get("trfcsb_ns", "?")} ns (SPD bytes 42-47)', log))
    iface: list[dict] = []
    if spd.get("die_organization"):
        iface.append(_row("Monolithic die density & configuration",
                          f'{spd["die_organization"]}, {spd.get("package", "")}', log))
    if spd.get("max_data_rate"):
        iface.append(_row("Fastest JEDEC bin programmed in SPD",
                          f'{spd["max_data_rate"]} MT/s', log))
    if spd.get("cas_latencies"):
        iface.append(_row("Supported CAS latencies",
                          ", ".join(str(c) for c in spd["cas_latencies"]) + " (SPD bytes 24-28)", log))
    for key, title, extra in (
            ("transistor", "4. Transistor Architecture & Electrical Characteristics", retention),
            ("interface", "5. Interface & Application Alignment", iface)):
        if not extra:
            continue
        target = next((s for s in sections if s["key"] == key), None)
        if target is None:
            target = {"key": key, "title": title, "rows": []}
            sections.append(target)
        target["rows"].extend(extra)

    return {
        "part_number": part_number,
        "manufacturer": manufacturer,
        "vendor": vendor,
        "slots": sorted({str(d.get("slot")) for d in dimms if d.get("slot")}),
        "mpn_decode": dec,
        "die_rev": die_rev,
        "node_name": (die or {}).get("node", {}).get("value"),
        "family_name": (family or {}).get("name"),
        "step1": step1,
        "sections": sections,
        "disclaimer": kb.get("disclaimer", ""),
        "basis_legend": kb.get("basis_legend") or {},
    }


def dram_die_reports(platform_infos: list[dict] | None) -> list[dict[str, Any]]:
    """One die/process-node report per unique populated part number."""
    by_pn: dict[str, list[dict]] = {}
    for pi in platform_infos or []:
        for d in pi.get("dimms") or []:
            if d.get("populated") and d.get("part_number"):
                by_pn.setdefault(d["part_number"], []).append(d)
    if not by_pn:
        return []
    kb = load_dram_nodes()
    return [build_die_report(pn, dimms, kb) for pn, dimms in by_pn.items()]
