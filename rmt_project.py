"""Branding, version and project registry shared by the MarginIQ CLI and GUI.

The project list lives in ``projects.json`` (repo root) so new platforms can be
added without touching code. Only the Python standard library is used here so
both ``rmt_log_pipeline.py`` and ``rmt_gui.py`` can import it cheaply.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

TOOL_NAME = "MarginIQ"
TOOL_SUBTITLE = "Intel CCG CVE Memory RMT Margin Analysis Tool"
TOOL_VERSION = "2.0.0"

HERE = Path(__file__).resolve().parent
PROJECTS_FILE = HERE / "projects.json"

# Header aliases understood for every project. Project entries in
# projects.json can add more.
BUILTIN_PARAM_ALIASES: dict[str, str] = {
    "RxVref": "RxDqVrefByte",
}

_FALLBACK_REGISTRY: dict[str, Any] = {
    "default": "NVL",
    "projects": {
        "NVL": {"name": "Nova Lake", "code": "NVL", "memory": "DDR5",
                "param_aliases": {}, "axis_config": "jmp_axis_settings.json"},
        "WCL": {"name": "Wildcat Lake", "code": "WCL", "memory": "DDR5 / LPDDR5",
                "param_aliases": {"RxVref": "RxDqVrefByte"},
                "rmt_fallback_markers": ["START_BYTE_RMT"],
                "require_rmt_task": True,
                "platform_variants": {"DDR5": "WCL DDR5 RVP", "LPDDR5": "WCL LPDDR5 RVP"},
                "dtr": {"layout": "rmt_task_runs",
                        "rmt_markers": ["START_RMT", "START_BYTE_RMT"],
                        "temps_from_filename": True},
                "axis_config": "jmp_axis_settings.json"},
    },
}


def load_projects(path: Path | None = None) -> dict[str, Any]:
    """Return the project registry (``{"default": key, "projects": {...}}``).

    Falls back to a built-in NVL/WCL registry when ``projects.json`` is
    missing or unreadable, printing the reason so misconfiguration is visible.
    """
    src = path or PROJECTS_FILE
    try:
        data = json.loads(src.read_text(encoding="utf-8"))
        projects = data.get("projects")
        if not isinstance(projects, dict) or not projects:
            raise ValueError("no 'projects' entries")
        default = data.get("default")
        if default not in projects:
            default = next(iter(projects))
        return {"default": default, "projects": projects}
    except FileNotFoundError:
        return dict(_FALLBACK_REGISTRY)
    except (PermissionError, ValueError, json.JSONDecodeError) as exc:
        print(f"Warning: could not read {src.name} ({exc}); using built-in project list.")
        return dict(_FALLBACK_REGISTRY)


def resolve_project_key(value: str | None, registry: dict[str, Any] | None = None) -> str | None:
    """Map a user value (key, code or name; case-insensitive) to a registry key.

    Returns the registry default for ``None``/empty input and ``None`` when
    the value does not match any project.
    """
    reg = registry or load_projects()
    if not value:
        return reg["default"]
    needle = value.strip().lower()
    for key, proj in reg["projects"].items():
        candidates = {key.lower(), str(proj.get("code", "")).lower(),
                      str(proj.get("name", "")).lower(), project_label(key, reg).lower()}
        if needle in candidates:
            return key
    return None


def get_project(key: str | None, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    reg = registry or load_projects()
    resolved = resolve_project_key(key, reg) or reg["default"]
    proj = dict(reg["projects"][resolved])
    proj.setdefault("code", resolved)
    proj.setdefault("name", resolved)
    proj["key"] = resolved
    return proj


def project_label(key: str, registry: dict[str, Any] | None = None) -> str:
    """Display label, e.g. ``"Nova Lake (NVL)"``."""
    reg = registry or load_projects()
    proj = reg["projects"].get(key, {})
    name = proj.get("name") or key
    code = proj.get("code") or key
    return name if name == code else f"{name} ({code})"


def project_param_aliases(key: str | None, registry: dict[str, Any] | None = None) -> dict[str, str]:
    aliases = dict(BUILTIN_PARAM_ALIASES)
    aliases.update(get_project(key, registry).get("param_aliases") or {})
    return aliases


def project_rmt_fallback_markers(key: str | None,
                                 registry: dict[str, Any] | None = None) -> list[str]:
    """Block start markers parsed only when a log has no ``START_RMT`` data.

    E.g. WCL ``["START_BYTE_RMT"]``; only blocks inside a "Rank Margin
    Tool" task run are used.
    """
    markers = get_project(key, registry).get("rmt_fallback_markers") or []
    return [str(m).strip().upper() for m in markers if str(m).strip()]


def project_require_rmt_task(key: str | None, registry: dict[str, Any] | None = None) -> bool:
    """True = RMT data only from inside a "Rank Margin Tool" task (WCL)."""
    return bool(get_project(key, registry).get("require_rmt_task", False))


def project_rmt_limits(key: str | None, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    """Spec RMT limits ``{family: {"frequencies": {freq: {"ticks": .., "physical": ..}}}}``.

    ``{}`` when the project has none (e.g. NVL): the axis-config +/-Ref is used.
    """
    limits = get_project(key, registry).get("rmt_limits")
    return dict(limits) if isinstance(limits, dict) else {}


def project_dtr_rules(key: str | None, registry: dict[str, Any] | None = None) -> dict[str, Any]:
    """Project-specific DTR (thermal) parsing rules; ``{}`` = built-in NVL flow.

    ``layout: "rmt_task_runs"`` - Boot RMT is the first "Rank Margin Tool"
    run and Run RMT the last one (no "Temp Drift RMT Test" task / "Hit
    Enter" pause). ``rmt_markers`` lists accepted block start markers and
    ``temps_from_filename`` takes Boot/Run temps from BCRH / BHRC in the name.
    """
    rules = get_project(key, registry).get("dtr")
    return dict(rules) if isinstance(rules, dict) else {}


def project_platform_variants(key: str | None,
                              registry: dict[str, Any] | None = None) -> dict[str, str]:
    """Memory family -> platform variant label, e.g. ``{"LPDDR5": "WCL LPDDR5 RVP"}``."""
    variants = get_project(key, registry).get("platform_variants")
    return {str(k).upper(): str(v) for k, v in variants.items()} if isinstance(variants, dict) else {}


def resolve_platform_variant(variants: dict[str, str], memory_tech: str | None,
                             board: str | None = None) -> str | None:
    """Pick the variant for the memory decoded from the log (e.g. "LPDDR5X").

    Tries the exact technology, then its family ("LPDDR5X" -> "LPDDR5"), then
    the board string ("WCL LP5 RVP2" -> LPDDR5). None when nothing matches.
    """
    if not variants:
        return None
    candidates: list[str] = []
    for tech in (memory_tech or "").upper().replace("/", " ").split():
        candidates.append(tech)
        fam = re.match(r"(LP)?DDR\d", tech)
        if fam:
            candidates.append(fam.group(0))
    board_u = (board or "").upper()
    if re.search(r"\bLP(?:DDR)?5X?\b", board_u):
        candidates.append("LPDDR5")
    elif re.search(r"\bDDR5\b", board_u):
        candidates.append("DDR5")
    for c in candidates:
        if c in variants:
            return variants[c]
    return None
