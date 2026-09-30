"""Branding, version and project registry shared by the MarginIQ CLI and GUI.

The project list lives in ``projects.json`` (repo root) so new platforms can be
added without touching code. Only the Python standard library is used here so
both ``rmt_log_pipeline.py`` and ``rmt_gui.py`` can import it cheaply.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

TOOL_NAME = "MarginIQ"
TOOL_SUBTITLE = "Intel CCG CVE DDR5 RMT Margin Analysis Tool"
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
        "WCL": {"name": "Wildcat Lake", "code": "WCL", "memory": "DDR5",
                "param_aliases": {"RxVref": "RxDqVrefByte"},
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
