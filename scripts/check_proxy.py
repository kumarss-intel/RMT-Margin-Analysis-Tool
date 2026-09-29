"""Pre-flight check for proxy settings that would make pip fail.

Run by setup.bat with the .venv interpreter before any pip install. pip aborts
with "OSError: Failed to parse: <url>" when a proxy value is malformed (for
example a template placeholder such as ``http://proxy-server:port``). This
script finds every proxy value pip will use, reports where each one comes
from, and exits non-zero with remediation steps when one cannot be parsed.

Exit codes: 0 = all proxy values usable (or none set), 1 = invalid value found.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from pip._vendor.urllib3.exceptions import LocationParseError
from pip._vendor.urllib3.util import parse_url

ENV_PROXY_VARS = ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY")


def _is_valid(url: str) -> bool:
    try:
        parsed = parse_url(url)
    except LocationParseError:
        return False
    return bool(parsed.host)


def _pip_config_proxies() -> list[tuple[str, str]]:
    """Return (source, value) for the proxy pip will actually use from its config.

    pip applies config in this order, later winning: [global] section,
    [install] section, then the PIP_PROXY environment variable.
    """
    found: list[tuple[str, str]] = []
    try:
        from pip._internal.configuration import Configuration

        config = Configuration(isolated=False)
        config.load()
        files_with_proxy = []
        for _variant, files in config.iter_config_files():
            for name in files:
                path = Path(name)
                try:
                    text = path.read_text(encoding="utf-8", errors="replace")
                except (FileNotFoundError, PermissionError, OSError):
                    continue
                if any(line.strip().lower().startswith("proxy") for line in text.splitlines()):
                    files_with_proxy.append(str(path))
        for key in (":env:.proxy", "install.proxy", "global.proxy"):
            try:
                value = config.get_value(key)
            except Exception:
                continue
            if not value:
                continue
            if key.startswith(":env:"):
                source = "environment variable PIP_PROXY"
            elif files_with_proxy:
                source = f"pip config file {', '.join(files_with_proxy)} (key '{key}')"
            else:
                source = f"pip config key '{key}'"
            found.append((source, value))
            break
    except Exception as err:  # pip internals changed - never block setup on this
        print(f"      [WARN] Could not inspect pip configuration: {err}")
    return found


def main() -> int:
    settings: list[tuple[str, str]] = _pip_config_proxies()
    for var in ENV_PROXY_VARS:
        value = os.environ.get(var) or os.environ.get(var.lower())
        if value:
            settings.append((f"environment variable {var}", value))

    bad = [(source, value) for source, value in settings if not _is_valid(value)]
    for source, value in settings:
        status = "INVALID" if (source, value) in bad else "ok"
        print(f"      proxy [{status}] {value}  <- {source}")
    if not settings:
        print("      proxy: none configured (direct connection)")
    if not bad:
        return 0

    print()
    print("[ERROR] pip cannot use the proxy setting(s) marked INVALID above.")
    print("        A value like http://proxy-server:port is a template placeholder;")
    print("        the port must be a number. Fix it in one of these ways:")
    print()
    print("        1) Quick: run setup with a valid proxy for this run only")
    print("             setup.bat --proxy http://proxy-chain.intel.com:912")
    print()
    print("        2) Permanent: correct or remove the setting at its source")
    if any("pip config" in s for s, _ in bad):
        print("             .venv\\Scripts\\python.exe -m pip config debug     (shows the file)")
        print("             .venv\\Scripts\\python.exe -m pip config set global.proxy http://proxy-chain.intel.com:912")
        print("               or delete the 'proxy =' line from that pip.ini")
    if any("environment variable" in s for s, _ in bad):
        print("             Windows Settings > 'Edit environment variables for your account':")
        print("               fix or delete the variable(s) listed above, then open a NEW terminal")
    return 1


if __name__ == "__main__":
    sys.exit(main())
