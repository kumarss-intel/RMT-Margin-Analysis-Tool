"""Rate-limited "is a newer MarginIQ available?" check for git clones.

How it stays light on the server:

* At most one network call per ``MARGINIQ_UPDATE_CHECK_HOURS`` (default 24 h)
  per clone. The result is cached in ``.git/marginiq-update-check.json`` and
  re-used in between, so launching the GUI or running the CLI many times a day
  costs nothing. A failed check (offline, proxy, no credentials) is cached too,
  so it is not retried on every launch.
* The network call is a single ``git ls-remote`` of the upstream branch: it
  reads one ref, downloads no objects and never modifies the working tree,
  index or remote-tracking branches.
* Git is told never to prompt (``GIT_TERMINAL_PROMPT=0``,
  ``credential.interactive=never``) and the call has a short timeout, so the
  check can never hang the GUI, the CLI or an agent run.

Set ``MARGINIQ_NO_UPDATE_CHECK=1`` to disable automatic checks entirely
(an explicit check from *Help > Check for Updates* still works).
Only the Python standard library is used.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent
DISABLE_ENV = "MARGINIQ_NO_UPDATE_CHECK"
INTERVAL_ENV = "MARGINIQ_UPDATE_CHECK_HOURS"
DEFAULT_INTERVAL_HOURS = 24.0
STATE_FILE_NAME = "marginiq-update-check.json"
LOCAL_TIMEOUT_S = 5
NETWORK_TIMEOUT_S = 10

STATE_UPDATE = "update"
STATE_CURRENT = "current"
STATE_UNKNOWN = "unknown"
STATE_DISABLED = "disabled"


@dataclass
class UpdateStatus:
    state: str
    detail: str = ""
    remote_ref: str = ""
    local_sha: str = ""
    remote_sha: str = ""
    checked_at: float = 0.0
    from_cache: bool = False

    @property
    def update_available(self) -> bool:
        return self.state == STATE_UPDATE

    def headline(self) -> str:
        if self.update_available:
            return f"A newer MarginIQ version is available on {self.remote_ref}."
        if self.state == STATE_CURRENT:
            return f"MarginIQ is up to date with {self.remote_ref}."
        if self.state == STATE_DISABLED:
            return f"Automatic update checks are disabled ({DISABLE_ENV} is set)."
        return f"Could not check for updates: {self.detail}"

    def message(self) -> str:
        """Multi-line, user-facing text including the update instructions."""
        lines = [self.headline()]
        if self.update_available:
            lines += [
                "",
                "To update, close MarginIQ and run in the tool folder:",
                f"    cd \"{REPO_DIR}\"",
                "    git pull",
                "",
                "If requirements.txt changed, also run:  setup.bat --no-prompt",
            ]
        if self.checked_at:
            stamp = time.strftime("%Y-%m-%d %H:%M", time.localtime(self.checked_at))
            lines += ["", f"Last checked with the server: {stamp}"]
        return "\n".join(lines)


def automatic_checks_disabled() -> bool:
    return os.environ.get(DISABLE_ENV, "").strip().lower() not in ("", "0", "false", "no")


def check_interval_seconds() -> float:
    raw = os.environ.get(INTERVAL_ENV, "").strip()
    try:
        hours = float(raw) if raw else DEFAULT_INTERVAL_HOURS
    except ValueError:
        hours = DEFAULT_INTERVAL_HOURS
    return max(0.0, hours) * 3600.0


def _git_env() -> dict[str, str]:
    env = dict(os.environ)
    env["GIT_TERMINAL_PROMPT"] = "0"
    env["GCM_INTERACTIVE"] = "never"
    env["LC_ALL"] = "C"
    return env


def _git(args: list[str], timeout: float) -> tuple[int | None, str]:
    """Run git in the repo. Returns (returncode, stdout); returncode is None
    when git could not be run or timed out."""
    git = shutil.which("git")
    if not git:
        return None, ""
    kwargs: dict = {}
    if sys.platform == "win32":
        # No console window flashes when called from the pythonw GUI.
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        proc = subprocess.run(
            [git, "-c", "credential.interactive=never", *args],
            cwd=str(REPO_DIR), env=_git_env(), stdin=subprocess.DEVNULL,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout, check=False, **kwargs)
    except (OSError, subprocess.SubprocessError):
        return None, ""
    return proc.returncode, proc.stdout.strip()


def _state_path() -> Path | None:
    rc, out = _git(["rev-parse", "--git-path", STATE_FILE_NAME], LOCAL_TIMEOUT_S)
    if rc != 0 or not out:
        return None
    path = Path(out)
    return path if path.is_absolute() else REPO_DIR / path


def _load_state(path: Path | None) -> dict:
    if path is None:
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, PermissionError, OSError, ValueError):
        return {}


def _save_state(path: Path | None, data: dict) -> None:
    if path is None:
        return
    tmp = path.with_name(path.name + ".tmp")
    try:
        tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
        os.replace(tmp, path)
    except (PermissionError, OSError):
        pass


def _upstream() -> tuple[str, str, str] | str:
    """Return (remote, merge_ref, display) for the current branch, or an
    error string when it cannot be determined."""
    rc, branch = _git(["symbolic-ref", "--quiet", "--short", "HEAD"], LOCAL_TIMEOUT_S)
    if rc != 0 or not branch:
        return "detached HEAD (no branch checked out)"
    _, remote = _git(["config", "--get", f"branch.{branch}.remote"], LOCAL_TIMEOUT_S)
    _, merge = _git(["config", "--get", f"branch.{branch}.merge"], LOCAL_TIMEOUT_S)
    if not remote or not merge or remote == ".":
        return f"branch '{branch}' has no remote upstream"
    short = merge[len("refs/heads/"):] if merge.startswith("refs/heads/") else merge
    return remote, merge, f"{remote}/{short}"


def check_for_updates(force: bool = False) -> UpdateStatus:
    """Compare the local HEAD with the upstream branch on the server.

    Hits the network only when the cached result is older than the check
    interval, or when *force* is true (explicit user request).
    """
    if not force and automatic_checks_disabled():
        return UpdateStatus(STATE_DISABLED)
    if not (REPO_DIR / ".git").exists():
        return UpdateStatus(STATE_UNKNOWN, "this copy is not a git clone (re-download it with git clone)")
    rc, head = _git(["rev-parse", "HEAD"], LOCAL_TIMEOUT_S)
    if rc is None:
        return UpdateStatus(STATE_UNKNOWN, "git is not installed or not on PATH")
    if rc != 0 or not head:
        return UpdateStatus(STATE_UNKNOWN, "could not read the local commit")
    upstream = _upstream()
    if isinstance(upstream, str):
        return UpdateStatus(STATE_UNKNOWN, upstream, local_sha=head)
    remote, merge, display = upstream

    state_path = _state_path()
    state = _load_state(state_path)
    now = time.time()
    checked_at = float(state.get("checked_at") or 0.0)
    same_target = state.get("remote") == remote and state.get("merge") == merge
    fresh = same_target and 0.0 <= now - checked_at < check_interval_seconds()

    if fresh and not force:
        remote_sha = str(state.get("remote_sha") or "")
        from_cache = True
        failure = str(state.get("error") or "")
    else:
        rc, out = _git(["ls-remote", "--exit-code", remote, merge], NETWORK_TIMEOUT_S)
        remote_sha = out.split()[0] if rc == 0 and out else ""
        if rc is None:
            failure = "the server did not answer in time (offline or proxy issue?)"
        elif rc == 2:
            failure = f"branch {display} no longer exists on the server"
        elif rc != 0 or not remote_sha:
            failure = "could not reach the server (network, proxy or access rights)"
        else:
            failure = ""
        checked_at, from_cache = now, False
        state.update({"remote": remote, "merge": merge, "checked_at": now,
                      "remote_sha": remote_sha, "error": failure})
        _save_state(state_path, state)

    base = dict(remote_ref=display, local_sha=head, checked_at=checked_at, from_cache=from_cache)
    if not remote_sha:
        return UpdateStatus(STATE_UNKNOWN, failure or "no cached result", remote_sha="", **base)
    if remote_sha == head:
        return UpdateStatus(STATE_CURRENT, remote_sha=remote_sha, **base)
    # 0: server commit is already contained in HEAD (local is ahead).
    # 1: commit known locally but not merged; 128: commit never fetched.
    rc, _ = _git(["merge-base", "--is-ancestor", remote_sha, "HEAD"], LOCAL_TIMEOUT_S)
    if rc == 0:
        return UpdateStatus(STATE_CURRENT, "local branch is ahead of the server",
                            remote_sha=remote_sha, **base)
    return UpdateStatus(STATE_UPDATE, remote_sha=remote_sha, **base)


def should_popup(status: UpdateStatus) -> bool:
    """True the first time a given server commit is reported, so the GUI pops
    a dialog once per new version instead of on every launch."""
    if not status.update_available:
        return False
    return _load_state(_state_path()).get("notified_sha") != status.remote_sha


def mark_notified(status: UpdateStatus) -> None:
    path = _state_path()
    state = _load_state(path)
    state["notified_sha"] = status.remote_sha
    _save_state(path, state)


def cli_notice(status: UpdateStatus) -> str:
    """One-line notice for console output (empty when nothing to report)."""
    if not status.update_available:
        return ""
    return (f"[MarginIQ update] A newer version is available on {status.remote_ref}. "
            f"Run 'git pull' in {REPO_DIR} to update "
            f"(then 'setup.bat --no-prompt' if requirements.txt changed).")


if __name__ == "__main__":
    result = check_for_updates(force="--force" in sys.argv)
    print(result.message())
    raise SystemExit(0)
