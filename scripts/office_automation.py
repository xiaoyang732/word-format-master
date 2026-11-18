"""Isolated Microsoft Word and WPS Office automation for render, refresh, and conversion."""

from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from runtime_detection import _word_application_com_server

AUTOMATION_MODES = frozenset({"inspect", "render", "refresh", "convert"})


def _find_powershell() -> str:
    shell = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        raise RuntimeError("Office/WPS COM automation requires PowerShell")
    return shell


def _stop_owned_processes(evidence: Path, shell: str) -> None:
    """A timed out worker may only terminate instances it proved it created."""
    identity_path = evidence.with_suffix(".identity.json")
    if not identity_path.is_file():
        return
    identity = json.loads(identity_path.read_text(encoding="utf-8-sig"))
    if not identity.get("owned"):
        return
    script = r"""param([string]$IdentityPath)
$item = Get-Content -LiteralPath $IdentityPath -Raw | ConvertFrom-Json
$targets = @()
if ($null -ne $item.owned_items -and $item.owned_items.Count -gt 0) {
  $targets = @($item.owned_items)
} elseif ($null -ne $item.pid -and $null -ne $item.started_ticks) {
  $targets = @(@{ pid = $item.pid; executable = $item.executable; started_ticks = $item.started_ticks })
}
foreach ($target in $targets) {
  $ownedProcess = Get-Process -Id $target.pid -ErrorAction SilentlyContinue
  if ($null -ne $ownedProcess -and $ownedProcess.Path -eq $target.executable -and
      $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() -eq $target.started_ticks) {
    Stop-Process -Id $ownedProcess.Id -Force -ErrorAction SilentlyContinue
  }
}
"""
    cleanup = evidence.with_suffix(".cleanup.ps1")
    cleanup.write_text(script, encoding="utf-8-sig")
    subprocess.run(
        [shell, "-NoProfile", "-NonInteractive", "-File", str(cleanup), str(identity_path)],
        capture_output=True,
        timeout=15,
        check=False,
    )


def _run_automation_worker(
    backend: str,
    source: str | Path,
    destination: str | Path | None = None,
    *,
    mode: str = "inspect",
    timeout: int = 180,
) -> dict[str, Any]:
    if mode not in AUTOMATION_MODES:
        raise ValueError(f"Unknown {backend} automation mode: {mode}")
    shell = _find_powershell()
    label = "Microsoft Word" if backend == "word" else "WPS Office"
    script = Path(__file__).with_suffix(".ps1")
    with tempfile.TemporaryDirectory(prefix=f"wfm-{backend}-") as folder:
        evidence = Path(folder) / "result.json"
        command = [
            shell,
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-InputPath",
            str(Path(source).resolve()),
            "-Mode",
            mode,
            "-Backend",
            backend,
            "-EvidencePath",
            str(evidence),
        ]
        if destination is not None:
            command += ["-OutputPath", str(Path(destination).resolve())]
        worker = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        try:
            stdout, stderr = worker.communicate(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            worker.kill()
            worker.communicate()
            _stop_owned_processes(evidence, shell)
            raise RuntimeError(f"{label} {mode} timed out after {timeout}s") from exc
        if worker.returncode or not evidence.is_file():
            detail = (stderr or stdout).decode("utf-8", errors="replace")[-1600:]
            raise RuntimeError(f"{label} {mode} failed: {detail}")
        return json.loads(evidence.read_text(encoding="utf-8-sig"))


def automate_word(
    source: str | Path,
    destination: str | Path | None = None,
    *,
    mode: str = "inspect",
    timeout: int = 180,
) -> dict[str, Any]:
    """Return verified Microsoft Word identity and computed semantics, or fail explicitly."""
    if mode not in AUTOMATION_MODES:
        raise ValueError("Unknown Microsoft Word automation mode")
    registered = _word_application_com_server()
    if registered is not None and registered.name.casefold() != "winword.exe":
        raise RuntimeError(
            f"Word.Application is registered to a different application: {registered}"
        )
    result = _run_automation_worker("word", source, destination, mode=mode, timeout=timeout)
    identity = result.get("application", {})
    if (
        identity.get("name") != "Microsoft Word"
        or Path(identity.get("executable", "")).name.casefold() != "winword.exe"
    ):
        raise RuntimeError("Actual automation application is not Microsoft Word")
    identity["registered_server"] = str(registered) if registered else None
    return result


def automate_wps(
    source: str | Path,
    destination: str | Path | None = None,
    *,
    mode: str = "convert",
    timeout: int = 180,
) -> dict[str, Any]:
    """Run WPS Office COM automation without claiming Microsoft Word identity."""
    result = _run_automation_worker("wps", source, destination, mode=mode, timeout=timeout)
    identity = result.get("application", {})
    if identity.get("name") != "WPS Office":
        raise RuntimeError("Actual automation application is not WPS Office")
    return result
