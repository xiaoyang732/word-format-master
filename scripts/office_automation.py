"""Isolated Microsoft Word automation shared by render, refresh and legacy loading."""

from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

from runtime_detection import _word_application_com_server


def _stop_owned_word(evidence: Path, shell: str) -> None:
    """A timed out worker may only terminate the instance it proved it created."""
    identity_path = evidence.with_suffix(".identity.json")
    if not identity_path.is_file():
        return
    identity = json.loads(identity_path.read_text(encoding="utf-8-sig"))
    if not identity.get("owned"):
        return
    script = r"""param([string]$IdentityPath)
$item = Get-Content -LiteralPath $IdentityPath -Raw | ConvertFrom-Json
$ownedProcess = Get-Process -Id $item.pid -ErrorAction SilentlyContinue
if ($null -ne $ownedProcess -and $ownedProcess.Path -eq $item.executable -and
    $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() -eq $item.started_ticks) {
  Stop-Process -Id $ownedProcess.Id -Force
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


def automate_word(source, destination=None, *, mode="inspect", timeout=180):
    """Return actual application identity and computed semantics, or fail explicitly."""
    if mode not in {"inspect", "render", "refresh", "convert"}:
        raise ValueError("Unknown Microsoft Word automation mode")
    shell = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        raise RuntimeError("Microsoft Word automation requires PowerShell")
    registered = _word_application_com_server()
    if registered is not None and registered.name.casefold() != "winword.exe":
        raise RuntimeError(
            f"Word.Application is registered to a different application: {registered}"
        )
    script = Path(__file__).with_suffix(".ps1")
    with tempfile.TemporaryDirectory(prefix="wfm-office-") as folder:
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
            _stop_owned_word(evidence, shell)
            raise RuntimeError(f"Microsoft Word {mode} timed out after {timeout}s") from exc
        if worker.returncode or not evidence.is_file():
            detail = (stderr or stdout).decode("utf-8", errors="replace")[-1600:]
            raise RuntimeError(f"Microsoft Word {mode} failed: {detail}")
        result = json.loads(evidence.read_text(encoding="utf-8-sig"))
    identity = result.get("application", {})
    if (
        identity.get("name") != "Microsoft Word"
        or Path(identity.get("executable", "")).name.casefold() != "winword.exe"
    ):
        raise RuntimeError("Actual automation application is not Microsoft Word")
    identity["registered_server"] = str(registered) if registered else None
    return result
