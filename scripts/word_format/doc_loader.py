"""Read-only loader for DOCX, DOTX, and legacy DOC/DOT/WPS files.

Legacy binary formats (.doc, .dot, .wps) are converted from an isolated
temporary copy into a SHA-256-addressed temporary DOCX cache so user authority
files are never modified, locked, or touched in place.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path

from config import DEFAULT_RUNTIME_ROOT, MAX_DOCUMENT_BYTES
from docx import Document
from runtime_detection import detect_renderers

from .contracts import FormatError, sha256

OOXML_TEMPLATE_SUFFIXES = frozenset({".docx", ".dotx"})
LEGACY_WORD_SUFFIXES = frozenset({".doc", ".dot", ".wps"})
SUPPORTED_WORD_SUFFIXES = OOXML_TEMPLATE_SUFFIXES | LEGACY_WORD_SUFFIXES
_LEGACY_CACHE_DIR = Path(tempfile.gettempdir()) / "word-format-master-doc-cache"

_DOTX_CONTENT_TYPE = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
)
_DOCX_CONTENT_TYPE = (
    b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)


def _is_valid_ooxml_docx(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size == 0 or not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as zf:
            return "word/document.xml" in zf.namelist()
    except zipfile.BadZipFile:
        return False


def _rewrite_dotx_bytes(path: Path) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in source.infolist():
            blob = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                blob = blob.replace(_DOTX_CONTENT_TYPE, _DOCX_CONTENT_TYPE)
            dst.writestr(item, blob)
    return buffer.getvalue()


def _word_com_to_docx(staged_source: Path, output_docx: Path) -> None:
    shell = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        raise RuntimeError("PowerShell is unavailable for Word/WPS COM conversion")
    # wdFormatXMLDocument = 12 forces Transitional DOCX regardless of Word's default save setting.
    script = r"""
param([string]$InputPath, [string]$OutputPath)
$app = $null
$doc = $null
$ErrorActionPreference = 'Stop'
try {
  foreach ($progId in @('Word.Application', 'KWPS.Application', 'WPS.Application')) {
    try {
      $app = New-Object -ComObject $progId
      break
    } catch {}
  }
  if ($null -eq $app) {
    throw 'No Word/WPS COM server available'
  }
  $app.Visible = $false
  try { $app.DisplayAlerts = 0 } catch {}
  $doc = $app.Documents.Open($InputPath, $false, $true)
  try {
    $doc.SaveAs2([ref]$OutputPath, [ref]12)
  } catch {
    $doc.SaveAs($OutputPath, 12)
  }
} finally {
  if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
  if ($null -ne $app) { try { $app.Quit() } catch {} }
  if ($null -ne $doc) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
  if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
""".strip()
    with tempfile.TemporaryDirectory(prefix="wfm-doc-com-") as temp_dir:
        script_path = Path(temp_dir) / "convert_to_docx.ps1"
        script_path.write_text(script + "\n", encoding="utf-8")
        result = subprocess.run(
            [
                shell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(staged_source),
                str(output_docx),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
    if result.returncode != 0 or not _is_valid_ooxml_docx(output_docx):
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Word/WPS COM conversion failed: {detail[-500:] or 'no DOCX output'}")


def _libreoffice_to_docx(staged_source: Path, output_docx: Path, executable: Path) -> None:
    with (
        tempfile.TemporaryDirectory(prefix="wfm-lo-doc-profile-") as profile_dir,
        tempfile.TemporaryDirectory(prefix="wfm-lo-doc-out-") as out_dir,
    ):
        profile_uri = Path(profile_dir).resolve().as_uri()
        env = os.environ.copy()
        env["HOME"] = profile_dir
        result = subprocess.run(
            [
                str(executable),
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--invisible",
                "--norestore",
                "--convert-to",
                "docx",
                "--outdir",
                out_dir,
                str(staged_source),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
            env=env,
        )
        generated = Path(out_dir) / f"{staged_source.stem}.docx"
        if result.returncode != 0 or not _is_valid_ooxml_docx(generated):
            detail = (result.stderr or result.stdout).strip()
            raise RuntimeError(
                f"LibreOffice conversion failed: {detail[-500:] or 'no DOCX output'}"
            )
        shutil.copy2(generated, output_docx)


def _convert_legacy_word_file(staged_source: Path, output_docx: Path) -> None:
    """Convert a staged read-only legacy Word file (.doc/.dot/.wps) to output_docx."""
    runtime = detect_renderers(DEFAULT_RUNTIME_ROOT)
    errors: list[str] = []
    if os.name == "nt" and runtime.get("word", {}).get("available"):
        try:
            _word_com_to_docx(staged_source, output_docx)
            return
        except Exception as exc:  # noqa: BLE001
            errors.append(str(exc))
    lo_entry = runtime.get("libreoffice") or {}
    if lo_entry.get("available") and lo_entry.get("path"):
        try:
            _libreoffice_to_docx(staged_source, output_docx, Path(lo_entry["path"]))
            return
        except Exception as exc:  # noqa: BLE001
            errors.append(str(exc))
    if os.name == "nt" and not runtime.get("word", {}).get("available"):
        # Attempt WPS ProgIDs even when Word.Application CLSID was not registered.
        try:
            _word_com_to_docx(staged_source, output_docx)
            return
        except Exception as exc:  # noqa: BLE001
            errors.append(str(exc))
    detail = "; ".join(errors) if errors else "Microsoft Word, WPS Office, or LibreOffice not found"
    raise FormatError(
        f"Cannot read legacy {staged_source.suffix.upper()} file without a local converter ({detail})",
        "unsupported",
    )


def ensure_readable_docx_path(path: str | Path) -> Path:
    """Return a readable OOXML package path (.docx/.dotx) without modifying ``path``."""
    source = Path(path).resolve()
    if not source.is_file():
        raise FormatError(f"Word file does not exist: {source}")
    if source.stat().st_size > MAX_DOCUMENT_BYTES:
        raise FormatError("Word file exceeds the maximum document size limit")
    suffix = source.suffix.lower()
    if suffix in OOXML_TEMPLATE_SUFFIXES:
        return source
    if suffix not in LEGACY_WORD_SUFFIXES:
        raise FormatError(
            f"Unsupported Word file extension {suffix!r}; expected DOCX, DOTX, DOC, DOT, or WPS",
            "unsupported",
        )
    original_hash = sha256(source)
    cache_dir = _LEGACY_CACHE_DIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    cached_docx = cache_dir / f"{original_hash}.docx"
    if _is_valid_ooxml_docx(cached_docx):
        if sha256(source) != original_hash:
            raise FormatError("Source file changed during legacy conversion check")
        return cached_docx

    # Handle OOXML packages that were saved or renamed with a .doc/.dot/.wps suffix.
    if _is_valid_ooxml_docx(source):
        tmp_path = cache_dir / f".{original_hash}.{uuid.uuid4().hex}.docx"
        try:
            tmp_path.write_bytes(_rewrite_dotx_bytes(source))
            os.replace(tmp_path, cached_docx)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()
        if sha256(source) != original_hash:
            raise FormatError("Source file changed during legacy package read")
        return cached_docx

    with tempfile.TemporaryDirectory(prefix="wfm-doc-stage-") as stage_dir:
        staged = Path(stage_dir) / f"input{suffix}"
        shutil.copy2(source, staged)
        try:
            staged.chmod(0o444)
        except OSError:
            pass
        tmp_out = Path(stage_dir) / "converted.docx"
        _convert_legacy_word_file(staged, tmp_out)
        if not _is_valid_ooxml_docx(tmp_out):
            raise FormatError("Converter did not produce a valid DOCX package", "unsupported")
        publish_tmp = cache_dir / f".{original_hash}.{uuid.uuid4().hex}.docx"
        try:
            publish_tmp.write_bytes(_rewrite_dotx_bytes(tmp_out))
            os.replace(publish_tmp, cached_docx)
        finally:
            if publish_tmp.exists():
                publish_tmp.unlink()
        try:
            staged.chmod(0o666)
        except OSError:
            pass

    if sha256(source) != original_hash:
        raise FormatError("Legacy source file changed during read-only conversion")
    return cached_docx


def open_word_document(path: str | Path) -> Document:
    """Open DOCX, DOTX, DOC, DOT, or WPS read-only as a ``python-docx`` Document."""
    readable = ensure_readable_docx_path(path)
    if readable.suffix.lower() == ".docx":
        return Document(readable)
    buffer = io.BytesIO(_rewrite_dotx_bytes(readable))
    return Document(buffer)
