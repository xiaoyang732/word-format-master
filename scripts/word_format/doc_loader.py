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
_MAX_CACHE_FILES = 64
_MAX_CACHE_BYTES = 128 * 1024 * 1024

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


def _prune_legacy_cache(cache_dir: Path, *, keep: Path | None = None) -> None:
    """Bound the temporary converted-DOCX cache by file count and total bytes."""
    try:
        entries = []
        for candidate in cache_dir.glob("*.docx"):
            if candidate.name.startswith(".") or (keep is not None and candidate == keep):
                continue
            try:
                stat = candidate.stat()
                entries.append((stat.st_mtime, stat.st_size, candidate))
            except OSError:
                continue
        entries.sort(key=lambda item: item[0], reverse=True)
        total_bytes = sum(item[1] for item in entries)
        while entries and (len(entries) > _MAX_CACHE_FILES or total_bytes > _MAX_CACHE_BYTES):
            _, size, oldest = entries.pop()
            total_bytes -= size
            try:
                oldest.unlink(missing_ok=True)
            except OSError:
                pass
    except OSError:
        pass


def _word_com_to_docx(staged_source: Path, output_docx: Path) -> None:
    from office_automation import automate_word

    automate_word(staged_source, output_docx, mode="convert", timeout=120)
    if not _is_valid_ooxml_docx(output_docx):
        raise RuntimeError("Microsoft Word did not produce a readable DOCX")


def _wps_com_to_docx(staged_source: Path, output_docx: Path) -> None:
    from office_automation import automate_wps

    automate_wps(staged_source, output_docx, mode="convert", timeout=120)
    if not _is_valid_ooxml_docx(output_docx):
        raise RuntimeError("WPS Office did not produce a readable DOCX")


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
    suffix = staged_source.suffix.lower()
    com_order = ("wps", "word") if suffix == ".wps" else ("word", "wps")
    if os.name == "nt":
        for backend in com_order:
            if not runtime.get(backend, {}).get("available"):
                continue
            converter = _wps_com_to_docx if backend == "wps" else _word_com_to_docx
            try:
                converter(staged_source, output_docx)
                return
            except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as exc:
                errors.append(str(exc))
    lo_entry = runtime.get("libreoffice") or {}
    if lo_entry.get("available") and lo_entry.get("path"):
        try:
            _libreoffice_to_docx(staged_source, output_docx, Path(lo_entry["path"]))
            return
        except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as exc:
            errors.append(str(exc))
    detail = (
        "; ".join(errors)
        if errors
        else "Microsoft Word, WPS Office, or LibreOffice not found"
    )
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
    _prune_legacy_cache(cache_dir, keep=cached_docx)
    return cached_docx


def open_word_document(path: str | Path) -> Document:
    """Open DOCX, DOTX, DOC, DOT, or WPS read-only as a ``python-docx`` Document."""
    readable = ensure_readable_docx_path(path)
    if readable.suffix.lower() == ".docx":
        return Document(readable)
    buffer = io.BytesIO(_rewrite_dotx_bytes(readable))
    return Document(buffer)
