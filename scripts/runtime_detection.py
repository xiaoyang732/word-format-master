#!/usr/bin/env python3
"""Detect and optionally provision local Word rendering runtimes."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from config import (
    DOWNLOAD_TIMEOUT,
    HTTP_REQUEST_TIMEOUT,
    LIBREOFFICE_DOWNLOAD_HOST,
    LIBREOFFICE_DOWNLOAD_PAGE,
)

OFFICIAL_DOWNLOAD_PAGE = LIBREOFFICE_DOWNLOAD_PAGE
DOWNLOAD_HOST = LIBREOFFICE_DOWNLOAD_HOST
DOWNLOAD_PATTERN = re.compile(
    r"https://download\.documentfoundation\.org/libreoffice/stable/[^\"'\s<>]+/"
    r"LibreOffice_[^\"'\s<>]+_Win_x86-64\.msi",
    re.IGNORECASE,
)
RuntimeProgress = Callable[[dict[str, Any]], None]


def _existing_file(candidates: list[Path]) -> Path | None:
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    return None


def _registry_candidates(product: str) -> list[Path]:
    if os.name != "nt":
        return []
    try:
        import winreg
    except ImportError:
        return []
    subkeys = {
        "word": [
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        ],
        "libreoffice": [r"SOFTWARE\LibreOffice\UNO\InstallPath"],
    }[product]
    result: list[Path] = []
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0)):
            for subkey in subkeys:
                try:
                    with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | view) as key:
                        for value_name in ("InstallPath", "InstallationPath", "ClientFolder"):
                            try:
                                value, _ = winreg.QueryValueEx(key, value_name)
                            except OSError:
                                continue
                            if value:
                                base = Path(str(value))
                                result.append(base / ("WINWORD.EXE" if product == "word" else Path("program") / "soffice.exe"))
                except OSError:
                    continue
    return result


def _find_word(runtime_root: Path) -> Path | None:
    candidates = [
        Path(item) for item in (
            shutil.which("WINWORD.EXE"),
            shutil.which("winword"),
        ) if item
    ]
    if os.name == "nt":
        program_files = [Path(os.environ.get("ProgramFiles", r"C:\Program Files"))]
        program_files.append(Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")))
        for root in program_files:
            candidates.extend([
                root / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE",
                root / "Microsoft Office" / "Office16" / "WINWORD.EXE",
            ])
    candidates.extend(_registry_candidates("word"))
    return _existing_file(candidates)


def _word_application_com_server() -> Path | None:
    """Resolve the executable registered for the Windows Word.Application COM class.

    WPS Office registers this class for Word-compatible automation on some
    Windows systems. Its ExportAsFixedFormat implementation can render DOCX
    files even when WINWORD.EXE is not installed.
    """
    if os.name != "nt":
        return None
    try:
        import winreg
    except ImportError:
        return None
    clsid = None
    for root, subkey in (
        (winreg.HKEY_CURRENT_USER, r"Software\Classes\Word.Application\CLSID"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Classes\Word.Application\CLSID"),
        (winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"),
    ):
        try:
            with winreg.OpenKey(root, subkey) as key:
                value, _ = winreg.QueryValueEx(key, None)
                if value:
                    clsid = str(value)
                    break
        except OSError:
            continue
    if not clsid:
        return None
    for root, prefix in (
        (winreg.HKEY_CURRENT_USER, r"Software\Classes\CLSID"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Classes\CLSID"),
        (winreg.HKEY_CLASSES_ROOT, "CLSID"),
    ):
        try:
            with winreg.OpenKey(root, rf"{prefix}\{clsid}\LocalServer32") as key:
                command, _ = winreg.QueryValueEx(key, None)
        except OSError:
            continue
        # A COM local server commonly has `"C:\\path\\app.exe" /Automation`.
        match = re.match(r'\s*"([^"]+\.exe)"|\s*([^\s]+\.exe)', str(command), re.IGNORECASE)
        executable = match.group(1) or match.group(2) if match else None
        if executable:
            path = Path(executable)
            if path.is_file():
                return path.resolve()
    return None


def _find_libreoffice(runtime_root: Path) -> Path | None:
    candidates = [
        Path(item) for item in (
            shutil.which("soffice"),
            shutil.which("libreoffice"),
        ) if item
    ]
    if os.name == "nt":
        program_files = [Path(os.environ.get("ProgramFiles", r"C:\Program Files"))]
        program_files.append(Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")))
        for root in program_files:
            candidates.append(root / "LibreOffice" / "program" / "soffice.exe")
    else:
        # The supported runtime is Windows; do not advertise or probe other OS layouts.
        return None
    candidates.extend(_registry_candidates("libreoffice"))
    manifest_path = runtime_root / "runtime-manifest.json"
    if manifest_path.is_file():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            managed_path = manifest.get("libreoffice_path")
            if managed_path:
                candidates.insert(0, Path(str(managed_path)))
        except (OSError, json.JSONDecodeError):
            pass
    if runtime_root.is_dir():
        candidates.extend(runtime_root.glob("**/program/soffice.exe"))
        candidates.extend(runtime_root.glob("**/soffice"))
    return _existing_file(candidates)


def _runtime_entry(runtime_id: str, name: str, path: Path | None, method: str) -> dict[str, Any]:
    return {
        "id": runtime_id,
        "name": name,
        "available": path is not None,
        "path": str(path) if path else None,
        "method": method if path else None,
    }


def _find_poppler_bin() -> Path | None:
    dependencies = Path(sys.executable).resolve().parent.parent
    candidates = [dependencies / "native" / "poppler" / "Library" / "bin"]
    for command in ("pdfinfo.exe", "pdfinfo"):
        executable = shutil.which(command)
        if executable and Path(executable).suffix.lower() not in {".cmd", ".bat"}:
            candidates.append(Path(executable).resolve().parent)
    for candidate in candidates:
        pdfinfo = _existing_file([candidate / "pdfinfo.exe", candidate / "pdfinfo"])
        pdftoppm = _existing_file([candidate / "pdftoppm.exe", candidate / "pdftoppm"])
        if pdfinfo and pdftoppm:
            return candidate.resolve()
    return None


def detect_pdf_rasterizer() -> dict[str, Any]:
    poppler_bin = _find_poppler_bin()
    if poppler_bin:
        return {
            "id": "poppler",
            "name": "Poppler",
            "available": True,
            "path": str(poppler_bin),
            "method": "native PDF page rasterizer",
        }
    pdfium = importlib.util.find_spec("pypdfium2")
    if pdfium is not None:
        return {
            "id": "pypdfium2",
            "name": "PDFium",
            "available": True,
            "path": str(pdfium.origin or "pypdfium2"),
            "method": "Python PDF page rasterizer",
        }
    return {
        "id": None,
        "name": "PDF page rasterizer",
        "available": False,
        "path": None,
        "method": None,
    }


def detect_renderers(runtime_root: str | Path) -> dict[str, Any]:
    root = Path(runtime_root).resolve()
    word = _find_word(root)
    word_name = "Microsoft Word"
    word_method = "installed executable"
    if not word:
        word = _word_application_com_server()
        if word:
            word_method = "Windows Word.Application COM"
            if "wps" in word.name.lower() or "kingsoft" in str(word).lower():
                word_name = "WPS Office（Word 兼容）"
    libreoffice = _find_libreoffice(root)
    preferred = "word" if word else "libreoffice" if libreoffice else None
    pdf_rasterizer = detect_pdf_rasterizer()
    return {
        "word": _runtime_entry("word", word_name, word, word_method),
        "libreoffice": _runtime_entry("libreoffice", "LibreOffice", libreoffice, "installed or managed runtime"),
        "preferred": preferred,
        "renderer_ready": preferred is not None,
        "pdf_rasterizer": pdf_rasterizer,
        "visual_ready": preferred is not None and pdf_rasterizer["available"],
        "official_download_page": OFFICIAL_DOWNLOAD_PAGE,
    }


def resolve_official_download() -> dict[str, str]:
    if os.name != "nt":
        raise RuntimeError("AI 下载运行包目前只支持 Windows；请从官网下载对应系统版本")
    request = Request(OFFICIAL_DOWNLOAD_PAGE, headers={"User-Agent": "Word-Format-Master/1.0"})
    with urlopen(request, timeout=30) as response:
        html = response.read(4 * 1024 * 1024).decode("utf-8", errors="replace")
    match = DOWNLOAD_PATTERN.search(html)
    if not match:
        raise RuntimeError("未能从 LibreOffice 官网解析当前 Windows 运行包地址")
    url = match.group(0)
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.netloc.lower() != DOWNLOAD_HOST:
        raise RuntimeError("下载地址不是受信任的 LibreOffice 官方域名")
    filename = Path(parsed.path).name
    return {"url": url, "filename": filename, "page": OFFICIAL_DOWNLOAD_PAGE}


def _emit(progress: RuntimeProgress | None, **payload: Any) -> None:
    if progress:
        progress(payload)


def download_libreoffice(runtime_root: str | Path, progress: RuntimeProgress | None = None) -> dict[str, Any]:
    root = Path(runtime_root).resolve()
    download_dir = root / "downloads"
    install_dir = root / "libreoffice"
    download_dir.mkdir(parents=True, exist_ok=True)
    root.mkdir(parents=True, exist_ok=True)
    package = resolve_official_download()
    archive = download_dir / package["filename"]
    partial = archive.with_suffix(archive.suffix + ".part")

    try:
        _emit(progress, status="downloading", phase="下载 LibreOffice 官方运行包", percent=1, downloaded=0, total=0)
        request = Request(package["url"], headers={"User-Agent": "Word-Format-Master/1.0"})

        # Download with SHA-256 computation
        sha256_hash = hashlib.sha256()
        with urlopen(request, timeout=60) as response:
            total = int(response.headers.get("Content-Length") or 0)
            downloaded = 0
            with partial.open("wb") as destination:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    destination.write(chunk)
                    sha256_hash.update(chunk)
                    downloaded += len(chunk)
                    percent = min(88, int(downloaded * 88 / total)) if total else 1
                    _emit(progress, status="downloading", phase="下载 LibreOffice 官方运行包", percent=percent, downloaded=downloaded, total=total)

        # Verify download completed successfully
        if total > 0 and downloaded != total:
            raise RuntimeError(f"下载不完整：期望 {total} 字节，实际下载 {downloaded} 字节")

        computed_hash = sha256_hash.hexdigest()
        _emit(progress, status="verifying", phase="验证下载文件完整性", percent=89, downloaded=downloaded, total=total)

        # Log the hash for verification (consider fetching official checksums in future)
        (download_dir / f"{archive.name}.sha256").write_text(
            f"{computed_hash}  {archive.name}\n",
            encoding="utf-8"
        )

        partial.replace(archive)
        _emit(progress, status="installing", phase="解包并检测 LibreOffice", percent=90, downloaded=downloaded, total=total)
    except Exception as exc:
        # Clean up partial download on any error
        if partial.exists():
            try:
                partial.unlink()
            except OSError:
                pass
        raise
    if archive.suffix.lower() != ".msi":
        raise RuntimeError(f"暂不支持的 LibreOffice 运行包格式：{archive.suffix}")
    msiexec = shutil.which("msiexec.exe") or str(Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "msiexec.exe")
    if not Path(msiexec).is_file():
        raise RuntimeError("未找到 Windows MSI 解包工具 msiexec.exe")
    install_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [msiexec, "/a", str(archive), "/qn", f"TARGETDIR={install_dir}"],
        check=False,
        capture_output=True,
        timeout=900,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 解包失败，msiexec 返回码 {result.returncode}")
    executable = _find_libreoffice(root)
    if not executable:
        raise RuntimeError("解包完成但未找到 soffice 可执行文件")
    (root / "runtime-manifest.json").write_text(
        json.dumps({"libreoffice_path": str(executable), "source_url": package["url"]}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    _emit(progress, status="ready", phase="LibreOffice 已安装并可用", percent=100, downloaded=downloaded, total=total)
    return {"status": "ready", "path": str(executable), "source_url": package["url"]}
