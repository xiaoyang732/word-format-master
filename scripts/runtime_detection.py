#!/usr/bin/env python3
"""Detect installed local Word rendering runtimes."""

from __future__ import annotations

import importlib.util
import json
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Any


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
    }
