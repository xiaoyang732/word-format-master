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


def _find_word() -> Path | None:
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


def _com_local_server_executable(prog_ids: tuple[str, ...]) -> tuple[str, Path] | None:
    """Resolve the first existing executable registered for any of ``prog_ids``."""
    if os.name != "nt":
        return None
    try:
        import winreg
    except ImportError:
        return None
    for prog_id in prog_ids:
        clsid = None
        for root, subkey in (
            (winreg.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}\CLSID"),
            (winreg.HKEY_LOCAL_MACHINE, rf"Software\Classes\{prog_id}\CLSID"),
            (winreg.HKEY_CLASSES_ROOT, rf"{prog_id}\CLSID"),
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
            continue
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
            match = re.match(
                r'\s*"([^"]+\.exe)"|\s*(.+?\.exe)(?:\s|$)', str(command), re.IGNORECASE
            )
            executable = match.group(1) or match.group(2) if match else None
            if executable:
                path = Path(executable)
                if path.is_file():
                    return prog_id, path.resolve()
    return None


def _word_application_com_server() -> Path | None:
    """Resolve the executable registered for the Windows Word.Application COM class."""
    found = _com_local_server_executable(("Word.Application",))
    return found[1] if found else None


def _find_wps() -> tuple[str, Path] | None:
    """Resolve WPS Writer COM automation server (KWPS.Application / WPS.Application)."""
    found = _com_local_server_executable(("KWPS.Application", "WPS.Application"))
    if found is not None:
        return found
    registered_word = _word_application_com_server()
    if registered_word is not None and registered_word.name.casefold() == "wps.exe":
        return "Word.Application", registered_word
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
        candidates.extend(_registry_candidates("libreoffice"))
    elif sys.platform == "darwin":
        candidates.extend([
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path.home() / "Applications" / "LibreOffice.app" / "Contents" / "MacOS" / "soffice",
        ])
    else:
        candidates.extend([
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/snap/bin/libreoffice"),
        ])
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
    word = _find_word()
    word_name = "Microsoft Word"
    word_method = "installed executable"
    registered = _word_application_com_server()
    if registered and registered.name.casefold() != "winword.exe":
        # Installed WINWORD does not prove that Word.Application starts Word.
        word = None
        word_method = "Word.Application registration mismatch"
    elif not word and registered and registered.name.casefold() == "winword.exe":
        word = registered
        word_method = "Windows Word.Application COM"
    wps_found = _find_wps()
    wps_prog_id = wps_found[0] if wps_found else None
    wps_path = wps_found[1] if wps_found else None
    wps_entry = _runtime_entry(
        "wps",
        "WPS Office",
        wps_path,
        f"Windows {wps_prog_id} COM" if wps_prog_id else "Windows WPS COM",
    )
    wps_entry["prog_id"] = wps_prog_id
    libreoffice = _find_libreoffice(root)
    preferred = (
        "word" if word else "wps" if wps_path else "libreoffice" if libreoffice else None
    )
    pdf_rasterizer = detect_pdf_rasterizer()
    word_entry = _runtime_entry("word", word_name, word, word_method)
    word_entry["registered_server"] = str(registered) if registered else None
    word_entry["diagnostic"] = word_method
    return {
        "word": word_entry,
        "wps": wps_entry,
        "libreoffice": _runtime_entry(
            "libreoffice", "LibreOffice", libreoffice, "installed or managed runtime"
        ),
        "preferred": preferred,
        "renderer_ready": preferred is not None,
        "pdf_rasterizer": pdf_rasterizer,
        "visual_ready": preferred is not None and pdf_rasterizer["available"],
    }
