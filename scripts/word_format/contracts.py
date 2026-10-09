"""Versioned, strict boundaries for requests and plans. No document writes."""

from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path
from typing import Any

SCHEMA_VERSION = "2.0"
MAX_OPERATIONS = 5000


class FormatError(ValueError):
    def __init__(self, message: str, status: str = "failed"):
        super().__init__(message)
        self.status = status


def require_keys(value: Any, allowed: set[str], required: set[str] = frozenset(), label: str = "object") -> dict:
    if not isinstance(value, dict):
        raise FormatError(f"{label} must be an object")
    unknown = set(value) - allowed
    missing = required - set(value)
    if unknown or missing:
        raise FormatError(f"{label}: unknown fields={sorted(unknown)}, missing fields={sorted(missing)}")
    return value


def number(value: Any, low: float, high: float, label: str = "value", integer: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value):
        raise FormatError(f"{label} must be a finite number")
    if not low <= value <= high or (integer and value != int(value)):
        raise FormatError(f"{label} must be {'an integer ' if integer else ''}between {low} and {high}")
    return value


def boolean(value: Any, label: str = "value") -> bool:
    if not isinstance(value, bool):
        raise FormatError(f"{label} must be boolean")
    return value


def text_value(value: Any, label: str = "value", empty: bool = False) -> str:
    if not isinstance(value, str) or (not empty and not value.strip()) or len(value) > 10000:
        raise FormatError(f"{label} must be {'non-empty ' if not empty else ''}text of at most 10000 characters")
    return value


def sha256(path: str | Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def json_hash(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, allow_nan=False).encode()).hexdigest()


def source_path(path: str | Path) -> Path:
    source = Path(path).resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise FormatError("Source must be an existing DOCX; convert DOC/DOTX through an explicit supported workflow")
    from config import MAX_DOCUMENT_BYTES
    if source.stat().st_size>MAX_DOCUMENT_BYTES: raise FormatError("Source exceeds the document size limit")
    # Reuse archive limits and conformance detection without creating a dashboard.
    from analyze_docx import package_inventory, detect_conformance

    package, _ = package_inventory(source)
    try:
        if detect_conformance(package) == "strict":
            raise FormatError("Strict OOXML requires an explicit conversion; direct writes support Transitional DOCX", "unsupported")
    finally:
        package.close()
    return source


def output_path(source: Path, output: str | Path) -> Path:
    destination = Path(output).resolve()
    if destination == source or destination.suffix.lower() != ".docx":
        raise FormatError("Output must be a different DOCX path")
    if destination.exists() and destination.samefile(source):
        raise FormatError("Output must not alias the source file")
    return destination
