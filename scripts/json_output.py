"""Shared protection and atomic publication for CLI JSON artifacts."""

import json
import os
from pathlib import Path
import tempfile


def ensure_protected_paths(destination, *protected):
    if not destination:
        return
    path = Path(destination).resolve()
    if path.suffix.lower() in {".doc", ".docx", ".dotx", ".docm", ".dotm"}:
        raise ValueError("JSON output must not use a Word document path")
    for value in protected:
        if not value:
            continue
        source = Path(value).resolve()
        if path == source or (path.exists() and source.exists() and path.samefile(source)):
            raise ValueError("JSON output must not overwrite document, request, plan or verification input")


def write_json(destination, value, *protected):
    ensure_protected_paths(destination, *protected)
    rendered = json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    path = Path(destination).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         prefix="." + path.name + ".", delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(rendered)
        ensure_protected_paths(path, *protected)
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
