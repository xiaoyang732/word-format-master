"""Shared protection and atomic publication for CLI JSON artifacts."""

import json
import os
from pathlib import Path
import tempfile

SKILL_ROOT = Path(__file__).resolve().parents[1]
PROTECTED_SKILL_SUBDIRS = ("templates", "scripts", "assets", "agents", "references")
PROTECTED_SKILL_FILES = ("SKILL.md", "README.md", "SECURITY.md", "requirements.txt", "LICENSE.txt")


def ensure_not_in_skill_root(destination):
    if not destination:
        return
    path = Path(destination).resolve()
    for filename in PROTECTED_SKILL_FILES:
        target = (SKILL_ROOT / filename).resolve()
        if path == target or (path.exists() and target.exists() and path.samefile(target)):
            raise ValueError("严禁写入或覆盖技能安装目录（$SKILL_DIR）内的文件；请将输出写入任务专属目录")
    for subdir in PROTECTED_SKILL_SUBDIRS:
        folder = (SKILL_ROOT / subdir).resolve()
        if path == folder or folder in path.parents:
            raise ValueError("严禁写入或覆盖技能安装目录（$SKILL_DIR）内的文件；请将输出写入任务专属目录")


def ensure_protected_paths(destination, *protected):
    if not destination:
        return
    ensure_not_in_skill_root(destination)
    path = Path(destination).resolve()
    if path.suffix.lower() in {".doc", ".docx", ".dotx", ".docm", ".dotm"}:
        raise ValueError("JSON output must not overwrite or use a Word document path")
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
