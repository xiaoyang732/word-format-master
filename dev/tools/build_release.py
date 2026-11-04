"""Build and verify a complete runtime archive; never include local dev data."""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

from validate_release import validate_release

ROOT = Path(__file__).resolve().parents[2]


def build_release(output, template_mode="native", root=ROOT):
    root, output = Path(root).resolve(), Path(output).resolve()
    if template_mode not in {"native", "json"}:
        raise ValueError("Unknown template mode")
    if output.suffix != ".zip" or output.exists():
        raise ValueError("请选择尚不存在的 ZIP 输出路径")
    files = {}
    for name in ("LICENSE.txt", "README.md", "requirements.txt", "SECURITY.md", "SKILL.md"):
        files[name] = (root / name).read_bytes()
    for folder, suffixes in (("scripts", {".py"}), ("references", {".md"}),
                             ("agents", {".yaml"}), ("assets", {".json"}), ("templates", {".json", ".md", ".docx"})):
        for source in sorted((root / folder).rglob("*")):
            if not source.is_file() or source.suffix not in suffixes or "__pycache__" in source.parts:
                continue
            if source.is_symlink() or root not in source.resolve().parents:
                raise ValueError(f"Unsafe runtime resource: {source}")
            files[source.relative_to(root).as_posix()] = source.read_bytes()
    catalog = json.loads(files["templates/reports.json"])
    for entry in catalog["templates"]:
        for key in ("path", "source_notes", "usage_notes"):
            resource = "templates/" + entry[key]
            candidate = (root / resource).resolve()
            if (root / "templates").resolve() not in candidate.parents or resource not in files:
                raise ValueError(f"缺少或非法模板依赖：{resource}")
        resource = "templates/" + entry["path"]
        data = files[resource]
        if hashlib.sha256(data).hexdigest() != entry["sha256"]:
            raise ValueError(f"模板摘要不匹配：{resource}")
        if template_mode == "json":
            files.pop(resource)
            files[resource + ".json"] = json.dumps({"format": "docx-base64-v1", "data": base64.b64encode(data).decode("ascii")}).encode("utf-8")
    version = re.search(r"^version:\s*(\S+)", files["SKILL.md"].decode("utf-8"), re.M).group(1)
    commit = subprocess.run(["git", "rev-parse", "HEAD"], cwd=root, capture_output=True, text=True, check=True).stdout.strip()
    dirty = bool(subprocess.run(["git", "status", "--porcelain"], cwd=root, capture_output=True, text=True, check=True).stdout)
    manifest = {"schema_version": "1.0", "version": version, "commit": commit, "dirty": dirty,
                "template_mode": template_mode, "files": {name: hashlib.sha256(data).hexdigest() for name, data in sorted(files.items())}}
    files["release-manifest.json"] = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".release-", dir=output.parent) as folder:
        staged = Path(folder) / "release.zip"
        with ZipFile(staged, "w", ZIP_DEFLATED) as archive:
            for name, data in sorted(files.items()):
                archive.writestr(name, data)
        validate_release(staged)
        # Publish atomically without replacing another release, even in a race.
        os.link(staged, output)
    return {"status": "passed", "output": str(output), "version": version,
            "sha256": hashlib.sha256(output.read_bytes()).hexdigest(), "file_count": len(files)}


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True)
    parser.add_argument("--template-mode", choices=["native", "json"], default="native")
    args = parser.parse_args()
    try:
        print(json.dumps(build_release(args.output, args.template_mode), ensure_ascii=False))
    except (ValueError, OSError, KeyError, subprocess.SubprocessError) as exc:
        parser.exit(2, str(exc) + "\n")
