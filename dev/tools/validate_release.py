"""Validate the delivered ZIP and run its report CLI in a fresh installation."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import subprocess
import sys
import tempfile
from zipfile import BadZipFile, ZipFile


def validate_release(path):
    with tempfile.TemporaryDirectory(prefix="word-release-check-") as folder:
        work = Path(folder)
        install = work / "中文 安装目录"
        with ZipFile(path) as archive:
            seen = set()
            if sum(i.file_size for i in archive.infolist()) > 300 * 1024 * 1024:
                raise ValueError("发布包超过验证大小限制")
            for item in archive.infolist():
                name = item.filename
                part = PurePosixPath(name)
                if (part.is_absolute() or ".." in part.parts or "\\" in name or ":" in name
                        or name.casefold() in seen or (item.external_attr >> 16) & 0o170000 == 0o120000):
                    raise ValueError(f"非法发布路径：{name}")
                seen.add(name.casefold())
            manifest = json.loads(archive.read("release-manifest.json"))
            expected = manifest["files"]
            actual = set(archive.namelist()) - {"release-manifest.json"}
            if set(expected) != actual:
                raise ValueError(f"发布文件不完整：{sorted(set(expected) ^ actual)}")
            for name, digest in expected.items():
                if hashlib.sha256(archive.read(name)).hexdigest() != digest:
                    raise ValueError(f"发布文件摘要不匹配：{name}")
            archive.extractall(install)
        cache = work / "cache"
        cache.mkdir()
        env = dict(os.environ, TMP=str(cache), TEMP=str(cache), TMPDIR=str(cache),
                   PYTHONPATH="", PYTHONDONTWRITEBYTECODE="1")
        for args in (("report", "list"), ("report", "prepare", str(work / "task"))):
            result = subprocess.run([sys.executable, "-E", "-B", str(install / "scripts/format_cli.py"), *args],
                                    cwd=work, env=env, capture_output=True, encoding="utf-8", timeout=90)
            if result.returncode:
                raise ValueError(f"分发包运行失败：{result.stderr}")
        workflow = json.loads((work / "task/workflow.json").read_text(encoding="utf-8"))
        if hashlib.sha256((work / "task/draft.docx").read_bytes()).hexdigest() != workflow["source"]["sha256"]:
            raise ValueError("分发包草稿与模板摘要不一致")
    return {"status": "passed", "checks": ["file_digests", "isolated_report_list", "isolated_report_prepare"]}


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("archive")
    args = parser.parse_args()
    try:
        print(json.dumps(validate_release(args.archive), ensure_ascii=False))
    except (ValueError, OSError, KeyError, BadZipFile, subprocess.SubprocessError) as exc:
        parser.exit(2, str(exc) + "\n")
