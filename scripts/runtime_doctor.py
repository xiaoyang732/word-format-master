"""Standard-library bootstrap for diagnosis before the engine can import."""
from __future__ import annotations

import argparse
import importlib
import json
from pathlib import Path
import sys
import subprocess
import tempfile

from json_output import ensure_protected_paths, write_json
from runtime_detection import detect_renderers

ROOT = Path(__file__).resolve().parents[1]


def diagnose(output_dir=None, probe_render=False):
    missing = []
    for module, package in (("docx", "python-docx"), ("lxml.etree", "lxml"), ("PIL.Image", "Pillow")):
        try:
            importlib.import_module(module)
        except (ImportError, OSError):
            missing.append(package)
    core = sys.version_info >= (3, 10) and not missing
    actions, checks = [], {}
    if not core:
        actions.append("使用 Python 3.10+，运行 python -m pip install -r requirements.txt 安装核心依赖")
    report = False
    if core:
        try:
            from word_format.reports import load_report_catalog
            load_report_catalog()
            report = True
            checks["report_resources"] = {"status": "passed"}
        except (ValueError, OSError, KeyError, TypeError) as exc:
            checks["report_resources"] = {"status": "failed", "code": getattr(exc, "code", "INVALID_CATALOG"), "detail": str(exc)}
            actions.append("报告模板或说明不可用，请重新安装完整发布包")
    writable = True
    try:
        # Do not create an arbitrary user directory during diagnosis.
        parent = Path(output_dir).resolve() if output_dir else Path(tempfile.gettempdir())
        with tempfile.TemporaryDirectory(prefix=".word-doctor-", dir=parent):
            pass
        checks["workspace"] = {"status": "passed", "path": str(parent)}
    except OSError as exc:
        writable = False
        checks["workspace"] = {"status": "failed", "detail": str(exc)}
        actions.append("选择已存在且可写的任务目录或修复临时目录权限")
    runtime = detect_renderers(ROOT / ".runtime")
    probe = {"status": "not_performed"}
    if not runtime["visual_ready"]:
        actions.append("逐页检查需要 Word/WPS/LibreOffice 及 PDF 页面转图片后端；核心排版可单独使用")
    if probe_render:
        try:
            from docx import Document
            from render_docx import render_document
            with tempfile.TemporaryDirectory(prefix="word-render-probe-") as folder:
                source = Path(folder) / "probe.docx"
                doc = Document()
                doc.add_paragraph("Word Format Master render probe")
                doc.save(source)
                render_document(source, Path(folder) / "pages")
            probe = {"status": "passed"}
        except (ImportError, OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
            probe = {"status": "failed", "detail": str(exc)}
            actions.append("实际渲染探测失败，请检查办公软件是否可正常启动")
    status = "passed" if core and report and writable and probe["status"] != "failed" else "failed"
    return {"schema_version": "1.0", "status": status, "missing_dependencies": missing,
            "capabilities": {"core_formatting": core, "default_report": report,
                             "page_rendering": runtime["renderer_ready"], "page_images": runtime["visual_ready"],
                             "workspace_writable": writable},
            "checks": checks, "runtime": runtime, "render_probe": probe,
            "visual_review": "not_performed", "next_actions": actions}


def main(argv=None):
    from cli_output import configure_cli_output
    configure_cli_output()
    parser = argparse.ArgumentParser(description="检查核心排版、默认模板及视觉环境")
    parser.add_argument("--output-dir", help="检查已存在任务目录的写入权限")
    parser.add_argument("--output", help="保存 JSON 自检结果")
    parser.add_argument("--probe-render", action="store_true", help="在临时目录实际探测渲染")
    args = parser.parse_args(argv)
    try:
        protected = [p for name in ("templates", "scripts", "assets", "agents", "references")
                     for p in (ROOT / name).rglob("*") if p.is_file()]
        protected.extend(ROOT / n for n in ("SKILL.md", "README.md", "SECURITY.md", "requirements.txt"))
        ensure_protected_paths(args.output, *protected)
        result = diagnose(args.output_dir, args.probe_render)
        if args.output:
            write_json(args.output, result, *protected)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "passed" else 2
    except (ValueError, OSError) as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2
