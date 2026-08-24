#!/usr/bin/env python3
"""Render a DOCX to page PNGs with Microsoft Word or LibreOffice."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from config import DEFAULT_RUNTIME_ROOT, RENDER_METHODS, SKILL_DIR
from runtime_detection import detect_pdf_rasterizer, detect_renderers


def resolve_renderer(runtime: dict[str, Any], requested: str) -> tuple[str, dict[str, Any]]:
    """Resolve a renderer without silently changing an explicit selection."""
    if requested not in RENDER_METHODS:
        raise ValueError(f"不支持的渲染方式：{requested}")
    selected = runtime.get("preferred") if requested == "auto" else requested
    if not selected:
        raise RuntimeError("未检测到 Microsoft Word 或 LibreOffice")
    entry = runtime.get(selected) or {}
    if not entry.get("available"):
        name = entry.get("name") or selected
        raise RuntimeError(f"已选择 {name}，但当前系统未检测到该渲染器")
    return selected, entry


def _nonempty(path: Path) -> bool:
    return path.is_file() and path.stat().st_size > 0


def _word_to_pdf(source: Path, output_pdf: Path) -> None:
    shell = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh")
    if not shell:
        raise RuntimeError("未找到 PowerShell，无法调用 Microsoft Word 导出 PDF")
    script = r'''
param([string]$InputPath, [string]$OutputPath)
$word = $null
$document = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $document = $word.Documents.Open($InputPath, $false, $true)
  foreach ($field in $document.Fields) { try { $null = $field.Update() } catch {} }
  $document.ExportAsFixedFormat($OutputPath, 17)
} finally {
  if ($document -ne $null) { $document.Close(0) }
  if ($word -ne $null) { $word.Quit() }
  if ($document -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
  if ($word -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
'''.strip()
    with tempfile.TemporaryDirectory(prefix="word-format-word-") as temp:
        script_path = Path(temp) / "export_pdf.ps1"
        script_path.write_text(script + "\n", encoding="utf-8")
        result = subprocess.run(
            [
                shell,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(source),
                str(output_pdf),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
    if result.returncode != 0 or not _nonempty(output_pdf):
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"Microsoft Word 导出 PDF 失败：{detail[-800:] or '没有生成 PDF'}")


def _libreoffice_to_pdf(source: Path, output_pdf: Path, executable: Path) -> None:
    with (
        tempfile.TemporaryDirectory(prefix="word-format-lo-profile-") as profile,
        tempfile.TemporaryDirectory(prefix="word-format-lo-output-") as conversion,
    ):
        profile_uri = Path(profile).resolve().as_uri()
        conversion_dir = Path(conversion)
        environment = os.environ.copy()
        environment["HOME"] = profile
        embedded_python = _nonempty(executable.parent / "python.exe") and executable.parent / "python.exe"
        uno_script = Path(__file__).resolve().with_name("libreoffice_render.py")
        if embedded_python and uno_script.is_file():
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
                listener.bind(("127.0.0.1", 0))
                port = listener.getsockname()[1]
            server = subprocess.Popen(
                [
                    str(executable),
                    f"-env:UserInstallation={profile_uri}",
                    "--headless",
                    "--invisible",
                    "--norestore",
                    f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
            )
            try:
                result = subprocess.run(
                    [str(embedded_python), str(uno_script), str(source), str(output_pdf), "--port", str(port)],
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=180,
                    env=environment,
                )
            finally:
                try:
                    server.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    server.terminate()
                    try:
                        server.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        server.kill()
            if result.returncode != 0 or not _nonempty(output_pdf):
                detail = (result.stderr or result.stdout).strip()
                raise RuntimeError(f"LibreOffice UNO 刷新域并导出 PDF 失败：{detail[-800:] or '没有生成 PDF'}")
            return
        result = subprocess.run(
            [
                str(executable),
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--invisible",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(conversion_dir),
                str(source),
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
            env=environment,
        )
        generated = conversion_dir / f"{source.stem}.pdf"
        if result.returncode != 0 or not _nonempty(generated):
            detail = (result.stderr or result.stdout).strip()
            raise RuntimeError(f"LibreOffice 导出 PDF 失败：{detail[-800:] or '没有生成 PDF'}")
        shutil.copy2(generated, output_pdf)


def _pdf_to_pngs(pdf: Path, output_dir: Path, dpi: int) -> list[Path]:
    for stale in output_dir.glob("page-*.png"):
        stale.unlink()
    rasterizer = detect_pdf_rasterizer()
    failures: list[str] = []
    if rasterizer.get("id") == "poppler":
        try:
            from pdf2image import convert_from_path

            images = convert_from_path(
                str(pdf),
                dpi=dpi,
                fmt="png",
                thread_count=4,
                poppler_path=str(rasterizer["path"]),
            )
            pages = []
            for index, image in enumerate(images, start=1):
                page = output_dir / f"page-{index}.png"
                image.save(page, format="PNG")
                image.close()
                pages.append(page)
            if pages:
                return pages
            failures.append("Poppler 没有生成任何页面")
        except Exception as exc:
            failures.append(f"Poppler: {exc}")
            for stale in output_dir.glob("page-*.png"):
                stale.unlink()
    try:
        import pypdfium2 as pdfium

        document = pdfium.PdfDocument(str(pdf))
        pages = []
        try:
            for index in range(len(document)):
                page_object = document[index]
                bitmap = page_object.render(scale=dpi / 72)
                image = bitmap.to_pil()
                page = output_dir / f"page-{index + 1}.png"
                image.save(page, format="PNG")
                image.close()
                bitmap.close()
                page_object.close()
                pages.append(page)
        finally:
            document.close()
        if pages:
            return pages
        failures.append("PDFium 没有生成任何页面")
    except Exception as exc:
        failures.append(f"PDFium: {exc}")
    raise RuntimeError(f"PDF 逐页转图失败：{'；'.join(failures) or '缺少 Poppler 或 PDFium'}")


def render_document(
    source: str | Path,
    output_dir: str | Path,
    *,
    method: str = "auto",
    runtime_root: str | Path = DEFAULT_RUNTIME_ROOT,
    dpi: int = 150,
    emit_pdf: bool = False,
) -> dict[str, Any]:
    input_path = Path(source).resolve()
    if not input_path.is_file() or input_path.suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("渲染输入必须是现有的 DOCX 或 DOTX 文件")
    target = Path(output_dir).resolve()
    target.mkdir(parents=True, exist_ok=True)
    runtime = detect_renderers(runtime_root)
    selected, entry = resolve_renderer(runtime, method)
    candidates = [(selected, entry)]
    if method == "auto" and selected == "word" and runtime.get("libreoffice", {}).get("available"):
        candidates.append(("libreoffice", runtime["libreoffice"]))
    temporary_pdf = target / f".{input_path.stem}.rendering.pdf"
    final_pdf = target / f"{input_path.stem}.pdf"
    if temporary_pdf.exists():
        temporary_pdf.unlink()
    try:
        failures = []
        for candidate_method, candidate_entry in candidates:
            temporary_pdf.unlink(missing_ok=True)
            try:
                if candidate_method == "word":
                    _word_to_pdf(input_path, temporary_pdf)
                else:
                    _libreoffice_to_pdf(input_path, temporary_pdf, Path(candidate_entry["path"]))
                selected, entry = candidate_method, candidate_entry
                break
            except (OSError, RuntimeError, subprocess.TimeoutExpired) as exc:
                failures.append(f"{candidate_entry.get('name') or candidate_method}: {exc}")
        else:
            raise RuntimeError("；".join(failures))
        pages = _pdf_to_pngs(temporary_pdf, target, dpi)
        page_manifest = {
            "schema_version": "1.0",
            "source": str(input_path),
            "source_sha256": hashlib.sha256(input_path.read_bytes()).hexdigest(),
            "selected_method": selected,
            "page_count": len(pages),
            "pages": [
                {
                    "number": index,
                    "path": str(page),
                    "sha256": hashlib.sha256(page.read_bytes()).hexdigest(),
                }
                for index, page in enumerate(pages, start=1)
            ],
        }
        manifest_path = target / "render-manifest.json"
        manifest_path.write_text(json.dumps(page_manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        if emit_pdf:
            shutil.copy2(temporary_pdf, final_pdf)
    finally:
        temporary_pdf.unlink(missing_ok=True)
    return {
        "status": "rendered",
        "requested_method": method,
        "selected_method": selected,
        "renderer": entry.get("name"),
        "page_count": len(pages),
        "pages": [str(page) for page in pages],
        "manifest": str(manifest_path),
        "pdf": str(final_pdf) if emit_pdf else None,
    }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Input DOCX or DOTX")
    parser.add_argument("--output-dir", required=True, help="Directory for page-<N>.png files")
    parser.add_argument("--method", choices=sorted(RENDER_METHODS), default="auto")
    parser.add_argument("--runtime-root", default=str(DEFAULT_RUNTIME_ROOT))
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--emit-pdf", action="store_true")
    args = parser.parse_args()
    if args.dpi < 72 or args.dpi > 300:
        raise SystemExit("--dpi 必须在 72 到 300 之间")
    try:
        report = render_document(
            args.input,
            args.output_dir,
            method=args.method,
            runtime_root=args.runtime_root,
            dpi=args.dpi,
            emit_pdf=args.emit_pdf,
        )
    except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False))
        return 1
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
