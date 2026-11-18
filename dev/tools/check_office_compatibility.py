"""Run real Microsoft Word regressions in a caller-selected workspace.

Usage: python -B dev/tools/check_office_compatibility.py --workspace <outside-repo-path>
Artifacts are diagnostic fixtures, never user manuscripts. No WPS fallback is used.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START

ROOT = Path(__file__).resolve().parents[2]


def run(workspace, name, script, *args, allow_pending=False):
    command = [sys.executable, "-B", str(ROOT / "scripts" / script), *map(str, args)]
    result = subprocess.run(
        command,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        env={**os.environ, "PYTHONUTF8": "1", "PYTHONDONTWRITEBYTECODE": "1"},
    )
    (workspace / f"{name}.log").write_text(result.stdout + "\n" + result.stderr, encoding="utf-8")
    if result.returncode and not allow_pending:
        raise RuntimeError(
            f"{name} failed ({result.returncode}): {result.stdout[-500:]} {result.stderr[-1200:]}"
        )
    print(name, result.returncode, flush=True)


def check(workspace, reference_style):
    workspace.mkdir(parents=True, exist_ok=False)
    draft = workspace / "draft.docx"
    document = Document()
    document.add_paragraph("Office兼容性回归测试", "Title")
    document.add_paragraph("软件功能测试材料")
    document.add_section(WD_SECTION_START.NEW_PAGE)
    document.add_paragraph("摘要")
    document.add_paragraph(
        "本测试验证标题、文献、普通列表、动态引用、页眉与正文页码在Word中的实际结果。"
    )
    document.add_paragraph("关键词：兼容性；原生编号")
    document.add_page_break()
    document.add_paragraph("Abstract")
    document.add_paragraph(
        "This software fixture checks native numbering and dynamic fields in Microsoft Word."
    )
    document.add_paragraph("Key words: Compatibility; Native numbering")
    document.add_page_break()
    document.add_heading("绪论", 1)
    document.add_heading("测试背景", 2)
    document.add_paragraph("项目文档说明了格式处理流程。")
    document.add_paragraph("记录输入", "List Number")
    document.add_paragraph("检查输出", "List Number")
    document.add_heading("方法", 1)
    document.add_heading("验收范围", 2)
    document.add_paragraph("直接接口规定了核验流程。")
    cap_p = document.add_paragraph("表 ")
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    seq_fld = OxmlElement("w:fldSimple")
    seq_fld.set(qn("w:instr"), "SEQ 表 \\* ARABIC")
    seq_r = OxmlElement("w:r")
    seq_t = OxmlElement("w:t")
    seq_t.text = "1"
    seq_r.append(seq_t)
    seq_fld.append(seq_r)
    cap_p._p.append(seq_fld)
    cap_p.add_run(" 兼容性检查项")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "模块"
    table.cell(0, 1).text = "状态"
    table.cell(1, 0).text = "目录与文献"
    table.cell(1, 1).text = "通过"
    document.add_paragraph("参考文献")
    if reference_style:
        document.styles.add_style("Bibliography", 1)
    for name in ("README.md", "references/direct-format-api.md"):
        document.add_paragraph(
            f"Word Format Master. {name}[EB/OL]. 本地项目文档.",
            "Bibliography" if reference_style else "Normal",
        )
    document.save(draft)
    request = {
        "operations": [
            {
                "action": "citation.insert",
                "target": {"type": "document"},
                "params": {
                    "placements": [
                        {
                            "text_contains": text,
                            "reference_ids": [f"ref-{number}"],
                            "citation_text": f"[{number}]",
                            "display": "superscript",
                            "insert": "before-terminal-punctuation",
                        }
                        for number, text in enumerate(("项目文档说明", "直接接口规定"), 1)
                    ]
                },
            }
        ]
    }
    request_path = workspace / "citations.json"
    request_path.write_text(json.dumps(request, ensure_ascii=False, indent=2), encoding="utf-8")
    cited = workspace / "cited.docx"
    plan = workspace / "citation-plan.json"
    applied = workspace / "citation-apply.json"
    run(
        workspace,
        "citation-plan",
        "format_cli.py",
        "plan",
        draft,
        cited,
        "--request",
        request_path,
        "--output",
        plan,
    )
    run(workspace, "citation-apply", "format_cli.py", "apply", "--plan", plan, "--report", applied)
    run(
        workspace,
        "citation-verify",
        "format_cli.py",
        "verify",
        "--plan",
        plan,
        "--report",
        applied,
        "--output",
        workspace / "citation-verify.json",
    )
    run(
        workspace,
        "pipeline",
        "run_format_pipeline.py",
        cited,
        workspace / "formatted.docx",
        "--template",
        "thesis-standard",
        "--report-dir",
        workspace / "reports",
        "--render",
        "--render-method",
        "word",
        "--check-quality",
        "--min-references",
        "2",
        "--require-all-references-cited",
    )
    manifest = json.loads(
        (workspace / "reports/render/render-manifest.json").read_text(encoding="utf-8")
    )
    evidence = json.loads(
        (workspace / "reports/render/office-evidence.json").read_text(encoding="utf-8")
    )
    assert manifest["application"]["name"] == "Microsoft Word", manifest
    assert manifest["office_validation"]["status"] == "passed", manifest["office_validation"]
    toc2 = [p for p in evidence["paragraphs"] if p.get("style") == "TOC 2"]
    assert toc2 and all(
        p["font_east_asia"] == "宋体" and p["left_indent_chars"] == 2 for p in toc2
    ), toc2
    bib = [p for p in evidence["paragraphs"] if "README.md" in p.get("text", "")]
    expected_bib_styles = {"书目", "书目1"} if reference_style else {"书目"}
    assert bib and bib[0]["style"] in expected_bib_styles and bib[0]["font_size_pt"] == 10.5, bib
    assert bib[0]["first_line_indent_chars"] == -2, bib
    cited_body = [p for p in evidence["paragraphs"] if "项目文档说明" in p.get("text", "")]
    assert cited_body and cited_body[0]["font_size_pt"] == 12, cited_body
    return manifest


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    args = parser.parse_args()
    workspace = args.workspace.resolve()
    if workspace == ROOT or (ROOT in workspace.parents and ROOT / "dev" not in workspace.parents):
        parser.error("Use a workspace outside the repository or under dev")
    for style in (False, True):
        manifest = check(workspace / ("bibliography" if style else "normal"), style)
        print(
            json.dumps(
                {
                    "case": style,
                    "pages": manifest["page_count"],
                    "application": manifest["application"],
                },
                ensure_ascii=False,
            ),
            flush=True,
        )


if __name__ == "__main__":
    main()
