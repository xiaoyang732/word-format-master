# -*- coding: utf-8 -*-
"""
执行 word-format-master 针对平面设计论文的规划、应用、视觉渲染与独立复核流水线
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

# 定位当前项目目录与工具路径。
SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent
PRESETS_PATH = SKILL_DIR / "assets" / "presets.json"
FORMAT_CLI = SCRIPTS_DIR / "format_cli.py"
RENDER_DOCX = SCRIPTS_DIR / "render_docx.py"
THESIS_QUALITY = SCRIPTS_DIR / "thesis_quality.py"

OUTPUT_DIR = SKILL_DIR / "output"


def run_pipeline(
    draft_path=None,
    output_dir=None,
    minimum_chinese_characters=None,
    maximum_chinese_characters=None,
):
    output_dir = Path(output_dir).expanduser().resolve() if output_dir else OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)
    draft_docx = (
        Path(draft_path).expanduser().resolve()
        if draft_path
        else output_dir / "生成式人工智能视域下现代平面设计的视觉语言重构与人机协同范式研究-初稿.docx"
    )
    final_docx = output_dir / "生成式人工智能视域下现代平面设计的视觉语言重构与人机协同范式研究-已排版.docx"
    request_json = output_dir / "gd_request.json"
    plan_json = output_dir / "gd_plan.json"
    report_json = output_dir / "gd_report.json"
    verify_json = output_dir / "gd_verification.json"
    quality_json = output_dir / "gd_content_quality.json"
    render_dir = output_dir / "render_gd_thesis"

    print("=" * 60)
    print("开始执行平面设计毕业论文 word-format-master 自动化排版流水线")
    print("=" * 60)

    # 1. 提取 thesis-standard 规范
    presets_data = json.loads(PRESETS_PATH.read_text(encoding="utf-8"))
    thesis_preset = next(p for p in presets_data["presets"] if p["id"] == "thesis-standard")

    request_data = {
        "schema_version": "2.0",
        "origin": "legacy",
        "request_text": "采用内置标准学位/毕业论文规范模板 (thesis-standard) 对平面设计论文全篇进行排版与格式规范化",
        "spec": thesis_preset,
    }
    request_json.write_text(json.dumps(request_data, ensure_ascii=False, indent=2), encoding="utf-8")
    quality_cmd = [sys.executable, str(THESIS_QUALITY), str(draft_docx), "--output", str(quality_json)]
    if minimum_chinese_characters is not None:
        quality_cmd.extend(["--min-chinese-characters", str(minimum_chinese_characters)])
    if maximum_chinese_characters is not None:
        quality_cmd.extend(["--max-chinese-characters", str(maximum_chinese_characters)])
    print("[1/6] 正在检查正文篇幅与引用映射（书目事实仍需人工核验）...")
    quality_result = subprocess.run(quality_cmd, capture_output=True, text=True, encoding="utf-8")
    if quality_result.returncode not in (0, 1):
        print("内容质量报告生成失败:", quality_result.stderr or quality_result.stdout)
        sys.exit(quality_result.returncode)
    if quality_result.returncode == 1:
        print(f"内容检查门槛未通过或无法评估，请查看报告: {quality_json}")
        sys.exit(1)
    quality_summary = json.loads(quality_json.read_text(encoding="utf-8"))
    print(
        "      内容状态: "
        f"{quality_summary['status']}；正文汉字数 "
        f"{quality_summary['measurement_scope']['body']['chinese_han_characters']}；"
        f"引用映射 {quality_summary['checks']['numeric_citation_mapping']['status']}；"
        f"书目来源 {quality_summary['checks']['bibliography_source_verification']['status']}"
    )
    print(f"      内容检查报告已生成: {quality_json.name}")
    print(f"[2/6] 排版规范请求文件已就绪: {request_json.name}")

    # 2. 生成排版计划 (plan)
    cmd_plan = [
        sys.executable,
        str(FORMAT_CLI),
        "plan",
        str(draft_docx),
        str(final_docx),
        "--request",
        str(request_json),
        "--output",
        str(plan_json),
    ]
    print("[3/6] 正在生成排版计划 (Plan)...")
    res_plan = subprocess.run(cmd_plan, capture_output=True, text=True, encoding="utf-8")
    if res_plan.returncode != 0:
        print("Plan 失败:")
        print("Stderr:", res_plan.stderr)
        print("Stdout:", res_plan.stdout)
        sys.exit(res_plan.returncode)
    print(f"      排版计划生成成功: {plan_json.name}")

    # 3. 执行排版修改 (apply)
    cmd_apply = [
        sys.executable,
        str(FORMAT_CLI),
        "apply",
        "--plan",
        str(plan_json),
        "--report",
        str(report_json),
        "--overwrite",
    ]
    print("[4/6] 正在应用确定性排版修改 (Apply)...")
    res_apply = subprocess.run(cmd_apply, capture_output=True, text=True, encoding="utf-8")
    if res_apply.returncode != 0:
        print("Apply 失败:")
        print("Stderr:", res_apply.stderr)
        print("Stdout:", res_apply.stdout)
        sys.exit(res_apply.returncode)
    print(f"      排版修改已安全写入: {final_docx.name}")

    # 4. 独立结构与格式复核 (verify)
    cmd_verify = [
        sys.executable,
        str(FORMAT_CLI),
        "verify",
        "--plan",
        str(plan_json),
        "--report",
        str(report_json),
        "--output",
        str(verify_json),
    ]
    print("[5/6] 正在执行独立格式与结构复核 (Verify)...")
    res_verify = subprocess.run(cmd_verify, capture_output=True, text=True, encoding="utf-8")
    if res_verify.returncode != 0:
        print("Verify 失败:")
        print("Stderr:", res_verify.stderr)
        print("Stdout:", res_verify.stdout)
        sys.exit(res_verify.returncode)
    print(f"      独立复核通过，报告已保存: {verify_json.name}")

    # 5. 页面视觉渲染与复核 (可选但推荐)
    render_dir.mkdir(parents=True, exist_ok=True)
    cmd_render = [
        sys.executable,
        str(RENDER_DOCX),
        str(final_docx),
        "--output-dir",
        str(render_dir),
        "--method",
        "auto",
    ]
    print("[6/6] 正在调用本地渲染器生成页面视觉图像 (Render)...")
    res_render = subprocess.run(cmd_render, capture_output=True, text=True, encoding="utf-8")
    if res_render.returncode == 0:
        print("      页面渲染成功，页面图像保存在:", render_dir)
        manifest_path = render_dir / "render-manifest.json"
        if manifest_path.exists():
            print(f"      渲染清单已生成: {manifest_path.name}")
    else:
        print("      渲染提示 (非阻断性):", res_render.stderr)

    print("=" * 60)
    print("论文排版流程结束；请结合内容检查报告处理书目核验和其他人工审校项。")
    print(f"最终成果文档: {final_docx}")
    print(f"内容质量报告: {quality_json}")
    print("=" * 60)


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", help="Draft DOCX (default: output folder's named draft)")
    parser.add_argument("--output-dir", help="Pipeline output directory (default: repository/output)")
    parser.add_argument("--min-chinese-characters", type=int, help="Optional minimum thesis body length")
    parser.add_argument("--max-chinese-characters", type=int, help="Optional maximum thesis body length")
    args = parser.parse_args()
    run_pipeline(
        draft_path=args.draft,
        output_dir=args.output_dir,
        minimum_chinese_characters=args.min_chinese_characters,
        maximum_chinese_characters=args.max_chinese_characters,
    )
