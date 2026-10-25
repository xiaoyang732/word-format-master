# -*- coding: utf-8 -*-
"""
执行 word-format-master 针对平面设计论文的规划、应用、视觉渲染与独立复核流水线
"""

import json
import subprocess
import sys
from pathlib import Path

# 定位当前技能目录与工具路径
SKILL_DIR = Path(__file__).resolve().parent / "skills" / "@user_bb897d7b" / "docx-format-master"
PRESETS_PATH = SKILL_DIR / "assets" / "presets.json"
FORMAT_CLI = SKILL_DIR / "scripts" / "format_cli.py"
RENDER_DOCX = SKILL_DIR / "scripts" / "render_docx.py"

OUTPUT_DIR = Path("output").resolve()
DRAFT_DOCX = OUTPUT_DIR / "生成式人工智能视域下现代平面设计的视觉语言重构与人机协同范式研究-初稿.docx"
FINAL_DOCX = OUTPUT_DIR / "生成式人工智能视域下现代平面设计的视觉语言重构与人机协同范式研究-已排版.docx"
REQUEST_JSON = OUTPUT_DIR / "gd_request.json"
PLAN_JSON = OUTPUT_DIR / "gd_plan.json"
REPORT_JSON = OUTPUT_DIR / "gd_report.json"
VERIFY_JSON = OUTPUT_DIR / "gd_verification.json"
RENDER_DIR = OUTPUT_DIR / "render_gd_thesis"


def run_pipeline():
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
    REQUEST_JSON.write_text(json.dumps(request_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[1/5] 排版规范请求文件已就绪: {REQUEST_JSON.name}")

    # 2. 生成排版计划 (plan)
    cmd_plan = [
        sys.executable,
        str(FORMAT_CLI),
        "plan",
        str(DRAFT_DOCX),
        str(FINAL_DOCX),
        "--request",
        str(REQUEST_JSON),
        "--output",
        str(PLAN_JSON),
    ]
    print(f"[2/5] 正在生成排版计划 (Plan)...")
    res_plan = subprocess.run(cmd_plan, capture_output=True, text=True, encoding="utf-8")
    if res_plan.returncode != 0:
        print("Plan 失败:")
        print("Stderr:", res_plan.stderr)
        print("Stdout:", res_plan.stdout)
        sys.exit(res_plan.returncode)
    print(f"      排版计划生成成功: {PLAN_JSON.name}")

    # 3. 执行排版修改 (apply)
    cmd_apply = [
        sys.executable,
        str(FORMAT_CLI),
        "apply",
        "--plan",
        str(PLAN_JSON),
        "--report",
        str(REPORT_JSON),
        "--overwrite",
    ]
    print(f"[3/5] 正在应用确定性排版修改 (Apply)...")
    res_apply = subprocess.run(cmd_apply, capture_output=True, text=True, encoding="utf-8")
    if res_apply.returncode != 0:
        print("Apply 失败:")
        print("Stderr:", res_apply.stderr)
        print("Stdout:", res_apply.stdout)
        sys.exit(res_apply.returncode)
    print(f"      排版修改已安全写入: {FINAL_DOCX.name}")

    # 4. 独立结构与格式复核 (verify)
    cmd_verify = [
        sys.executable,
        str(FORMAT_CLI),
        "verify",
        "--plan",
        str(PLAN_JSON),
        "--report",
        str(REPORT_JSON),
        "--output",
        str(VERIFY_JSON),
    ]
    print(f"[4/5] 正在执行独立格式与结构复核 (Verify)...")
    res_verify = subprocess.run(cmd_verify, capture_output=True, text=True, encoding="utf-8")
    if res_verify.returncode != 0:
        print("Verify 失败:")
        print("Stderr:", res_verify.stderr)
        print("Stdout:", res_verify.stdout)
        sys.exit(res_verify.returncode)
    print(f"      独立复核通过，报告已保存: {VERIFY_JSON.name}")

    # 5. 页面视觉渲染与复核 (可选但推荐)
    RENDER_DIR.mkdir(parents=True, exist_ok=True)
    cmd_render = [
        sys.executable,
        str(RENDER_DOCX),
        str(FINAL_DOCX),
        "--output-dir",
        str(RENDER_DIR),
        "--method",
        "auto",
    ]
    print(f"[5/5] 正在调用本地渲染器生成页面视觉图像 (Render)...")
    res_render = subprocess.run(cmd_render, capture_output=True, text=True, encoding="utf-8")
    if res_render.returncode == 0:
        print("      页面渲染成功，页面图像保存在:", RENDER_DIR)
        manifest_path = RENDER_DIR / "render-manifest.json"
        if manifest_path.exists():
            print(f"      渲染清单已生成: {manifest_path.name}")
    else:
        print("      渲染提示 (非阻断性):", res_render.stderr)

    print("=" * 60)
    print("平面设计论文排版流水线全部顺利完成！")
    print(f"最终成果文档: {FINAL_DOCX}")
    print("=" * 60)


if __name__ == "__main__":
    run_pipeline()
