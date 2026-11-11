"""Automated diagnostics and 1-step remediation playbooks for acceptance stages."""

from __future__ import annotations

from pathlib import Path
from typing import Any

VISUAL_DEFECT_PLAYBOOK: dict[str, dict[str, str]] = {
    "cover_overflow_two_pages": {
        "symptom": "封面溢出为2页（日期掉到第2页）或误删含校名图片（<w:drawing>）的空文本段落",
        "cause": "长题目自动换行增加行高，叠加模板日期上方的多个空行；或 cover_body 规则将空行字号放大；或手动删除空段落时未检查 <w:drawing> 节点。",
        "fix": (
            "使用内置命令安全压缩封面纯空占位段（自动跳过含图片/形状/分节符的段落）："
            "python scripts/format_cli.py cover <样例> <输出.docx> --fields <fields.json> --compact-spacers 3 --overwrite"
        ),
    },
    "cover_banner_font_shrunk": {
        "symptom": "封面大字类型/横幅（如 48pt 黑体“文献综述”）被覆盖为小字号",
        "cause": "无标签封面横幅段与封面其他段落共用 cover_body 语义角色且未限定文本选择器。",
        "fix": (
            "在规则 target 中限定精确文本：{\"type\": \"paragraph\", \"module\": \"cover\", "
            "\"semantic_role\": \"cover_body\", \"text\": \"文献综述\"}（最新版 template_profile 已自动添加）。"
        ),
    },
    "cover_date_off_center_or_red_hint": {
        "symptom": "封面日期视觉偏右未居中，或残留模板红色字号批注（如“（小二号居中”）、字体/颜色报混合冲突",
        "cause": "样例封面日期末尾带有红色格式提示 run（FF0000）及补偿性首行缩进（firstLine=7字符）。",
        "fix": (
            "在 cover 的 --fields 中传入 \"date\" 字段，或重新执行 format_cli.py cover"
            "（内置封面卫生清理会自动剥离红色格式提示 run、清除居中日期的首行缩进并同步段落标记 rPr）。"
        ),
    },
    "abstract_or_toc_page_flow_drift": {
        "symptom": "中英文摘要未独立成页、关键词后直接连排下一模块，或目录前后分页漂移",
        "cause": "中文关键词段末缺少分页符（<w:br w:type=\"page\"/>），或英文关键词与目录之间缺少 nextPage 分节符。",
        "fix": "在中文关键词段末追加分页符，并在英文关键词段末与目录段末配置 nextPage 分节符（<w:sectPr>）。",
    },
    "english_title_justify_stretched": {
        "symptom": "英文摘要页长标题单词间距被异常拉大",
        "cause": "英文摘要标题（abstract_title, language=en）继承了两端对齐（justify）。",
        "fix": (
            "为英文摘要标题显式设置居中对齐：{\"action\": \"paragraph.alignment.set\", "
            "\"target\": {\"type\": \"paragraph\", \"semantic_role\": \"abstract_title\", "
            "\"language\": \"en\", \"all\": true}, \"params\": {\"value\": \"center\"}}。"
        ),
    },
    "stale_toc_or_field_numbers": {
        "symptom": "目录页码与正文章节实际渲染页码不一致，或图表交叉引用/题注域未刷新",
        "cause": "扩写或排版导致页码变化后，未在最终渲染前通过 Word 重新计算 TOC/REF/SEQ/STYLEREF 域。",
        "fix": "运行内置域刷新命令：python scripts/format_cli.py refresh-fields <文档.docx> --in-place",
    },
}

NUMBERING_KIND_REMEDIATIONS: dict[str, dict[str, str]] = {
    "heading": {
        "cause": "章节标题中包含手打编号前缀（如“第1章”、“1.1”），或标题引用的 Word 编号定义缺失。",
        "fix": (
            "去掉标题文本中的手打编号前缀，改用原生多级编号（w:numPr），"
            "或在格式化操作中使用 heading.numbering.configure / heading.scheme.configure 自动转换。"
        ),
    },
    "caption": {
        "cause": "图表题注使用了手打编号（如“表 1.1”、“图 2.1”），未使用 Word 动态域。",
        "fix": (
            "将题注编号替换为 STYLEREF 1 \\s 与 SEQ Table \\* ARABIC \\s 1（或 SEQ Figure）组合域，"
            "并用书签（如 Ref_Table_1_1）包裹编号部分，正文中通过 REF Ref_Table_1_1 \\h 交叉引用。"
        ),
    },
    "reference_or_citation": {
        "cause": "正文或参考文献列表中存在手打数字文献标号（如“[1]”），未使用动态域与原生编号。",
        "fix": (
            "删除正文中的手打 [1] 文本，在请求中使用 citation.insert（支持通过 text_contains 或 paragraph_sha256 定位段落）"
            "自动生成 REF WFMRef_N \\n \\h 动态引用域与参考文献原生 [%1] 编号列表。"
        ),
    },
    "list": {
        "cause": "正文段落开头使用了手打序号（如“1、”、“(1)”、“一、”）。",
        "fix": "使用 list.numbering.configure 转换为 Word 原生编号列表，或改用段内加粗引导词（如“首先，……；其次，……”）。",
    },
    "appendix": {
        "cause": "附录标题包含手打序号（如“附录A”）。",
        "fix": "使用 Word 编号或题注/样式编号管理附录序号。",
    },
    "equation": {
        "cause": "公式段落末尾包含手打公式编号（如“(1-1)”）。",
        "fix": "将公式编号替换为 STYLEREF + SEQ Equation 动态域。",
    },
    "page": {
        "cause": "页眉或页脚中包含手打页码数字，未使用 PAGE / NUMPAGES 域。",
        "fix": "使用 page_number.configure 自动生成原生 PAGE 域页码。",
    },
}


VISUAL_TEMPLATE_PLACEHOLDER_SUMMARY = (
    "请先调用图像查看工具逐页检查全部 page-*.png 渲染图并核对 checklist 各项后，"
    "在此填写真实的版面检查结论，并将 status 从 pending_visual_review 改为 passed（或 failed）。"
)


def build_visual_report_template(
    *,
    source_sha256: str,
    manifest_path: str,
    manifest_sha256: str,
    selected_method: str,
    page_count: int,
) -> dict[str, Any]:
    """Build a visual review template bound to the exact output and render manifest, requiring explicit sign-off."""
    return {
        "schema_version": "1.0",
        "kind": "word-format-visual-review",
        "status": "pending_visual_review",
        "output_sha256": source_sha256,
        "manifest_sha256": manifest_sha256,
        "manifest_path": manifest_path,
        "renderer": {"method": selected_method},
        "page_count": page_count,
        "reviewed_pages": list(range(1, page_count + 1)),
        "model": {"id": "multimodal-reviewer", "supports_image_input": True},
        "checklist": {
            "paper_and_margins_consistent": "全篇纸张尺寸与页边距一致（如全篇 A4，无封面/正文纸张跳变）",
            "cover_single_page_and_title_consistent": "封面单页完整收口无溢出空白页，封面题目与摘要页题目完全一致，无红色批注残留",
            "body_alignment_and_indentation": "正文段落两端对齐（justify）无右侧参差锯齿，首行缩进 2 字符",
            "headers_footers_and_page_numbers": "页眉文字与页眉横线符合规范，前置节（摘要/目录）与正文节页码格式及起始页码正确",
            "toc_completeness_and_page_match": "目录包含全部应收录章节及模块（如参考文献），且目录页码与正文实际起始页码完全一致",
            "headings_tables_and_references": "各级标题编号与标题文字间距正常无粘连，表格无出界，参考文献悬挂缩进规范",
        },
        "findings": [],
        "summary": VISUAL_TEMPLATE_PLACEHOLDER_SUMMARY,
    }


def diagnose_structure_remediation(comparison: dict[str, Any]) -> list[dict[str, Any]]:
    if comparison.get("status") == "passed":
        return []
    remediations: list[dict[str, Any]] = []
    for diff in comparison.get("differences", []):
        missing = diff.get("missing", [])
        unexpected = diff.get("unexpected", [])
        if (
            len(missing) == 1
            and len(unexpected) == 1
            and missing[0].get("kind") == unexpected[0].get("kind") == "chapters"
            and missing[0].get("outline_levels") != unexpected[0].get("outline_levels")
        ):
            remediations.append(
                {
                    "issue": "chapter_hierarchy_difference",
                    "expected_outline_levels": missing[0].get("outline_levels"),
                    "actual_outline_levels": unexpected[0].get("outline_levels"),
                    "fix": (
                        "草稿章节标题层级序列与样例不完全相同。若本任务在样例章节级别范围内扩写了多个章节或小节，"
                        "可追加 --flexible-chapters（或在复刻样例精确小节结构时调整章节标题层级与样例一致）。"
                    ),
                }
            )
        elif (
            len(missing) == 1
            and len(unexpected) == 1
            and missing[0].get("kind") in {"cover", "title_block"}
            and unexpected[0].get("kind") in {"cover", "title_block"}
        ):
            remediations.append(
                {
                    "issue": "front_matter_page_break_difference",
                    "fix": (
                        "封面/标题块分页状态与样例不一致。若因 Word 更新域注入了 <w:lastRenderedPageBreak/>，"
                        "请运行 python scripts/format_cli.py refresh-fields <草稿.docx> --in-place 自动清理分页缓存标记。"
                    ),
                }
            )
        else:
            remediations.append(
                {
                    "issue": "module_sequence_mismatch",
                    "missing": [
                        f"{m.get('kind')}:{m.get('variant')}" if m.get("variant") else m.get("kind")
                        for m in missing
                    ],
                    "unexpected": [
                        f"{m.get('kind')}:{m.get('variant')}" if m.get("variant") else m.get("kind")
                        for m in unexpected
                    ],
                    "fix": (
                        "请严格按照样例的 ordered_modules 顺序增删模块（不得臆造样例中不存在的承诺书、英文摘要、致谢或附录）。"
                    ),
                }
            )
    return remediations


def diagnose_numbering_remediation(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not issues:
        return []
    seen_kinds: list[str] = []
    for item in issues:
        kind = str(item.get("kind") or "")
        if kind and kind not in seen_kinds:
            seen_kinds.append(kind)
    result = []
    for kind in seen_kinds:
        guide = NUMBERING_KIND_REMEDIATIONS.get(
            kind,
            {
                "cause": "检测到手打编号或缺失编号定义。",
                "fix": "请改用 Word 原生自动编号（w:numPr）或动态域（REF/SEQ/STYLEREF/PAGE）。",
            },
        )
        affected = [i.get("paragraph") for i in issues if i.get("kind") == kind]
        result.append(
            {
                "kind": kind,
                "affected_paragraphs": affected[:20],
                "cause": guide["cause"],
                "fix": guide["fix"],
            }
        )
    return result


def diagnose_verification_remediation(verification: dict[str, Any]) -> list[dict[str, Any]]:
    if verification.get("status") == "passed":
        return []
    remediations: list[dict[str, Any]] = []
    fmt = verification.get("format", {})
    if isinstance(fmt, dict) and fmt.get("status") == "failed":
        failed_checks = [c for c in fmt.get("checks", []) if c.get("status") == "failed"]
        remediations.append(
            {
                "check": "format",
                "failed_count": len(failed_checks),
                "sample_failures": failed_checks[:5],
                "fix": (
                    "目标段落或结构操作未达到预期属性（例如局部空 run、红色格式批注 run 或段落标记 pPr/rPr 不一致，"
                    "或目录/引用域未刷新）。请清理目标段落多余空 run，或使用 format_cli.py cover / refresh-fields 修复。"
                ),
            }
        )
    struct = verification.get("structure", {})
    if isinstance(struct, dict) and struct.get("status") == "failed":
        remediations.append(
            {
                "check": "structure",
                "errors": struct.get("errors", []),
                "fix": (
                    "后置结构核验未通过：若为 table-of-contents 结果无效，请运行 "
                    "python scripts/format_cli.py refresh-fields <输出.docx> --in-place 更新目录域；"
                    "若为模块顺序差异，请对齐样例 ordered_modules。"
                ),
            }
        )
    if verification.get("visual", {}).get("status") in {"failed", "pending"}:
        remediations.append(
            {
                "check": "visual",
                "status": verification["visual"].get("status"),
                "fix": (
                    "请检查渲染目录下的 page-1.png..page-N.png，如发现版面瑕疵可对照 visual_defect_playbook 修复；"
                    "确认无误后使用预生成的 visual-report.template.json 执行 --resume 完成视觉签收。"
                ),
                "visual_defect_playbook": VISUAL_DEFECT_PLAYBOOK,
            }
        )
    return remediations


def build_pipeline_remediation(
    checks: dict[str, str],
    failure: dict[str, Any] | None,
    reports_dir: Path,
    output_path: Path,
) -> dict[str, Any]:
    template_visual = reports_dir / "visual-report.template.json"
    guide: dict[str, Any] = {
        "visual_report_template": str(template_visual) if template_visual.is_file() else None,
        "actions": [],
    }
    if failure:
        stage = failure.get("stage")
        detail = str(failure.get("detail") or failure.get("error") or "")
        if stage == "structure":
            guide["actions"].append(
                "结构核对未通过：请查看 structure.json 中的 remediation 字段；若因 Word 域更新注入了 lastRenderedPageBreak，"
                "运行 python scripts/format_cli.py refresh-fields <输入草稿.docx> --in-place 即可一键修复。"
            )
        elif stage == "plan":
            guide["actions"].append(
                "计划构建需要澄清：请查看 plan.json 中的 issues；若为 section-mapping，在 mapping.json 中映射节号；"
                "若为规则冲突，在 resolutions.json 中使用 choose/replace/ignore 决策（replace 规则会自动覆盖同目标基础规则）。"
            )
        elif stage == "verify":
            guide["actions"].append(
                "格式核验未通过：请查看 verification.json 中的 remediation 字段定位具体属性或复合操作偏差。"
            )
        elif stage == "numbering":
            guide["actions"].append(
                "编号核对未通过：请查看 numbering.json 中的 remediation 字段，将手打章节号/图表题注号/文献标号替换为原生编号或动态域。"
            )
        elif stage == "quality":
            guide["actions"].append(
                f"内容与引用核验未通过：请查看 {reports_dir / 'quality.json'}，核对正文字数、摘要字数、文献总数/近三年数量及正文与参考文献双向引用映射（不得有未引用的悬空文献或缺失文献）。"
            )
        elif stage == "visual":
            guide["actions"].append(
                f"视觉核验未通过（{detail[:160]}）：请先调用图像查看工具逐页检查 {reports_dir / 'render'} 下的 page-*.png，"
                f"确认版面无误后将 {template_visual} 中的 status 从 pending_visual_review 改为 passed 并填写具体检查结论 summary，"
                "再作为 --visual-report 传入并执行 --resume（严禁未看图直接原样提交模板）。"
            )
    if checks.get("visual") == "pending":
        guide["actions"].append(
            f"页面渲染已完成：必须先调用图像查看工具（如 read_file）逐页审阅 {reports_dir / 'render'} 下的全部 page-*.png。"
            f"核对 checklist 各项无误后，编辑 {template_visual} 将 status 从 pending_visual_review 改为 passed 并填写真实 summary，"
            f"再追加 --visual-report \"{template_visual}\" --resume 完成最终签收（直接原样传入未修改的模板会被硬门禁拒绝）。"
        )
        guide["visual_defect_playbook"] = VISUAL_DEFECT_PLAYBOOK
    return guide
