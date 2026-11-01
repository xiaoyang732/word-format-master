"""Universal multi-template compatibility regression tests.

Covers diverse real-world university/journal templates and user drafts:
1. Table-based covers (2-col & 4-col, colonless labels, vMerge multi-row slots, custom keys)
2. Single-section templates without cover section breaks, underlined/underscore slots,
   multi-field lines, standalone title placeholders, and CLI `format_cli.py cover`
3. Inline abstracts (`abstract_inline` label/value segments), colon subtitles, 2-line titles,
   expanded module markers (`主要符号表`, `引用文献`, `谢辞`, `主题词`, `Index Terms`),
   and immunity against false positives on chapter prose (`Abstract classes...`)
4. TOC insertion anchor across bilingual `abstract_title` lines (ZH->EN and EN->ZH)
5. Preserving black/auto color when `Heading N` is assigned to `Normal` headings
   under written requirements without explicit color clauses
6. Read-only legacy `.doc` / `.dot` / `.wps` loading for templates, requirements, and covers
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from word_format.api import apply_plan, build_plan, verify_plan
from word_format.contracts import FormatError, sha256
from word_format.cover import clone_template_cover
from word_format.inspect import DocumentIndex
from word_format.profiles import requirements_profile, template_profile
from word_format.properties import read_font
from word_format.templates import compile_template, template_plan


class UniversalTemplateCompatTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_table_based_cover_without_colons_vmerge_and_custom_fields(self):
        tpl_path = self.root / "table_cover_template.docx"
        tpl = Document()
        normal = tpl.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(12)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

        banner = tpl.add_paragraph("某某大学本科毕业设计说明书")
        banner.runs[0].bold = True
        banner.runs[0].font.size = Pt(22)
        banner.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")

        # 4-column table: Row 0 & Row 1 form a 2-line vMerge title box; Row 2 has two fields!
        tbl = tpl.add_table(rows=4, cols=4)
        # Row 0: [论文题目 | 样例第一行题目 (gridSpan 3)]
        tbl.cell(0, 0).text = "论文题目"
        tbl.cell(0, 1).merge(tbl.cell(0, 3))
        tbl.cell(0, 1).text = "样例旧题目第一行"
        tbl.cell(0, 0)._tc.get_or_add_tcPr().append(
            parse_xml(f'<w:vMerge {nsdecls("w")} w:val="restart"/>')
        )
        tbl.cell(0, 1)._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="6"/></w:tcBorders>'
            )
        )
        # Row 1: [vMerge continue | 样例旧题目第二行 (gridSpan 3)]
        tbl.cell(1, 0).text = ""
        tbl.cell(1, 1).merge(tbl.cell(1, 3))
        tbl.cell(1, 1).text = "——样例旧题目第二行"
        tbl.cell(1, 0)._tc.get_or_add_tcPr().append(parse_xml(f'<w:vMerge {nsdecls("w")}/>'))
        tbl.cell(1, 1)._tc.get_or_add_tcPr().append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="6"/></w:tcBorders>'
            )
        )
        # Row 2: 4-col [学生姓名 | 张三 | 学    号 | 20200001]
        tbl.cell(2, 0).text = "学生姓名"
        tbl.cell(2, 1).text = "张三"
        tbl.cell(2, 2).text = "学    号"
        tbl.cell(2, 3).text = "20200001"
        # Row 3: 4-col [院（系） | 旧学院 | 课程名称 | 旧课程] (includes custom field key!)
        tbl.cell(3, 0).text = "院（系）"
        tbl.cell(3, 1).text = "旧学院"
        tbl.cell(3, 2).text = "课程名称"
        tbl.cell(3, 3).text = "旧课程"

        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(14)
                        r.font.name = "Times New Roman"
                        r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "仿宋")

        h1 = tpl.add_heading("绪论", level=1)
        h1.runs[0].font.size = Pt(16)
        h1.runs[0].font.name = "Times New Roman"
        h1.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        p_body = tpl.add_paragraph("样例第一章正文。")
        p_body.runs[0].font.size = Pt(12)
        tpl.save(tpl_path)

        # Clone cover from single-section table template (must NOT clone 绪论!)
        draft_path = self.root / "draft.docx"
        out_path = self.root / "formatted.docx"
        draft = clone_template_cover(
            tpl_path,
            fields={
                "title": "分布式仓储调度系统的设计与实现",
                "student": "王潇杨",
                "student_id": "2305010628",
                "college": "大数据与计算机科学学院",
                "课程名称": "企业级Java开发综合实践",
            },
        )
        draft.add_heading("第1章 系统总体架构", level=1)
        draft.add_paragraph("这是目标论文自己的第一章正文。")
        draft.save(draft_path)

        cloned_reopened = Document(draft_path)
        all_cover_cells = [
            c.text.strip() for row in cloned_reopened.tables[0].rows for c in row.cells
        ]
        self.assertIn("分布式仓储调度系统的设计与实现", all_cover_cells)
        self.assertIn("王潇杨", all_cover_cells)
        self.assertIn("2305010628", all_cover_cells)
        self.assertIn("大数据与计算机科学学院", all_cover_cells)
        self.assertIn("企业级Java开发综合实践", all_cover_cells)
        self.assertNotIn("样例旧题目第一行", all_cover_cells)
        self.assertNotIn("——样例旧题目第二行", all_cover_cells)
        self.assertNotIn("样例第一章正文", "".join(p.text for p in cloned_reopened.paragraphs))
        # Bottom border on Row 1 Cell 1 is preserved even though continuation text was cleared
        self.assertTrue(
            cloned_reopened.tables[0].cell(1, 1)._tc.xpath("./w:tcPr/w:tcBorders/w:bottom")
        )

        # Profile & compile template from tpl_path and apply cleanly onto draft_path
        profile = template_profile(tpl_path)
        resolutions = {
            f"s1.{i['id']}": {"decision": "ignore", "reason": "Keep target defaults for test"}
            for i in profile["issues"]
        }
        compiled = compile_template([profile], identifier="table-cover-uni", resolutions=resolutions)
        plan = template_plan(draft_path, out_path, compiled, mapping={"sections": {"1": 2}})
        self.assertNotIn("status", plan, msg=f"Unexpected review issues: {plan.get('issues')}")
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

    def test_single_section_cover_underlines_multifield_and_cli_replace_cover(self):
        tpl_path = self.root / "underline_cover_template.docx"
        tpl = Document()
        p_banner = tpl.add_paragraph("本科毕业论文")
        p_banner.runs[0].font.size = Pt(24)

        p_title_ph = tpl.add_paragraph("（在此输入论文题目）")
        p_title_ph.runs[0].font.size = Pt(18)
        p_title_ph.runs[0].bold = True

        # Multi-field paragraph with ASCII underscore blanks
        p_multi = tpl.add_paragraph()
        r_m1 = p_multi.add_run("姓  名：")
        r_m1.bold = True
        p_multi.add_run("__________")
        r_m2 = p_multi.add_run("    学  号：")
        r_m2.bold = True
        p_multi.add_run("____________")

        # Styled underline slot: bold label + non-bold underlined padded value
        p_adv = tpl.add_paragraph()
        r_lbl = p_adv.add_run("指导教师：")
        r_lbl.bold = True
        r_val = p_adv.add_run("    旧导师    ")
        r_val.bold = False
        r_val.font.underline = True

        p_date = tpl.add_paragraph("2024 年 6 月")
        # Separate cover from abstract ONLY with a page break (no section break!)
        p_date.add_run().add_break(
            __import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE
        )
        tpl.add_paragraph("摘  要")
        tpl.add_paragraph("模板中的旧摘要正文，绝不能被克隆进封面。")
        tpl.save(tpl_path)
        tpl_hash = sha256(tpl_path)

        # Existing draft already has a plain cover that should be replaced
        draft_path = self.root / "existing_draft.docx"
        draft = Document()
        draft.add_paragraph("简陋旧封面标题")
        draft.add_paragraph("姓名：旧姓名")
        draft.add_page_break()
        draft.add_paragraph("摘  要")
        draft.add_paragraph("这是保留的用户真实摘要正文。")
        draft.add_heading("第1章 绪论", level=1)
        draft.add_paragraph("这是保留的用户真实第一章正文。")
        draft.save(draft_path)
        draft_hash = sha256(draft_path)

        fields_path = self.root / "cover_fields.json"
        fields_path.write_text(
            json.dumps(
                {
                    "title": "高校智能排课算法研究",
                    "姓名": "李明",
                    "学号": "20260809",
                    "指导教师": "赵教授",
                    "date": "二〇二六年六月",
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        out_path = self.root / "with_official_cover.docx"
        report_path = self.root / "cover_report.json"

        # Safety check: attempting to overwrite template_path via CLI must fail
        bad = subprocess.run(
            [
                sys.executable,
                str(ROOT / "scripts/format_cli.py"),
                "cover",
                str(tpl_path),
                str(tpl_path),
            ],
            capture_output=True,
            text=True,
        )
        self.assertNotEqual(bad.returncode, 0)
        self.assertEqual(sha256(tpl_path), tpl_hash)

        res = subprocess.run(
            [
                sys.executable,
                str(ROOT / "scripts/format_cli.py"),
                "cover",
                str(tpl_path),
                str(out_path),
                "--into",
                str(draft_path),
                "--replace-existing-cover",
                "--fields",
                str(fields_path),
                "--report",
                str(report_path),
            ],
            capture_output=True,
            text=True,
        )
        self.assertEqual(res.returncode, 0, msg=res.stderr)
        self.assertEqual(sha256(tpl_path), tpl_hash)
        self.assertEqual(sha256(draft_path), draft_hash)

        out = Document(out_path)
        texts = [p.text for p in out.paragraphs if p.text.strip()]
        joined = "\n".join(texts)
        self.assertIn("高校智能排课算法研究", joined)
        self.assertIn("李明", joined)
        self.assertIn("20260809", joined)
        self.assertIn("赵教授", joined)
        self.assertIn("二〇二六年六月", joined)
        self.assertIn("这是保留的用户真实摘要正文。", joined)
        self.assertIn("这是保留的用户真实第一章正文。", joined)
        self.assertNotIn("简陋旧封面标题", joined)
        self.assertNotIn("旧姓名", joined)
        self.assertNotIn("模板中的旧摘要正文", joined)

        # Verify underline & non-bold run preservation on 指导教师 and multi-field line
        adv_p = next(p for p in out.paragraphs if "指导教师：" in p.text)
        val_runs = [r for r in adv_p.runs if "赵教授" in r.text]
        self.assertEqual(len(val_runs), 1)
        self.assertFalse(bool(val_runs[0].bold))
        self.assertIsNotNone(val_runs[0]._r.rPr.find(qn("w:u")))

        multi_p = next(p for p in out.paragraphs if "姓  名：" in p.text and "学  号：" in p.text)
        name_runs = [r for r in multi_p.runs if "李明" in r.text]
        id_runs = [r for r in multi_p.runs if "20260809" in r.text]
        self.assertEqual(len(name_runs), 1)
        self.assertEqual(len(id_runs), 1)
        self.assertIsNotNone(name_runs[0]._r.rPr.find(qn("w:u")))
        self.assertIsNotNone(id_runs[0]._r.rPr.find(qn("w:u")))

    def test_inline_abstract_segments_colon_subtitles_and_expanded_module_markers(self):
        sample_path = self.root / "journal_sample.docx"
        sample = Document()
        normal = sample.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(10.5)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

        def styled_p(doc, text, *, ea="宋体", latin="Times New Roman", size=10.5, bold=False):
            p = doc.add_paragraph(text)
            r = p.runs[0]
            r.font.name = latin
            r.font.size = Pt(size)
            r.bold = bold
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), ea)
            return p

        # Chinese title with colon subtitle
        styled_p(sample, "智慧物流调度方法：基于多智能体强化学习", ea="黑体", size=16, bold=True)
        # Inline Chinese abstract: bold 黑体 label + regular 楷体 value
        p_abs_zh = sample.add_paragraph()
        r_az_lbl = p_abs_zh.add_run("【摘 要】")
        r_az_lbl.bold = True
        r_az_lbl.font.size = Pt(10.5)
        r_az_lbl.font.name = "Times New Roman"
        r_az_lbl._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        r_az_val = p_abs_zh.add_run("本文提出一种面向大规模仓储的多智能体调度方法。")
        r_az_val.bold = False
        r_az_val.font.size = Pt(10.5)
        r_az_val.font.name = "Times New Roman"
        r_az_val._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "楷体")

        # Variant keywords: 【主题词】
        p_kw_zh = sample.add_paragraph()
        r_kz_lbl = p_kw_zh.add_run("【主题词】")
        r_kz_lbl.bold = True
        r_kz_lbl.font.size = Pt(10.5)
        r_kz_lbl.font.name = "Times New Roman"
        r_kz_lbl._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        r_kz_val = p_kw_zh.add_run("智慧物流；强化学习")
        r_kz_val.bold = False
        r_kz_val.font.size = Pt(10.5)
        r_kz_val.font.name = "Times New Roman"
        r_kz_val._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "楷体")

        # 2-line English title + colon subtitle before inline English Abstract
        styled_p(sample, "Intelligent Logistics Scheduling", size=14, bold=True)
        styled_p(sample, "— Multi-Agent Learning: Theory and Practice", size=14, bold=True)

        p_abs_en = sample.add_paragraph()
        r_ae_lbl = p_abs_en.add_run("Abstract: ")
        r_ae_lbl.bold = True
        r_ae_lbl.font.size = Pt(10.5)
        r_ae_lbl.font.name = "Times New Roman"
        r_ae_lbl._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
        r_ae_val = p_abs_en.add_run(
            "This paper studies multi-agent scheduling in Shanxi (山西) logistics hubs."
        )
        r_ae_val.bold = False
        r_ae_val.font.size = Pt(10.5)
        r_ae_val.font.name = "Times New Roman"
        r_ae_val._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")

        p_kw_en = sample.add_paragraph()
        r_ke_lbl = p_kw_en.add_run("Index Terms: ")
        r_ke_lbl.bold = True
        r_ke_lbl.font.size = Pt(10.5)
        r_ke_lbl.font.name = "Times New Roman"
        r_ke_lbl._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
        r_ke_val = p_kw_en.add_run("logistics, reinforcement learning")
        r_ke_val.bold = False
        r_ke_val.font.size = Pt(10.5)
        r_ke_val.font.name = "Times New Roman"
        r_ke_val._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")

        styled_p(sample, "主要符号表", ea="黑体", size=14, bold=True)
        styled_p(sample, "α —— 学习率参数", ea="宋体", size=10.5, bold=False)

        h1 = sample.add_heading("1 系统模型", level=1)
        h1.runs[0].font.size = Pt(14)
        h1.runs[0].font.name = "Times New Roman"
        h1.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        styled_p(
            sample,
            "Abstract classes in Java provide base implementations while Keywords in Java are reserved.",
            ea="宋体",
            size=10.5,
        )
        styled_p(sample, "引用文献", ea="黑体", size=14, bold=True)
        styled_p(sample, "[1] Sutton R S. Reinforcement Learning[M]. MIT Press, 2018.", size=9)
        styled_p(sample, "谢 辞", ea="黑体", size=14, bold=True)
        styled_p(sample, "感谢课题组全体成员。", size=10.5)
        sample.save(sample_path)

        # Target unformatted draft using colon inline abstract ("摘 要：...") and 2-line EN title
        draft_path = self.root / "journal_draft.docx"
        out_path = self.root / "journal_out.docx"
        draft = Document()
        for line in (
            "仓库管理系统重构：基于领域驱动设计",
            "摘 要：本文围绕仓储核心领域建模展开系统重构与性能优化。",
            "关键词：领域驱动设计；仓库管理",
            "Warehouse Management Refactoring",
            "— Domain-Driven Design: Architecture and Practice",
            "Abstract: This paper refactors warehouse software for Shanxi (山西) enterprises.",
            "Index Terms: domain-driven design, warehouse",
            "缩略语表",
            "DDD —— 领域驱动设计",
            "1 领域模型设计",
            "Abstract classes in Java define domain contracts.",
            "引用文献",
            "[1] Evans E. Domain-Driven Design[M]. Addison-Wesley, 2003.",
            "后 记",
            "感谢审稿专家提出宝贵建议。",
        ):
            draft.add_paragraph(line)
        draft.save(draft_path)

        profile = template_profile(sample_path)
        resolutions = {
            f"s1.{i['id']}": {"decision": "ignore", "reason": "Test ignores header/numbering issues"}
            for i in profile["issues"]
        }
        compiled = compile_template([profile], identifier="journal-inline", resolutions=resolutions)
        plan = template_plan(draft_path, out_path, compiled)
        self.assertNotIn("status", plan, msg=f"Unexpected review issues: {plan.get('issues')}")
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(out_path)
        out_idx = DocumentIndex(out)

        # Check colon subtitle & 2-line EN title got abstract_title
        abs_titles = [
            (r["text"], r["language"])
            for r in out_idx.records.values()
            if r.get("semantic_role") == "abstract_title"
        ]
        self.assertEqual(
            abs_titles,
            [
                ("仓库管理系统重构：基于领域驱动设计", "zh"),
                ("Warehouse Management Refactoring", "en"),
                ("— Domain-Driven Design: Architecture and Practice", "en"),
            ],
        )

        # Check Chinese inline abstract split into bold 黑体 label ("摘 要：") + non-bold 楷体 value!
        zh_inline = next(
            r
            for r in out_idx.records.values()
            if r.get("semantic_role") == "abstract_inline" and r.get("language") == "zh"
        )
        self.assertEqual(out_idx.font_values(zh_inline["id"], "font_east_asia", 0, 4), ["黑体"])
        self.assertEqual(out_idx.font_values(zh_inline["id"], "bold", 0, 4), [True])
        self.assertEqual(out_idx.font_values(zh_inline["id"], "font_east_asia", 4, 20), ["楷体"])
        self.assertEqual(out_idx.font_values(zh_inline["id"], "bold", 4, 20), [False])

        # English inline abstract with Chinese proper noun "(山西)" is still language='en'
        en_inline = next(
            r
            for r in out_idx.records.values()
            if r.get("semantic_role") == "abstract_inline" and r.get("language") == "en"
        )
        self.assertEqual(out_idx.font_values(en_inline["id"], "bold", 0, 9), [True])
        self.assertEqual(out_idx.font_values(en_inline["id"], "bold", 10, 30), [False])

        # Chapter prose starting with "Abstract classes..." remained body_text
        prose_rec = next(
            r
            for r in out_idx.records.values()
            if r.get("kind") == "paragraph" and "Abstract classes in Java" in r.get("text", "")
        )
        self.assertEqual(prose_rec["semantic_role"], "body_text")
        self.assertEqual(prose_rec["module"], "chapters")

    def test_toc_insertion_skips_bilingual_abstract_titles_in_both_orders(self):
        # Order 1: Chinese abstract -> 2-line English title (with colon) -> English Abstract -> Ch 1
        draft_zh_first = self.root / "zh_first.docx"
        out_zh_first = self.root / "zh_first_out.docx"
        doc1 = Document()
        for line in (
            "智慧仓储管理系统设计：基于微服务架构",
            "摘  要",
            "中文摘要内容。",
            "关键词：仓储；微服务",
            "Smart Warehouse Management System",
            "— Microservice Architecture: Design and Implementation",
            "Abstract",
            "English abstract content.",
            "Key words: warehouse; microservice",
        ):
            doc1.add_paragraph(line)
        doc1.add_heading("第1章 绪论", level=1)
        doc1.add_paragraph("第一章正文。")
        doc1.save(draft_zh_first)

        req = {
            "operations": [
                {
                    "action": "toc.configure",
                    "target": {"type": "document"},
                    "params": {"enabled": True, "title": "目 录", "max_heading_level": 3},
                }
            ]
        }
        plan1 = build_plan(draft_zh_first, out_zh_first, req)
        rep1 = apply_plan(plan1)
        self.assertEqual(verify_plan(plan1, rep1)["status"], "passed")
        out1 = Document(out_zh_first)
        texts1 = [p.text.strip() for p in out1.paragraphs if p.text.strip()]
        self.assertGreater(texts1.index("目 录"), texts1.index("Key words: warehouse; microservice"))
        self.assertLess(texts1.index("目 录"), texts1.index("第1章 绪论"))

        # Order 2: English Abstract first -> Chinese title -> Chinese 摘要 -> Chapter 1
        draft_en_first = self.root / "en_first.docx"
        out_en_first = self.root / "en_first_out.docx"
        doc2 = Document()
        for line in (
            "Smart Warehouse System: A Study",
            "Abstract",
            "English abstract first.",
            "Key words: warehouse; study",
            "智慧仓储系统研究：架构分析",
            "摘  要",
            "后置中文摘要正文。",
            "关键词：仓储；架构",
        ):
            doc2.add_paragraph(line)
        doc2.add_heading("Chapter 1 Introduction", level=1)
        doc2.add_paragraph("Chapter body.")
        doc2.save(draft_en_first)

        plan2 = build_plan(draft_en_first, out_en_first, req)
        rep2 = apply_plan(plan2)
        self.assertEqual(verify_plan(plan2, rep2)["status"], "passed")
        out2 = Document(out_en_first)
        texts2 = [p.text.strip() for p in out2.paragraphs if p.text.strip()]
        self.assertGreater(texts2.index("目 录"), texts2.index("关键词：仓储；架构"))
        self.assertLess(texts2.index("目 录"), texts2.index("Chapter 1 Introduction"))

    def test_normal_headings_with_written_requirements_do_not_turn_accent_blue(self):
        draft_path = self.root / "plain_headings.docx"
        out_path = self.root / "styled_headings.docx"
        doc = Document()
        # Authors frequently write headings as Normal paragraphs without Heading 1/2 styles
        doc.add_paragraph("第1章 绪论")
        doc.add_paragraph("正文段落一。")
        doc.add_paragraph("1.1 研究背景")
        doc.add_paragraph("正文段落二。")
        doc.save(draft_path)
        normal_before_xml = doc.styles["Normal"].element.xml

        req_txt = self.root / "req.txt"
        req_txt.write_text(
            "一级标题黑体，三号，居中。二级标题黑体，四号，左对齐。正文宋体，小四。",
            encoding="utf-8",
        )
        req_prof = requirements_profile(req_txt)
        compiled = compile_template([req_prof], identifier="written-req-no-color")
        compiled["rules"].append(
            {
                "id": "scheme",
                "state": "set",
                "action": "heading.scheme.configure",
                "target": {"type": "document"},
                "params": {
                    "levels": [
                        {"format": "decimal", "text": "第%1章", "start": 1, "suffix": "space"},
                        {"format": "decimal", "text": "%1.%2", "start": 1, "suffix": "space"},
                    ]
                },
                "evidence": [{"anchor": "user:scheme"}],
            }
        )
        plan = template_plan(draft_path, out_path, compiled)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(out_path)
        self.assertEqual(out.styles["Normal"].element.xml, normal_before_xml)
        h1_p = out.paragraphs[0]
        h2_p = out.paragraphs[2]
        self.assertEqual(h1_p.style.name, "Heading 1")
        self.assertEqual(h2_p.style.name, "Heading 2")
        # Neither heading text nor paragraph-mark numbering rPr may become Office blue (365F91/4F81BD)
        for hp in (h1_p, h2_p):
            self.assertEqual(read_font(out, hp._p, None, "color"), "auto")
            for r in hp.runs:
                self.assertEqual(read_font(out, hp._p, r._r, "color"), "auto")

    def test_legacy_doc_and_wps_readonly_loader_for_profile_requirements_and_cover(self):
        real_docx = self.root / "backing.docx"
        backing = Document()
        backing.styles["Normal"].font.name = "Times New Roman"
        backing.styles["Normal"].font.size = Pt(12)
        backing.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        p_cov = backing.add_paragraph("题  目：旧题目")
        p_cov.runs[0].font.size = Pt(16)
        p_cov.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        backing.add_page_break()
        h1 = backing.add_heading("第1章 绪论", level=1)
        h1.runs[0].font.size = Pt(16)
        h1.runs[0].font.name = "Times New Roman"
        h1.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        backing.add_paragraph("正文小四。")
        backing.save(real_docx)

        legacy_doc = self.root / "school_template.doc"
        legacy_wps = self.root / "school_requirements.wps"
        # Fake OLE2 binary headers so zipfile.is_zipfile is False and converter is invoked
        legacy_doc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1LEGACY-DOC-CONTENT")
        legacy_wps.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1LEGACY-WPS-REQ")
        doc_hash = sha256(legacy_doc)
        wps_hash = sha256(legacy_wps)

        req_backing = self.root / "req_backing.docx"
        req_doc = Document()
        req_doc.add_paragraph("正文宋体，小四，1.5倍行距。")
        req_doc.save(req_backing)

        seen_staged_paths: list[Path] = []

        def fake_converter(staged_source: Path, output_docx: Path) -> None:
            seen_staged_paths.append(staged_source)
            self.assertNotEqual(staged_source.resolve(), legacy_doc.resolve())
            self.assertNotEqual(staged_source.resolve(), legacy_wps.resolve())
            src_to_copy = req_backing if staged_source.suffix.lower() == ".wps" else real_docx
            shutil.copy2(src_to_copy, output_docx)

        with patch(
            "word_format.doc_loader._LEGACY_CACHE_DIR",
            self.root / "isolated_doc_cache",
        ), patch("word_format.doc_loader._convert_legacy_word_file", side_effect=fake_converter):
            t_prof = template_profile(legacy_doc)
            self.assertEqual(t_prof["source"]["path"], str(legacy_doc.resolve()))
            self.assertEqual(t_prof["source"]["sha256"], doc_hash)
            self.assertTrue(t_prof["rules"])

            r_prof = requirements_profile(legacy_wps)
            self.assertEqual(r_prof["source"]["path"], str(legacy_wps.resolve()))
            self.assertEqual(r_prof["source"]["sha256"], wps_hash)
            self.assertTrue(
                any(
                    r["action"] == "font.size.set" and r["params"]["value"] == 12
                    for r in r_prof["rules"]
                )
            )

            cloned = clone_template_cover(legacy_doc, fields={"title": "新转换封面题目"})
            out_cover = self.root / "from_doc_cover.docx"
            cloned.save(out_cover)
            reopened = Document(out_cover)
            self.assertIn("新转换封面题目", "\n".join(p.text for p in reopened.paragraphs))
            self.assertNotIn("正文小四", "\n".join(p.text for p in reopened.paragraphs))

        # Originals must remain 100% byte-identical
        self.assertEqual(sha256(legacy_doc), doc_hash)
        self.assertEqual(sha256(legacy_wps), wps_hash)
        self.assertTrue(seen_staged_paths)

        # When no converter is installed, raise FormatError(..., status="unsupported")
        uncached_doc = self.root / "no_converter.doc"
        uncached_doc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1UNCACHED-DOC")
        with patch(
            "word_format.doc_loader.detect_renderers",
            return_value={"word": {"available": False}, "libreoffice": {"available": False}},
        ), patch(
            "word_format.doc_loader._word_com_to_docx",
            side_effect=RuntimeError("No COM"),
        ):
            with self.assertRaises(FormatError) as ctx:
                template_profile(uncached_doc)
            self.assertEqual(ctx.exception.status, "unsupported")


if __name__ == "__main__":
    unittest.main()
