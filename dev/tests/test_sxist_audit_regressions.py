"""Regression tests for the 5 user-reported thesis formatting defects:
1. Official template cover cloning (VML textboxes, media relationships, field updates)
2. Blue automatic heading/list numbers when text runs were already black
3. Standalone semantic role extraction & TOC isolation for 摘要 / Abstract / 诚信声明 / 参考文献 / 致谢
4. Standalone page breaks (page_break_before) for 参考文献 and 致谢
5. East Asian theme font resolution when w:themeFontLang is ja-JP (ＭＳ 明朝 vs 宋体)
"""

from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

from word_format.api import apply_plan, audit_document, build_plan, verify_plan
from word_format.contracts import sha256
from word_format.cover import clone_template_cover
from word_format.inspect import DocumentIndex
from word_format.profiles import requirements_profile, template_profile
from word_format.properties import outline_level, read_font, read_paragraph
from word_format.templates import compile_template, load_template, template_plan
from word_format.toc import headings as toc_headings, result_status


class SxistAuditRegressionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "source.docx"
        self.output = self.root / "output.docx"

    def test_ja_jp_theme_font_ms_mincho_is_detected_and_overwritten_by_songti(self):
        """Issue 5: python-docx default.docx sets w:themeFontLang eastAsia='ja-JP' (ＭＳ 明朝)."""
        doc = Document()
        doc.add_paragraph("仓库管理系统的设计与实现")
        doc.save(self.source)

        # Inspect must truthfully report ＭＳ 明朝 from theme1.xml + settings.xml, never 宋体.
        index = DocumentIndex(Document(self.source))
        self.assertEqual(index.font_values("p0", "font_east_asia"), ["ＭＳ 明朝"])

        request = {
            "operations": [
                {
                    "action": "font.east_asia.set",
                    "target": {"type": "paragraph", "id": "p0"},
                    "params": {"value": "宋体"},
                }
            ]
        }
        self.assertEqual(audit_document(self.source, request)["status"], "failed")

        plan = build_plan(self.source, self.output, request)
        report = apply_plan(plan)
        self.assertTrue(report["checks"][0]["changed"])
        self.assertEqual(verify_plan(plan, report)["status"], "passed")

        out = Document(self.output)
        self.assertEqual(DocumentIndex(out).font_values("p0", "font_east_asia"), ["宋体"])
        rfonts = out.paragraphs[0].runs[0]._r.rPr.rFonts
        self.assertEqual(rfonts.get(qn("w:eastAsia")), "宋体")

    def test_blue_heading_numbering_with_black_text_runs_syncs_paragraph_mark_rpr(self):
        """Issue 2: Pre-colored black text runs must not prevent pPr/rPr from turning black."""
        doc = Document()
        for level, title in ((1, "第1章 绪论"), (2, "1.1 研究背景"), (3, "1.1.1 核心问题")):
            h = doc.add_heading("", level=level)
            run = h.add_run(title)
            run.font.color.rgb = RGBColor(0, 0, 0)
        doc.add_paragraph("正文内容。")
        doc.save(self.source)

        # Before formatting, Heading 1/2/3 paragraph-mark rPr inherits Office accent1 blue.
        raw_doc = Document(self.source)
        self.assertNotEqual(read_font(raw_doc, raw_doc.paragraphs[0]._p, None, "color"), "000000")
        self.assertEqual(read_font(raw_doc, raw_doc.paragraphs[1]._p, None, "color"), "4F81BD")

        template, authority = load_template("thesis-standard")
        plan = template_plan(self.source, self.output, template, template_source=authority)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(self.output)
        for p in out.paragraphs:
            if p.style.name.startswith("Heading "):
                self.assertEqual(read_font(out, p._p, None, "color"), "000000")
                self.assertEqual(read_font(out, p._p, None, "font_east_asia"), "黑体")
                self.assertTrue(read_font(out, p._p, None, "bold"))
                ppr_color = p._p.find("./" + qn("w:pPr") + "/" + qn("w:rPr") + "/" + qn("w:color"))
                self.assertIsNotNone(ppr_color)
                self.assertEqual(ppr_color.get(qn("w:val")), "000000")
                self.assertIsNone(ppr_color.get(qn("w:themeColor")))

    def test_standalone_abstract_and_module_headings_isolated_from_body_and_toc(self):
        """Issue 3: 摘要/Abstract/诚信声明/参考文献/致谢 use standalone roles and never pollute TOC."""
        sample_path = self.root / "sample_template.docx"
        sample = Document()
        normal = sample.styles["Normal"]
        normal.font.name = "Times New Roman"
        normal.font.size = Pt(12)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

        def add_styled(doc, text, *, ea="宋体", latin="Times New Roman", size=12, bold=False):
            p = doc.add_paragraph(text)
            r = p.runs[0]
            r.font.name = latin
            r.font.size = Pt(size)
            r.bold = bold
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), ea)
            return p

        add_styled(sample, "诚信声明", ea="宋体", size=16, bold=True)
        add_styled(sample, "本人郑重声明：", ea="宋体", size=12, bold=False)
        add_styled(sample, "本论文由本人独立完成。", ea="宋体", size=12, bold=False)
        add_styled(sample, "本人签名：   2026 年月日", ea="宋体", size=14, bold=False)
        add_styled(sample, "样例系统设计与实现", ea="黑体", size=16, bold=False)
        add_styled(sample, "摘  要", ea="黑体", size=16, bold=False)
        add_styled(sample, "样例中文摘要正文。", ea="宋体", size=12, bold=False)
        kw_zh = add_styled(sample, "关键词：", ea="黑体", size=12, bold=False)
        kw_zh.add_run("仓储；管理").bold = False
        add_styled(sample, "Sample System Design", ea="宋体", size=16, bold=True)
        add_styled(sample, "Abstract", ea="宋体", size=16, bold=True)
        add_styled(sample, "Sample English abstract prose.", ea="宋体", size=12, bold=False)
        kw_en = add_styled(sample, "Key words: ", ea="宋体", size=12, bold=True)
        kw_en.add_run("warehouse; JavaEE").bold = False
        add_styled(sample, "目 录", ea="宋体", size=16, bold=True)
        h1 = sample.add_heading("绪论", level=1)
        h1.runs[0].font.name = "Times New Roman"
        h1.runs[0].font.size = Pt(16)
        h1.runs[0].bold = True
        h1.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
        add_styled(sample, "第一章正文。", ea="宋体", size=12, bold=False)
        ref_p = add_styled(sample, "参考文献", ea="宋体", size=16, bold=True)
        ref_p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        add_styled(sample, "[1] 示例文献.", ea="宋体", size=12, bold=False)
        ack_p = add_styled(sample, "致谢", ea="宋体", size=16, bold=True)
        ack_p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        add_styled(sample, "感谢指导老师。", ea="宋体", size=12, bold=False)
        sample.save(sample_path)

        # Build unformatted target draft where everything starts as 10.5pt MS Mincho
        target = Document()
        for line in (
            "诚信声明",
            "本人郑重声明：",
            "所呈交的毕业设计由本人独立完成。",
            "本人签名：   2026 年月日",
            "基于JavaEE的仓库管理系统的设计与实现",
            "摘  要",
            "本文设计并实现了仓库管理系统。",
            "关键词：仓库管理，JavaEE",
            "Design and Implementation of Warehouse Management System",
            "Abstract",
            "This thesis designs a warehouse management system.",
            "Key words: Warehouse, JavaEE",
            "目录",
            "第1章 绪论",
            "仓库管理正文段落。",
            "参考文献",
            "[1] 王珊. 数据库系统概论[M]. 北京: 高等教育出版社, 2023.",
            "致谢",
            "谨向导师致以诚挚谢意。",
        ):
            if line == "摘  要":
                target.add_heading(line, level=1)
            else:
                target.add_paragraph(line)
        target.save(self.source)

        profile = template_profile(sample_path)
        resolutions = {
            f"s1.{issue['id']}": {"decision": "ignore", "reason": "Test fixture ignores empty headers"}
            for issue in profile["issues"]
        }
        compiled = compile_template([profile], identifier="sxist-sample", resolutions=resolutions)
        compiled["rules"].append(
            {
                "id": "toc-enable",
                "state": "set",
                "action": "toc.configure",
                "target": {"type": "document"},
                "params": {"enabled": True, "title": "目 录", "max_heading_level": 3},
                "evidence": [{"anchor": "user:toc"}],
            }
        )
        plan = template_plan(self.source, self.output, compiled)
        self.assertNotIn("status", plan)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(self.output)
        out_index = DocumentIndex(out)
        zh_abs = next(
            r
            for r in out_index.records.values()
            if r.get("semantic_role") == "abstract_heading" and r.get("language") == "zh"
        )
        en_abs = next(
            r
            for r in out_index.records.values()
            if r.get("semantic_role") == "abstract_heading" and r.get("language") == "en"
        )
        decl_h = next(
            r
            for r in out_index.records.values()
            if r.get("semantic_role") == "declaration_heading"
        )
        chap_h = next(
            r for r in out_index.records.values() if r.get("semantic_role") == "chapter_heading"
        )

        # "摘  要" is 16pt 黑体 NOT bold, distinct from Chapter 1 (16pt 宋体 bold) and Body (12pt 宋体)
        self.assertEqual(out_index.font_values(zh_abs["id"], "font_east_asia"), ["黑体"])
        self.assertEqual(out_index.font_values(zh_abs["id"], "font_size_pt"), [16.0])
        self.assertEqual(out_index.font_values(zh_abs["id"], "bold"), [False])
        self.assertEqual(out_index.font_values(en_abs["id"], "bold"), [True])
        self.assertEqual(out_index.font_values(chap_h["id"], "font_east_asia"), ["宋体"])
        self.assertEqual(out_index.font_values(chap_h["id"], "bold"), [True])

        # Front-matter headings get body outline_level (9 / None) and never appear in TOC
        for front in (decl_h, zh_abs, en_abs):
            self.assertEqual(read_paragraph(out, out_index.elements[front["id"]], "outline_level"), 9)
            self.assertIn(outline_level(out, out_index.elements[front["id"]]), {None, 9})
        indexed_titles = [p.text.strip() for p, _ in toc_headings(out, 3)]
        self.assertEqual(indexed_titles, ["第1章 绪论", "参考文献", "致谢"])
        self.assertTrue(result_status(out, 3)["valid"])

    def test_references_and_acknowledgements_start_on_standalone_pages(self):
        """Issue 4: 参考文献 and 致谢 must each have page_break_before=True."""
        doc = Document()
        doc.add_heading("第一章 绪论", level=1)
        doc.add_paragraph("正文。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 文献条目.")
        doc.add_paragraph("致谢")
        doc.add_paragraph("致谢正文。")
        doc.save(self.source)

        template, authority = load_template("thesis-standard")
        plan = template_plan(self.source, self.output, template, template_source=authority)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(self.output)
        by_text = {p.text.strip(): p for p in out.paragraphs if p.text.strip()}
        self.assertTrue(read_paragraph(out, by_text["参考文献"]._p, "page_break_before"))
        self.assertTrue(read_paragraph(out, by_text["致谢"]._p, "page_break_before"))

        # Written requirement parser also recognizes standalone page-break clauses
        req_file = self.root / "req.txt"
        req_file.write_text("参考文献另起一页。致谢单独开一页。", encoding="utf-8")
        req_profile = requirements_profile(req_file)
        self.assertEqual(req_profile["issues"], [])
        targets = {
            r["target"].get("semantic_role"): r["params"]["value"]
            for r in req_profile["rules"]
            if r["action"] == "paragraph.page_break_before.set"
        }
        self.assertEqual(
            targets,
            {"reference_heading": True, "acknowledgements_heading": True},
        )

    def test_clone_template_cover_preserves_vml_and_media_and_updates_fields(self):
        """Issue 1: Cloning template cover keeps VML/images/styles and coexists with body tables."""
        tpl_path = self.root / "school_cover_template.docx"
        img_path = self.root / "badge.png"
        Image.new("RGB", (24, 24), "navy").save(img_path)

        tpl = Document()
        badge_p = tpl.add_paragraph()
        badge_p.add_run().add_picture(str(img_path))
        vml_p = tpl.add_paragraph()
        vml_p._p.append(
            parse_xml(
                f'<w:r {nsdecls("w")} xmlns:v="urn:schemas-microsoft-com:vml">'
                "<w:pict><v:shape id='_x0000_s1' style='width:120pt;height:40pt'>"
                "<v:textbox><w:txbxContent>"
                "<w:p><w:r><w:t>学校代码：13597</w:t></w:r></w:p>"
                "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>学    号：</w:t></w:r></w:p>"
                "</w:txbxContent></v:textbox>"
                "</v:shape></w:pict></w:r>"
            )
        )
        for line in (
            "信息检索与论文写作期末作业",
            "题    目：",
            "学生姓名：",
            "二级学院：默认学院",
            "专    业：",
            "指导教师：",
            "二 〇 二六 年 五 月",
        ):
            p = tpl.add_paragraph()
            r = p.add_run(line)
            r.bold = True
            r.font.size = Pt(14)
            r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")
        from docx.enum.section import WD_SECTION_START

        tpl.add_section(WD_SECTION_START.NEW_PAGE)
        tpl.add_paragraph("模板后续占位内容")
        tpl.save(tpl_path)
        tpl_hash = sha256(tpl_path)

        draft = clone_template_cover(
            tpl_path,
            fields={
                "学号": "2305010628",
                "题目": "基于JavaEE的仓库管理系统的设计与实现",
                "学生姓名": "王潇杨",
                "二级学院": "大数据与计算机科学学院",
                "专业": "计算机科学与技术",
                "指导教师": "张老师",
            },
        )
        draft.add_heading("第1章 绪论", level=1)
        draft.add_paragraph("正文内容。")
        tbl = draft.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "字段"
        tbl.cell(0, 1).text = "类型"
        tbl.cell(1, 0).text = "id"
        tbl.cell(1, 1).text = "INT"
        draft.save(self.source)

        self.assertEqual(sha256(tpl_path), tpl_hash)

        # Ensure table_text operations do NOT crash on the cover VML textbox
        request = {
            "operations": [
                {
                    "action": "font.size.set",
                    "target": {
                        "type": "paragraph",
                        "role": "table_text",
                        "part": "word/document.xml",
                        "all": True,
                    },
                    "params": {"value": 10.5},
                }
            ]
        }
        plan = build_plan(self.source, self.output, request)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")

        with zipfile.ZipFile(self.output) as pkg:
            doc_xml = pkg.read("word/document.xml").decode("utf-8")
            media_files = [n for n in pkg.namelist() if n.startswith("word/media/")]
        self.assertTrue(media_files)
        self.assertIn("2305010628", doc_xml)
        self.assertIn("基于JavaEE的仓库管理系统的设计与实现", doc_xml)
        self.assertIn("王潇杨", doc_xml)
        self.assertIn("大数据与计算机科学学院", doc_xml)
        self.assertNotIn("默认学院", doc_xml)
        self.assertNotIn("模板后续占位内容", doc_xml)

    def test_numbering_audit_accepts_gbt7714_iso_access_date_and_english_abstract_title(self):
        """GB/T 7714 [EB/OL] access dates like [2026-05-10] are dates, not typed citations."""
        from word_format.numbering_audit import audit_numbering

        doc = Document()
        doc.add_paragraph("基于JavaEE的仓库管理系统的设计与实现")
        doc.add_paragraph("摘  要")
        doc.add_paragraph("中文摘要正文。")
        doc.add_paragraph("关键词：仓库管理；JavaEE")
        doc.add_paragraph("Design and Implementation of Warehouse Management System")
        doc.add_paragraph("Abstract")
        doc.add_paragraph("English abstract body.")
        doc.add_paragraph("Key words: warehouse; JavaEE")
        doc.add_paragraph("参考文献")
        p_ref = doc.add_paragraph(
            "Apache Software Foundation. Apache POI[EB/OL]. (2024-01-15)[2026-05-10]. https://poi.apache.org/."
        )
        # Bind real Word numbering on p_ref
        p_ref._p.get_or_add_pPr().append(
            parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
        )

        idx = DocumentIndex(doc)
        en_title_rec = next(
            r for r in idx.records.values() if r.get("kind") == "paragraph" and "Design and Implementation" in r.get("text", "")
        )
        self.assertEqual(en_title_rec["semantic_role"], "abstract_title")
        self.assertEqual(en_title_rec["language"], "en")

        audit = audit_numbering(doc)
        self.assertEqual(audit["status"], "passed")
        self.assertEqual(audit["issues"], [])


if __name__ == "__main__":
    unittest.main()
