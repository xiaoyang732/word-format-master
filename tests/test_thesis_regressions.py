"""Regressions from the thesis generation trial (fixtures only, no user files)."""

from pathlib import Path
import sys
import tempfile
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from document_structure import heading_level
from word_format import apply_plan, build_plan, verify_plan
from word_format.properties import read_paragraph


class ThesisRegressionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "source.docx"
        self.output = self.folder / "output.docx"

    def test_chapter_descriptions_stay_body_through_numbering_and_toc(self):
        doc = Document()
        doc.add_paragraph("第一章 绪论", style="Heading 1")
        descriptions = [
            "第一章：绪论。介绍研究背景与贡献。",
            "第二章：相关理论与技术基础。阐述技术基础。",
            "第3章介绍实验方法与结果。",
            "1.5倍行距用于正文。",
            "1. 本文首先介绍研究问题。",
        ]
        for text in descriptions:
            doc.add_paragraph(text)
        doc.add_paragraph("第二章 技术基础")
        doc.add_paragraph("2.1 静态分析")
        doc.add_paragraph("本节内容。")
        doc.save(self.source)
        spec = {"body": {"font_size_pt": 12, "first_line_indent_chars": 2},
                "headings": [{"level": 1, "font_size_pt": 16}, {"level": 2, "font_size_pt": 14}],
                "lists": {"headings": {"level1": "chinese", "level2": "arabic", "level3": "arabic"}},
                "table_of_contents": {"enabled": True}}
        plan = build_plan(self.source, self.output, {"spec": spec})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        for text in descriptions:
            p = next(p for p in out.paragraphs if p.text == text)
            self.assertEqual(p.style.name, "Normal")
            self.assertIsNone(p._p.find("./" + qn("w:pPr") + "/" + qn("w:numPr")))
            self.assertEqual(read_paragraph(out, p._p, "first_line_indent"), {"value": 2, "unit": "chars"})
        self.assertEqual(sum(p.style.name == "Heading 1" for p in out.paragraphs), 2)

    def test_explicit_headings_override_text_heuristics(self):
        self.assertEqual(heading_level("第一章：研究问题", "Heading 1"), 1)
        self.assertEqual(heading_level("1.2 A heading", "Heading 3"), 3)
        self.assertIsNone(heading_level("第一章 绪论", "TOC 1"))
        self.assertIsNone(heading_level("1. 项目", "List Paragraph"))
        self.assertIsNone(heading_level("1.5倍行距", "Normal"))
        self.assertIsNone(heading_level("86.4%", "Normal"))
        self.assertIsNone(heading_level("普通正文", "Normal", 9))
        doc = Document()
        doc.add_paragraph("第一章：研究问题", style="Heading 1")
        doc.add_paragraph("正文")
        doc.add_paragraph("1.2 A heading", style="Heading 3")
        doc.add_paragraph("更深层次", style="Heading 4")
        doc.save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {
            "lists": {"headings": {"level1": "chinese", "level2": "arabic", "level3": "arabic"}}}})
        apply_plan(plan)
        out = Document(self.output)
        self.assertEqual(out.paragraphs[2].style.name, "Heading 3")
        self.assertEqual(out.paragraphs[3].style.name, "Heading 4")

    def test_batching_retains_global_rules_and_module_overrides(self):
        doc = Document()
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要正文")
        doc.add_paragraph("第一章 绪论", style="Heading 1")
        for i in range(24):
            p = doc.add_paragraph(f"正文段落 {i}")
            p.runs[0].font.size = Pt(20)
            p.paragraph_format.first_line_indent = Pt(0)
        doc.save(self.source)
        spec = {"body": {"font_size_pt": 12, "first_line_indent_chars": 2, "alignment": "justify"},
                "headings": [{"level": 1, "font_size_pt": 16}],
                "document_structure": {"ordered_modules": [
                    {"kind": "abstract", "style_roles": {"body": {"tokens": {"font_size_pt": 10, "space_after_pt": 5}}}},
                    {"kind": "chapters"}]}}
        plan = build_plan(self.source, self.output, {"spec": spec})
        self.assertLess(len(plan["operations"]), 20)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        # Explicit global body values win; the unrelated module spacing survives.
        self.assertEqual(out.paragraphs[1].runs[0].font.size.pt, 12)
        self.assertEqual(out.paragraphs[1].paragraph_format.space_after.pt, 5)
        self.assertEqual(out.paragraphs[2].runs[0].font.size.pt, 16)
        for p in out.paragraphs[3:]:
            self.assertEqual(p.runs[0].font.size.pt, 12)
            self.assertEqual(read_paragraph(out, p._p, "first_line_indent"), {"value": 2, "unit": "chars"})

    def test_blank_paragraphs_and_page_breaks_preserve_contents(self):
        doc = Document()
        doc.add_paragraph("封面", style="Title")
        doc.add_paragraph()
        doc.add_paragraph()
        doc.add_page_break()
        doc.add_paragraph("独创性声明")
        doc.add_paragraph("声明正文")
        doc.add_page_break()
        doc.add_paragraph("第一章 绪论", style="Heading 1")
        doc.add_paragraph()
        doc.add_paragraph("正文")
        doc.save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {"body": {"font_size_pt": 12, "first_line_indent_chars": 2}}})
        report = apply_plan(plan)
        out = Document(self.output)
        self.assertEqual([p.text for p in doc.paragraphs], [p.text for p in out.paragraphs])
        self.assertEqual(len(list(doc.element.iter(qn("w:br")))), len(list(out.element.iter(qn("w:br")))))
        self.assertEqual(report["verification"]["preservation"]["status"], "passed")

    def test_body_rules_do_not_reformat_cover_metadata(self):
        doc = Document()
        cover = doc.add_paragraph("合成论文封面标题")
        cover.runs[0].font.size = Pt(22)
        cover.paragraph_format.alignment = 1
        cover.paragraph_format.first_line_indent = Pt(0)
        doc.add_paragraph()
        doc.add_paragraph("独创性声明")
        doc.add_paragraph("声明正文")
        doc.add_paragraph("第一章 绪论", style="Heading 1")
        doc.add_paragraph("章节正文")
        doc.save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {"body": {
            "font_size_pt": 12, "alignment": "justify", "first_line_indent_chars": 2}}})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        self.assertEqual(out.paragraphs[0]._p.xml, doc.paragraphs[0]._p.xml)
        self.assertEqual(out.paragraphs[-1].runs[0].font.size.pt, 12)
        # A document with no front-matter marker must remain a body-only input.
        self.source = self.folder / "plain.docx"
        self.output = self.folder / "plain-output.docx"
        plain = Document()
        plain.add_paragraph("普通正文").runs[0].font.size = Pt(20)
        plain.save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {"body": {"font_size_pt": 12}}})
        apply_plan(plan)
        self.assertEqual(Document(self.output).paragraphs[0].runs[0].font.size.pt, 12)

    def test_indent_repairs_conflicting_xml_and_replaces_style_hanging(self):
        doc = Document()
        style = doc.styles.add_style("Hanging body", WD_STYLE_TYPE.PARAGRAPH)
        style.paragraph_format.first_line_indent = Pt(-24)
        for size in (10.5, 12, 16):
            p = doc.add_paragraph(f"正文字号 {size}", style=style)
            p.runs[0].font.size = Pt(size)
            ind = OxmlElement("w:ind")
            for name, value in {"firstLineChars": "200", "firstLine": "0", "hanging": "0", "hangingChars": "0"}.items():
                ind.set(qn("w:" + name), value)
            p._p.get_or_add_pPr().append(ind)
            self.assertEqual(read_paragraph(doc, p._p, "first_line_indent")["value"], 0)
        doc.save(self.source)
        plan = build_plan(self.source, self.output, {"operations": [{
            "action": "paragraph.first_line_indent.set", "target": {"type": "paragraph", "all": True},
            "params": {"value": 2, "unit": "chars"}}]})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        for p in out.paragraphs:
            ind = p._p.pPr.find(qn("w:ind"))
            self.assertEqual(dict(ind.attrib), {qn("w:firstLineChars"): "200"})
            self.assertEqual(read_paragraph(out, p._p, "first_line_indent"), {"value": 2, "unit": "chars"})
            self.assertEqual(read_paragraph(out, p._p, "hanging_indent")["value"], -2)
        self.assertEqual([p.runs[0].font.size.pt for p in out.paragraphs], [10.5, 12, 16])

    def test_physical_and_zero_indents_replace_opposite_inherited_units(self):
        doc = Document()
        style = doc.styles.add_style("Character indent", WD_STYLE_TYPE.PARAGRAPH)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), "200")
        style.element.get_or_add_pPr().append(ind)
        for text in ("physical first", "physical hanging", "clear first"):
            doc.add_paragraph(text, style=style)
        hanging = doc.styles.add_style("Inherited hanging", WD_STYLE_TYPE.PARAGRAPH)
        hanging.paragraph_format.first_line_indent = Pt(-24)
        doc.add_paragraph("clear hanging", style=hanging)
        doc.add_paragraph("clear first through hanging", style=style)
        doc.save(self.source)
        ops = []
        for i, (action, value) in enumerate((("first_line_indent", 7.4), ("hanging_indent", 7.4), ("first_line_indent", 0))):
            ops.append({"action": f"paragraph.{action}.set", "target": {"type": "paragraph", "id": f"p{i}"},
                        "params": {"value": value, "unit": "mm"}})
        ops.append({"action": "paragraph.first_line_indent.set", "target": {"type": "paragraph", "id": "p3"},
                    "params": {"value": 0, "unit": "chars"}})
        ops.append({"action": "paragraph.hanging_indent.set", "target": {"type": "paragraph", "id": "p4"},
                    "params": {"value": 0, "unit": "mm"}})
        plan = build_plan(self.source, self.output, {"operations": ops})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        self.assertAlmostEqual(read_paragraph(out, out.paragraphs[0]._p, "first_line_indent")["value"], 7.4, delta=25.4 / 1440)
        self.assertAlmostEqual(read_paragraph(out, out.paragraphs[1]._p, "first_line_indent")["value"], -7.4, delta=25.4 / 1440)
        self.assertAlmostEqual(read_paragraph(out, out.paragraphs[1]._p, "hanging_indent")["value"], 7.4, delta=25.4 / 1440)
        self.assertEqual(read_paragraph(out, out.paragraphs[2]._p, "first_line_indent")["value"], 0)
        self.assertEqual(read_paragraph(out, out.paragraphs[3]._p, "first_line_indent")["value"], 0)
        self.assertEqual(read_paragraph(out, out.paragraphs[4]._p, "first_line_indent")["value"], 0)
        self.assertTrue(report["checks"][-1]["changed"])


if __name__ == "__main__":
    unittest.main()
