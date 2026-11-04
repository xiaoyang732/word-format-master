"""Effective-format regressions for legacy body requests (synthetic files only)."""

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

from word_format import apply_plan, build_plan, verify_plan
from word_format.inspect import DocumentIndex
from word_format.properties import read_font, read_paragraph, text_runs


FONT_PROPERTIES = ("font_east_asia", "font_latin", "font_size_pt", "bold", "italic", "color")
PARAGRAPH_PROPERTIES = (
    "alignment", "space_before_pt", "space_after_pt", "line_spacing",
    "first_line_indent", "left_indent", "right_indent", "hanging_indent",
    "keep_with_next", "keep_together", "page_break_before", "widow_control",
    "outline_level", "numbering_suppressed",
)


def effective_format(index, oid):
    paragraph = index.elements[oid]
    return {
        "paragraph": {name: read_paragraph(index.document, paragraph, name)
                      for name in PARAGRAPH_PROPERTIES},
        "runs": [{name: read_font(index.document, paragraph, run, name)
                  for name in FONT_PROPERTIES} for run in text_runs(paragraph)],
    }


class LegacyBodyScopeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "source.docx"
        self.output = self.folder / "output.docx"

    def test_body_spec_preserves_effective_format_of_every_nonbody_paragraph(self):
        # Changing the shared Normal style must fail this test even when all
        # unselected paragraph XML nodes remain byte-for-byte unchanged.
        doc = Document()
        normal = doc.styles["Normal"]
        normal.font.name = "Courier New"
        normal.font.size = Pt(11)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        normal.paragraph_format.line_spacing = 1.15
        doc.add_paragraph("测试文档", style="Title")
        doc.add_paragraph("学生姓名：测试作者")
        doc.add_paragraph("独创性声明")
        doc.add_paragraph("声明内容。")
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要内容。")
        doc.add_paragraph("关键词：排版；边界")
        doc.add_heading("第一章 绪论", 1)
        doc.add_heading("1.1 研究背景", 2)
        doc.add_heading("1.1.1 研究范围", 3)
        doc.add_paragraph("章节正文。")
        doc.add_table(rows=1, cols=1).cell(0, 0).text = "表格内容"
        doc.add_paragraph("参考文献", style="Heading 1")
        doc.add_paragraph("[1] 测试文献。")
        doc.add_paragraph("致谢", style="Heading 1")
        doc.add_paragraph("感谢所有测试参与者。")
        doc.add_paragraph("附录", style="Heading 1")
        doc.add_paragraph("附录内容。")
        doc.sections[0].header.paragraphs[0].text = "页眉内容"
        doc.sections[0].footer.paragraphs[0].text = "页脚内容"
        doc.save(self.source)
        source_bytes = self.source.read_bytes()
        before = DocumentIndex(Document(self.source))
        selected = {oid for oid, record in before.records.items()
                    if record.get("kind") == "paragraph" and record.get("role") == "body"
                    and record.get("module") in {None, "chapters"}}
        self.assertEqual([before.records[oid]["text"] for oid in selected], ["章节正文。"])
        unselected = {oid: effective_format(before, oid) for oid, record in before.records.items()
                      if record.get("kind") == "paragraph" and oid not in selected}

        plan = build_plan(self.source, self.output, {"spec": {"body": {
            "font_east_asia": "宋体", "font_latin": "Times New Roman", "font_size_pt": 12,
            "alignment": "justify", "first_line_indent_chars": 2,
            "line_spacing": {"kind": "multiple", "value": 1.5},
        }}})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        after = DocumentIndex(Document(self.output))
        for oid, original in unselected.items():
            with self.subTest(text=before.records[oid]["text"]):
                self.assertEqual(effective_format(after, oid), original)
        for oid in selected:
            actual = effective_format(after, oid)
            self.assertEqual(actual["runs"][0]["font_east_asia"], "宋体")
            self.assertEqual(actual["runs"][0]["font_latin"], "Times New Roman")
            self.assertEqual(actual["runs"][0]["font_size_pt"], 12)
            self.assertEqual(actual["paragraph"]["alignment"], "justify")
            self.assertEqual(actual["paragraph"]["first_line_indent"], {"value": 2, "unit": "chars"})
            self.assertEqual(actual["paragraph"]["line_spacing"],
                             {"kind": "multiple", "value": 1.5, "unit": "multiple"})
        self.assertEqual(self.source.read_bytes(), source_bytes)

    def test_partial_body_spec_preserves_other_effective_properties_in_body_only_input(self):
        doc = Document()
        normal = doc.styles["Normal"]
        normal.font.name = "Courier New"
        normal.font.size = Pt(11)
        normal.font.italic = True
        normal.font.color.rgb = RGBColor.from_string("123456")
        normal.paragraph_format.space_before = Pt(7)
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph = doc.add_paragraph("没有章节标记的普通正文。")
        paragraph.runs[0].bold = True
        doc.save(self.source)
        source_bytes = self.source.read_bytes()
        before = DocumentIndex(Document(self.source))
        original = effective_format(before, "p0")

        plan = build_plan(self.source, self.output, {"spec": {"body": {"font_size_pt": 12}}})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        after = DocumentIndex(Document(self.output))
        actual = effective_format(after, "p0")
        self.assertEqual(actual["runs"][0]["font_size_pt"], 12)
        original["runs"][0]["font_size_pt"] = 12
        self.assertEqual(actual, original)
        self.assertEqual(after.document.styles["Normal"].font.size, Pt(11))
        self.assertEqual(self.source.read_bytes(), source_bytes)


if __name__ == "__main__":
    unittest.main()
