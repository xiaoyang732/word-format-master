"""Self-contained regressions for effective indents, declarations and font ranges."""

import sys
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

from word_format.cover import _replace_paragraph_text_pattern, _set_labeled_paragraph_value
from word_format.inspect import DocumentIndex
from word_format.properties import element, read_paragraph


def numbered_document(*, style_numbering=False, override=False):
    doc = Document()
    style = doc.styles.add_style("Reference fixture", WD_STYLE_TYPE.PARAGRAPH)
    element(style.element.get_or_add_pPr(), "ind").set(qn("w:firstLineChars"), "200")
    paragraph = doc.add_paragraph("[1] Reference", style)
    pr = paragraph._p.get_or_add_pPr()
    element(pr, "ind").set(qn("w:firstLineChars"), "0")
    num_pr = element(style.element.get_or_add_pPr() if style_numbering else pr, "numPr")
    element(num_pr, "numId").set(qn("w:val"), "42")
    element(num_pr, "ilvl").set(qn("w:val"), "0")
    numbering = doc.part.numbering_part.element
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), "42")
    level = element(abstract, "lvl")
    level.set(qn("w:ilvl"), "0")
    indent = element(element(level, "pPr"), "ind")
    indent.set(qn("w:left"), "920")
    indent.set(qn("w:hanging"), "440")
    numbering.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), "42")
    element(num, "abstractNumId").set(qn("w:val"), "42")
    if override:
        item = element(num, "lvlOverride")
        item.set(qn("w:ilvl"), "0")
        level = element(item, "lvl")
        level.set(qn("w:ilvl"), "0")
        indent = element(element(level, "pPr"), "ind")
        indent.set(qn("w:left"), "1440")
        indent.set(qn("w:hanging"), "720")
    numbering.append(num)
    return doc, paragraph


class TestDeclarationRefHF(unittest.TestCase):
    def test_signature_slot_preserves_inline_date(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("本人签名：").font.size = Pt(14)
        p.add_run("        2026 年月日")
        self.assertTrue(_set_labeled_paragraph_value(p._p, "王潇杨"))
        self.assertIn("本人签名：王潇杨", p.text)
        self.assertIn("2026 年月日", p.text)
        self.assertTrue(_replace_paragraph_text_pattern(
            p._p, r"(?:20\d\d|\d{2,4})\s*年\s*月\s*日|年\s*月\s*日", "2026年6月6日"
        ))
        self.assertIn("本人签名：王潇杨", p.text)
        self.assertIn("2026年6月6日", p.text)

    def test_numbering_hanging_indent_inheritance(self):
        doc, p = numbered_document()
        self.assertAlmostEqual(read_paragraph(doc, p._p, "left_indent")["value"], 16.2278, places=3)
        self.assertAlmostEqual(read_paragraph(doc, p._p, "hanging_indent")["value"], 7.7611, places=3)
        self.assertAlmostEqual(read_paragraph(doc, p._p, "first_line_indent")["value"], -7.7611, places=3)

    def test_plain_paragraph_zero_cancels_style_first_line(self):
        doc = Document()
        element(doc.styles["Normal"].element.get_or_add_pPr(), "ind").set(qn("w:firstLineChars"), "200")
        p = doc.add_paragraph("Ordinary text")
        element(p._p.get_or_add_pPr(), "ind").set(qn("w:firstLineChars"), "0")
        self.assertEqual(read_paragraph(doc, p._p, "first_line_indent"), {"value": 0.0, "unit": "chars"})
        self.assertEqual(read_paragraph(doc, p._p, "hanging_indent")["value"], 0)

    def test_zero_left_and_right_indents_override_style(self):
        doc = Document()
        inherited = element(doc.styles["Normal"].element.get_or_add_pPr(), "ind")
        inherited.set(qn("w:leftChars"), "200")
        inherited.set(qn("w:rightChars"), "200")
        p = doc.add_paragraph("Ordinary text")
        direct = element(p._p.get_or_add_pPr(), "ind")
        direct.set(qn("w:leftChars"), "0")
        direct.set(qn("w:rightChars"), "0")
        for name in ("left_indent", "right_indent"):
            self.assertEqual(read_paragraph(doc, p._p, name), {"value": 0.0, "unit": "chars"})

    def test_style_numbering_inherits_hanging(self):
        doc, p = numbered_document(style_numbering=True)
        self.assertAlmostEqual(read_paragraph(doc, p._p, "hanging_indent")["value"], 7.7611, places=3)

    def test_numbering_level_override_wins(self):
        doc, p = numbered_document(override=True)
        self.assertEqual(read_paragraph(doc, p._p, "left_indent"), {"value": 25.4, "unit": "mm"})
        self.assertEqual(read_paragraph(doc, p._p, "hanging_indent"), {"value": 12.7, "unit": "mm"})

    def test_direct_zero_num_id_disables_style_numbering(self):
        doc, p = numbered_document(style_numbering=True)
        element(element(p._p.get_or_add_pPr(), "numPr"), "numId").set(qn("w:val"), "0")
        self.assertEqual(read_paragraph(doc, p._p, "left_indent")["value"], 0)
        self.assertEqual(read_paragraph(doc, p._p, "first_line_indent")["value"], 0)

    def test_partial_num_pr_inherits_style_num_id(self):
        doc, p = numbered_document(style_numbering=True)
        element(element(p._p.get_or_add_pPr(), "numPr"), "ilvl").set(qn("w:val"), "0")
        self.assertAlmostEqual(read_paragraph(doc, p._p, "hanging_indent")["value"], 7.7611, places=3)

    def test_whitespace_only_range_uses_its_actual_run_font(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("   ").font.size = Pt(18)
        p.add_run("text").font.size = Pt(12)
        p.add_run("  ").font.size = Pt(16)
        index = DocumentIndex(doc)
        self.assertEqual(index.font_values("p0", "font_size_pt"), [12.0])
        self.assertEqual(index.font_values("p0", "font_size_pt", 0, 3), [18.0])
        self.assertEqual(index.font_values("p0", "font_size_pt", 7, 9), [16.0])


if __name__ == "__main__":
    unittest.main()
