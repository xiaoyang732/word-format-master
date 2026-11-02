"""Regression tests for Declaration cloning, Reference indent inheritance, and Header/Page-number profiling."""

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

import unittest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word_format.cover import clone_template_cover
from word_format.inspect import DocumentIndex
from word_format.profiles import template_profile
from word_format.properties import read_paragraph, read_font


class TestDeclarationRefHF(unittest.TestCase):
    def setUp(self):
        # Locate cached template docx if available
        import tempfile, glob, os
        cache_dir = os.path.join(tempfile.gettempdir(), "word-format-master-doc-cache")
        cached = glob.glob(os.path.join(cache_dir, "1e7908811a36*.docx"))
        self.sample_path = cached[0] if cached else None

    def test_signature_slot_preserves_inline_date(self):
        """_set_labeled_paragraph_value must clamp before inline date patterns and preserve date."""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("                      本人签名：")
        r2 = p.add_run("        2026 年月日")
        # Ensure r1/r2 have 14pt (sz=28) font
        r1.font.size = None
        r1._r.get_or_add_rPr().append(OxmlElement("w:sz"))
        r1._r.rPr.find(qn("w:sz")).set(qn("w:val"), "28")

        from word_format.cover import _set_labeled_paragraph_value, _replace_paragraph_text_pattern
        ok = _set_labeled_paragraph_value(p._p, "王潇杨")
        self.assertTrue(ok)
        self.assertIn("本人签名：王潇杨", p.text)
        self.assertIn("2026 年月日", p.text)

        # Update date
        ok_date = _replace_paragraph_text_pattern(
            p._p,
            r"(?:20\d\d|\d{2,4})\s*年\s*月\s*日|年\s*月\s*日",
            "2026年6月6日",
        )
        self.assertTrue(ok_date)
        self.assertIn("本人签名：王潇杨", p.text)
        self.assertIn("2026年6月6日", p.text)

    def test_numbering_hanging_indent_inheritance(self):
        """Paragraph with numPr referencing abstractNum must inherit hanging and left indent."""
        if not self.sample_path:
            self.skipTest("Sample docx not found in cache")
        doc = Document(self.sample_path)
        # Find reference paragraph (p105)
        p_ref = next(p for p in doc.paragraphs if "考试招生信息化平台" in p.text)
        left = read_paragraph(doc, p_ref._p, "left_indent")
        hanging = read_paragraph(doc, p_ref._p, "hanging_indent")
        first_line = read_paragraph(doc, p_ref._p, "first_line_indent")
        line_spacing = read_paragraph(doc, p_ref._p, "line_spacing")

        # 920 twips = 16.23 mm
        self.assertAlmostEqual(left["value"], 16.23, delta=0.1)
        # hanging = 440 twips = 7.76 mm (0.78 cm)
        self.assertAlmostEqual(hanging["value"], 7.76, delta=0.1)
        # first-line indent is negative hanging indent (-7.76 mm)
        self.assertAlmostEqual(first_line["value"], -7.76, delta=0.1)
        # line spacing is exact 22 pt
        self.assertEqual(line_spacing["kind"], "exact")
        self.assertEqual(line_spacing["value"], 22.0)

    def test_template_profile_extracts_headers_and_page_numbers(self):
        """template_profile must extract header content and footer page number rules."""
        if not self.sample_path:
            self.skipTest("Sample docx not found in cache")
        profile = template_profile(self.sample_path)
        actions = [r["action"] for r in profile["rules"]]
        self.assertIn("header.content.set", actions)
        self.assertIn("header.alignment.set", actions)
        self.assertIn("page_number.configure", actions)

        # Check reference_item rules
        ref_rules = [
            r for r in profile["rules"]
            if r["target"].get("semantic_role") == "reference_item"
        ]
        ref_actions = {r["action"] for r in ref_rules}
        self.assertIn("paragraph.hanging_indent.set", ref_actions)
        self.assertIn("paragraph.left_indent.set", ref_actions)
        self.assertIn("paragraph.line_spacing.set", ref_actions)

        # Check declaration signature font size (14.0 pt)
        decl_sig_rules = [
            r for r in profile["rules"]
            if r["target"].get("semantic_role") == "declaration_signature" and r["action"] == "font.size.set"
        ]
        self.assertTrue(len(decl_sig_rules) > 0)
        self.assertEqual(decl_sig_rules[0]["params"]["value"], 14.0)


if __name__ == "__main__":
    unittest.main()
