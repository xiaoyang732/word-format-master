"""Scoped page numbers must survive independent verification and linked parts."""

import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

from word_format import apply_plan, build_plan, verify_plan
from word_format.contracts import FormatError
from word_format.fields import field_codes
from word_format.headers import effective_root
from word_format.properties import element
from word_format.registry import REGISTRY
from word_format.structures import append_field
from verify_output import verify_structure


class ScopedPageNumberTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.source = Path(temporary.name) / "source.docx"
        self.output = Path(temporary.name) / "output.docx"
        doc = Document()
        doc.add_paragraph("封面")
        doc.sections[0].footer.paragraphs[0].text = "保留页脚文字"
        append_field(doc.sections[0].footer.add_paragraph(), "PAGE")
        doc.settings.odd_and_even_pages_header_footer = True
        doc.sections[0].different_first_page_header_footer = True
        append_field(doc.sections[0].first_page_footer.paragraphs[0], "PAGE")
        append_field(doc.sections[0].even_page_footer.paragraphs[0], "PAGE")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_heading("摘要", 1)
        doc.add_paragraph("摘要第一页")
        doc.add_page_break()
        doc.add_paragraph("摘要第二页")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_heading("第1章 绪论", 1)
        doc.add_paragraph("正文")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_heading("第2章 实现", 1)
        element(doc.sections[-1]._sectPr, "pgNumType").set(qn("w:start"), "9")
        doc.save(self.source)

    def plan(self, **params):
        return build_plan(self.source, self.output, {"operations": [{
            "action": "page_number.configure", "target": {"type": "document"},
            "params": {"enabled": True, "start": 1, **params},
        }]})

    def execute(self, **params):
        plan = self.plan(**params)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        return Document(self.output)

    def test_body_scope_cleans_all_front_variants_and_restarts_body(self):
        doc = self.execute(scope="body", show_on_first_page=False, show_on_body_first_page=True)
        for i, section in enumerate(doc.sections):
            for attr in ("footer", "first_page_footer", "even_page_footer"):
                codes = field_codes(effective_root(getattr(section, attr)))
                self.assertEqual(sum(c == "PAGE" for c in codes), 0 if i < 2 else 1)
            if i >= 2:
                page_type = section._sectPr.find(qn("w:pgNumType"))
                self.assertEqual(page_type.get(qn("w:fmt")), "decimal")
                self.assertEqual(page_type.get(qn("w:start")), "1" if i == 2 else None)
        self.assertIn("保留页脚文字", "".join(n.text or "" for n in doc.sections[0].footer._element.iter(qn("w:t"))))

    def test_explicit_one_based_start_section(self):
        doc = self.execute(start_section=4, start=5)
        self.assertEqual(doc.sections[3]._sectPr.find(qn("w:pgNumType")).get(qn("w:start")), "5")
        for section in list(doc.sections)[:3]:
            self.assertNotIn("PAGE", field_codes(effective_root(section.footer)))

    def test_default_all_scope_keeps_front_page_numbers_but_resets_body(self):
        doc = self.execute(show_on_body_first_page=True)
        self.assertIn("PAGE", field_codes(effective_root(doc.sections[0].footer)))
        self.assertEqual(doc.sections[2]._sectPr.find(qn("w:pgNumType")).get(qn("w:start")), "1")

    def test_bad_scope_and_start_section_are_rejected(self):
        for params in ({"scope": "guess"}, {"start_section": True}, {"start_section": 0},
                       {"start_section": 5}, {"start_section": 1.5}):
            with self.subTest(params=params), self.assertRaises(FormatError):
                self.plan(**params)

    def test_missing_body_marker_does_not_silently_number_front_matter(self):
        doc = Document(self.source)
        for p in doc.paragraphs:
            if "第1章" in p.text or "第2章" in p.text:
                p.text = "业务分析"
        doc.save(self.source)
        with self.assertRaises(FormatError):
            self.execute(scope="body")

    def test_verifier_rejects_wrong_body_reset(self):
        plan = self.plan(scope="body")
        cap = REGISTRY["page_number.configure"]
        def bad(index, ref, params):
            cap.handler(index, ref, params)
            element(index.document.sections[2]._sectPr, "pgNumType").set(qn("w:start"), "7")
        with patch.dict(REGISTRY, {cap.action: replace(cap, handler=bad)}):
            with self.assertRaises(FormatError):
                apply_plan(plan)
        self.assertFalse(self.output.exists())

    def test_legacy_structure_verifier_obeys_explicit_scope_and_reset(self):
        self.execute(start_section=4, start=5)
        spec = {"page_numbers": {"enabled": True, "start_section": 4, "start": 5}}
        self.assertEqual(verify_structure(self.output, spec)["status"], "passed")
        doc = Document(self.output)
        element(doc.sections[3]._sectPr, "pgNumType").set(qn("w:start"), "7")
        doc.save(self.output)
        self.assertEqual(verify_structure(self.output, spec)["status"], "failed")


if __name__ == "__main__":
    unittest.main()
