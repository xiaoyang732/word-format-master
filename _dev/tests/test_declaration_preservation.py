"""Declaration cloning boundaries and explicit preservation, without cached files."""

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))

from word_format import apply_plan, verify_plan
from word_format.cover import clone_template_cover
from word_format.profiles import profile_plan, requirements_profile
from word_format.properties import element


def declaration_template(section_break):
    doc = Document()
    doc.add_paragraph("毕业设计封面")
    if section_break:
        doc.add_section(WD_SECTION_START.NEW_PAGE)
    else:
        doc.add_page_break()
    doc.add_paragraph("诚信声明")
    blank = doc.add_paragraph()
    spacing = element(blank._p.get_or_add_pPr(), "spacing")
    spacing.set(qn("w:line"), "360")
    spacing.set(qn("w:lineRule"), "auto")
    element(blank._p.get_or_add_pPr(), "ind").set(qn("w:firstLine"), "420")
    doc.add_paragraph("本人签名：        2026 年月日")
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_paragraph("必须保留的后续普通正文")
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    doc.add_heading("摘要", 1)
    return doc


class DeclarationPreservationTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_clone_and_replace_stop_at_semantic_declaration_end(self):
        for section_break in (False, True):
            with self.subTest(section_break=section_break):
                template = declaration_template(section_break)
                path = self.root / "template.docx"
                template.save(path)
                result = clone_template_cover(path, fields={"signature": "王潇杨", "signature_date": "2026年6月6日"})
                texts = [p.text for p in result.paragraphs]
                self.assertIn("诚信声明", texts)
                self.assertFalse(any("后续普通正文" in t for t in texts))
                self.assertFalse(any("摘要" in t for t in texts))
                self.assertTrue(any("2026年6月6日" in t for t in texts))
                replaced = clone_template_cover(path, declaration_template(section_break), replace_existing_cover=True)
                self.assertEqual(sum("后续普通正文" in p.text for p in replaced.paragraphs), 1)
                self.assertEqual(sum(p.text == "诚信声明" for p in replaced.paragraphs), 1)

    def test_materialization_preserves_existing_blank_spacing_and_indent_xml(self):
        template = declaration_template(False)
        source_blank = next(p for p in template.paragraphs if p._p.find("./w:pPr/w:ind", p._p.nsmap) is not None)
        path = self.root / "template.docx"
        template.save(path)
        result = clone_template_cover(path)
        blank = next(p for p in result.paragraphs if p._p.find("./w:pPr/w:ind", p._p.nsmap) is not None)
        for tag in ("spacing", "ind"):
            self.assertEqual(dict(blank._p.pPr.find(qn("w:" + tag)).attrib),
                             dict(source_blank._p.pPr.find(qn("w:" + tag)).attrib))

    def test_preserve_excludes_blank_and_nonblank_paragraphs_from_broad_rules(self):
        path, output = self.root / "target.docx", self.root / "output.docx"
        doc = Document()
        doc.add_paragraph("诚信声明")
        blank = doc.add_paragraph()
        element(blank._p.get_or_add_pPr(), "spacing").set(qn("w:line"), "360")
        doc.add_paragraph("正文需要格式化")
        doc.save(path)
        original = Document(path)
        requirements = self.root / "requirements.txt"
        requirements.write_text("正文小四。", encoding="utf-8")
        profile = requirements_profile(requirements)
        for rule in profile["rules"]:
            rule["target"] = {"type": "paragraph", "all": True, "part": "word/document.xml"}
        plan = profile_plan(path, output, profile, mapping={"preserve": {"p0": "克隆声明原样保留", "p1": "声明空行原样保留"}})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        result = Document(output)
        for i in (0, 1):
            self.assertEqual(result.paragraphs[i]._p.xml, original.paragraphs[i]._p.xml)
        self.assertEqual(result.paragraphs[2].runs[0].font.size.pt, 12)
        self.assertFalse(any(oid in {"p0", "p1"} for m in report["profile_context"]["mapping"] for oid in m["target_ids"]))


if __name__ == "__main__":
    unittest.main()
