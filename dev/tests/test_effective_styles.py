"""Real DOCX regressions for default styles and inherited style toggles."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
from word_format.api import apply_plan, audit_document, build_plan, verify_plan
from word_format.contracts import sha256
from word_format.inspect import DocumentIndex
from word_format.profiles import template_profile
from word_format.properties import read_paragraph


class EffectiveStyleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.output = self.root / "source.docx", self.root / "output.docx"

    def default_document(self, remove_normal=False):
        document = Document()
        normal = document.styles["Normal"]
        normal.font.size = Pt(12)
        normal.element.attrib.pop(qn("w:default"), None)
        actual = document.styles.add_style("Actual Default", WD_STYLE_TYPE.PARAGRAPH)
        actual.element.set(qn("w:default"), "1")
        actual.font.size = Pt(18)
        actual.font.bold = True
        actual.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph("Default style body")
        if remove_normal:
            document.styles.element.remove(normal.element)
        document.save(self.source)
        return actual.style_id

    def test_real_default_style_is_used_by_inspect_profile_audit_and_apply(self):
        for remove_normal in (False, True):
            with self.subTest(remove_normal=remove_normal):
                style_id = self.default_document(remove_normal)
                original_hash = sha256(self.source)
                document = Document(self.source)
                self.assertEqual(document.paragraphs[0].style.style_id, style_id)
                index = DocumentIndex(document)
                self.assertEqual(index.records["p0"]["style_id"], style_id)
                self.assertEqual(index.font_values("p0", "font_size_pt"), [18])
                self.assertEqual(
                    read_paragraph(document, document.paragraphs[0]._p, "alignment"), "center"
                )
                profile = template_profile(self.source)
                sizes = [
                    r["params"]["value"] for r in profile["rules"] if r["action"] == "font.size.set"
                ]
                self.assertIn(18, sizes)
                self.assertNotIn(12, sizes)
                request = {
                    "operations": [
                        {
                            "action": "font.size.set",
                            "target": {"type": "paragraph", "id": "p0"},
                            "params": {"value": 12},
                        },
                        {
                            "action": "paragraph.alignment.set",
                            "target": {"type": "paragraph", "id": "p0"},
                            "params": {"value": "left"},
                        },
                    ]
                }
                self.assertEqual(audit_document(self.source, request)["summary"]["failed"], 2)
                plan = build_plan(self.source, self.output, request)
                report = apply_plan(plan, overwrite=True)
                self.assertTrue(all(check["changed"] for check in report["checks"]))
                self.assertEqual(verify_plan(plan, report)["status"], "passed")
                out = Document(self.output).paragraphs[0]
                self.assertEqual(out.runs[0].font.size.pt, 12)
                self.assertEqual(out.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.LEFT)
                self.assertEqual(out.text, "Default style body")
                self.assertEqual(sha256(self.source), original_hash)

    def test_style_toggles_set_effective_bold_and_italic(self):
        for name in ("bold", "italic"):
            for inherited in (False, True):
                for wanted in (False, True):
                    with self.subTest(property=name, inherited=inherited, wanted=wanted):
                        document = Document()
                        base = document.styles.add_style("Base", WD_STYLE_TYPE.PARAGRAPH)
                        setattr(base.font, name, inherited)
                        child = document.styles.add_style("Child", WD_STYLE_TYPE.PARAGRAPH)
                        child.base_style = base
                        document.add_paragraph("Styled content", style=child)
                        document.save(self.source)
                        request = {
                            "operations": [
                                {
                                    "action": f"font.{name}.set",
                                    "target": {"type": "style", "name": "Child"},
                                    "params": {"value": wanted},
                                }
                            ]
                        }
                        plan = build_plan(self.source, self.output, request)
                        report = apply_plan(plan, overwrite=True)
                        self.assertEqual(verify_plan(plan, report)["status"], "passed")
                        output = Document(self.output)
                        self.assertEqual(DocumentIndex(output).font_values("p0", name), [wanted])
                        tag = "b" if name == "bold" else "i"
                        raw = output.styles["Child"].element.find(
                            "./" + qn("w:rPr") + "/" + qn("w:" + tag)
                        )
                        self.assertEqual(raw.get(qn("w:val")), "1" if wanted != inherited else "0")
                        self.assertIsNone(output.paragraphs[0].runs[0]._r.rPr)


if __name__ == "__main__":
    unittest.main()
