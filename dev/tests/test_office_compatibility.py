"""Regressions for the Office audit's confirmed failures and publication contracts."""
# Project imports require the repository's scripts path below.
# ruff: noqa: E402

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from word_format import apply_plan, build_plan, verify_plan
from word_format.contracts import FormatError
from word_format.fields import refresh_document_fields
from word_format.inspect import DocumentIndex
from word_format.numbering import effective_numbering
from word_format.office_validation import expected_labels, field_errors, validate_office
from word_format.ooxml import sequence_warnings
from word_format.properties import element


class OfficeCompatibilityTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)
        self.source = self.folder / "source.docx"
        self.output = self.folder / "output.docx"

    def execute(self, document, operations):
        document.save(self.source)
        plan = build_plan(self.source, self.output, {"operations": operations})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        return Document(self.output)

    def test_heading_and_normal_references_keep_independent_sequences(self):
        for style in ("Normal", "Bibliography"):
            with self.subTest(style=style):
                self.output.unlink(missing_ok=True)
                doc = Document()
                doc.styles.add_style("Bibliography", 1)
                doc.add_heading("绪论", 1)
                doc.add_heading("背景", 2)
                doc.add_paragraph("正文")
                doc.add_paragraph("步骤", "List Number")
                doc.add_heading("方法", 1)
                doc.add_paragraph("参考文献")
                doc.add_paragraph("Author. One[M]. 2020.", style)
                doc.add_paragraph("Author. Two[M]. 2021.", style)
                operations = [
                    {
                        "action": "heading.numbering.configure",
                        "target": {"type": "document"},
                        "params": {"level1": "chinese", "level2": "arabic", "level3": "keep"},
                    },
                    {
                        "action": "list.numbering.configure",
                        "target": {"type": "document"},
                        "params": {"style": "decimal-period", "start": 1},
                    },
                    {
                        "action": "reference.numbering.configure",
                        "target": {"type": "document"},
                        "params": {"numbering_mode": "word-numbering"},
                    },
                ]
                # The public registry names are part of the behavior under test.
                from word_format.registry import REGISTRY

                operations[-1]["action"] = next(
                    name for name, cap in REGISTRY.items() if cap.property == "reference_numbering"
                )
                result = self.execute(doc, operations)
                root = result.part.numbering_part.element
                self.assertEqual(sequence_warnings(root), [])
                paragraphs = result.paragraphs
                ids = [effective_numbering(result, paragraphs[i]._p)[0] for i in (0, 3, 6)]
                self.assertEqual(len(set(ids)), 3)
                labels = expected_labels(result)
                self.assertEqual(
                    [labels[paragraphs[i]._p] for i in (0, 1, 4, 6, 7)],
                    ["第一章", "1.1", "第二章", "[1]", "[2]"],
                )

    def test_body_scope_inserts_real_boundary_and_preserves_section_geometry(self):
        doc = Document()
        doc.add_paragraph("封面", "Title")
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要文字")
        doc.add_page_break()
        doc.add_heading("业务需求", 1)
        doc.add_paragraph("正文")
        width = doc.sections[0].page_width
        result = self.execute(
            doc,
            [
                {
                    "action": "page_number.configure",
                    "target": {"type": "document"},
                    "params": {
                        "enabled": True,
                        "scope": "body",
                        "start": 1,
                        "show_on_body_first_page": True,
                    },
                }
            ],
        )
        self.assertEqual(len(result.sections), 2)
        self.assertTrue(all(section.page_width == width for section in result.sections))
        self.assertEqual(
            result.sections[-1]._sectPr.find(qn("w:pgNumType")).get(qn("w:start")), "1"
        )
        self.assertFalse(result.paragraphs[3]._p.findall(".//" + qn("w:br")))

    def test_scoped_page_numbers_follow_toc_and_survive_rebuild_with_shared_story(self):
        from word_format.fields import field_codes
        from word_format.headers import effective_root

        doc = Document()
        doc.sections[0].different_first_page_header_footer = True
        doc.sections[0].footer.paragraphs[0].text = "保留页脚"
        doc.sections[0].first_page_footer.paragraphs[0].text = "保留首页页脚"
        doc.add_paragraph("题目", "Title")
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要内容")
        doc.add_paragraph("关键词：文档")
        doc.add_heading("绪论", 1)
        doc.add_paragraph("正文")
        actions = [
            {
                "action": "page_number.configure", "target": {"type": "document"},
                "params": {"enabled": True, "scope": "body", "show_on_body_first_page": True},
            },
            {
                "action": "toc.configure", "target": {"type": "document"},
                "params": {"enabled": True, "page_break_after": True},
            },
        ]
        for _ in range(2):
            doc = self.execute(doc, actions)
            front, body = doc.sections
            for attr in ("footer", "first_page_footer"):
                self.assertEqual(field_codes(effective_root(getattr(front, attr))), [])
                self.assertEqual(field_codes(effective_root(getattr(body, attr))), ["PAGE"])
            self.assertIn("保留页脚", front.footer.paragraphs[0].text)
            self.assertIn("保留首页页脚", front.first_page_footer.paragraphs[0].text)
            paragraphs = [p.text for p in doc.paragraphs]
            self.assertLess(paragraphs.index("关键词：文档"), paragraphs.index("目录"))
            self.assertLess(paragraphs.index("目录"), paragraphs.index("绪论"))
            self.output.unlink()

    def test_sequence_rules_accept_repeated_run_property_choices(self):
        run = OxmlElement("w:r")
        pr = element(run, "rPr")
        for name in ("sz", "b", "rFonts", "b"):
            pr.append(OxmlElement("w:" + name))
        self.assertEqual(sequence_warnings(run), [])

    def test_ordered_writers_do_not_sort_unknown_extensions(self):
        pr = OxmlElement("w:pPr")
        element(pr, "sectPr")
        extension = OxmlElement("w:customXml")
        pr.append(extension)
        element(pr, "jc")
        element(pr, "pStyle")
        self.assertEqual(sequence_warnings(pr), [])
        self.assertIs(pr[-1], extension)

    def test_failed_word_refresh_does_not_publish_cleanup_or_modify_input(self):
        doc = Document()
        doc.add_paragraph("保留输入")
        doc.save(self.source)
        original = self.source.read_bytes()
        self.output.write_bytes(b"prior output")
        with patch("office_automation.automate_word", side_effect=RuntimeError("COM failure")):
            with self.assertRaisesRegex(FormatError, "not published"):
                refresh_document_fields(self.source, self.output, overwrite=True)
        self.assertEqual(self.source.read_bytes(), original)
        self.assertEqual(self.output.read_bytes(), b"prior output")
        self.assertEqual(list(self.folder.glob(".*.docx")), [])

    def test_only_field_results_trigger_error_scan(self):
        evidence = {
            "fields": [
                {
                    "code": "REF missing \\n",
                    "result": "错误！未找到引用源。",
                    "target_exists": False,
                },
                {"code": "DOCPROPERTY Title", "result": "错误分析研究"},
            ]
        }
        self.assertEqual(len(field_errors(evidence)), 1)
        evidence["fields"][0] = {
            "code": "REF WFMRef_1 \\n",
            "result": "[9]",
            "target_exists": True,
            "target_label": "[1]",
        }
        self.assertEqual(len(field_errors(evidence)), 1)

    def test_builtin_styleref_uses_language_independent_selector(self):
        doc = Document()
        doc.add_heading("绪论", 1)
        result = self.execute(
            doc,
            [
                {
                    "action": "header.content.set",
                    "target": {"type": "section", "all": True},
                    "params": {"mode": "chapter_title", "style_name": "Heading 1"},
                }
            ],
        )
        instructions = [
            n.get(qn("w:instr")) for n in result.sections[0].header._element.iter(qn("w:fldSimple"))
        ]
        self.assertEqual(instructions, ["STYLEREF 1"])

    def test_custom_style_with_heading_like_name_is_not_builtin(self):
        from word_format.headers import style_reference

        doc = Document()
        style = doc.styles.add_style('Heading Custom "A"', 1)
        self.assertEqual(style_reference(style), '"Heading Custom \\"A\\""')

    def test_office_check_rejects_wrong_computed_native_label(self):
        from word_format.numbering import apply_heading_scheme

        doc = Document()
        doc.add_heading("绪论", 1)
        apply_heading_scheme(
            DocumentIndex(doc),
            {},
            {"levels": [{"format": "decimal", "text": "%1", "start": 1, "suffix": "space"}]},
        )
        evidence = {
            "paragraphs": [{"text": "绪论\r", "list_string": "[1]", "page": 1}],
            "fields": [],
            "pagination_stable": True,
            "application": {"name": "Microsoft Word"},
        }
        result = validate_office(doc, evidence)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["errors"][0]["actual_label"], "[1]")

    def test_existing_theme_fonts_do_not_look_like_direct_format_mismatches(self):
        doc = Document()
        doc.add_paragraph("Existing theme text")
        evidence = {
            "paragraphs": [{
                "text": "Existing theme text\r",
                "font_east_asia": "+中文正文",
                "font_latin": "+西文正文",
                "font_size_pt": 11,
                "alignment": 0,
            }],
            "fields": [],
            "pagination_stable": True,
            "application": {"name": "Microsoft Word"},
        }
        result = validate_office(doc, evidence)
        self.assertEqual(result["status"], "passed")
        self.assertEqual(len(result["theme_font_aliases"]), 2)


if __name__ == "__main__":
    unittest.main()
