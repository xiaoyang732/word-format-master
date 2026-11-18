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

from word_format import (
    apply_plan,
    build_plan,
    normalize_universal_chinese_base,
    verify_plan,
)
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
            n.text.strip()
            for n in result.sections[0].header._element.iter(qn("w:instrText"))
        ]
        self.assertEqual(instructions, ["STYLEREF 1 \\* CHARFORMAT"])

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

    def test_numbered_references_and_body_lists_receive_template_typography(self):
        from word_format.properties import read_font, read_paragraph
        from word_format.templates import load_template, template_plan

        doc = Document()
        doc.add_heading("绪论", 1)
        doc.add_paragraph("正文说明。")
        doc.add_paragraph("记录输入", "List Number")
        doc.add_paragraph("检查输出", "List Number")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] Author A. Work One[M]. Beijing, 2024.")
        doc.add_paragraph("[2] Author B. Work Two[J]. Shanghai, 2025.")
        doc.save(self.source)

        # Simulate pre-numbered references produced by citation.insert before template formatting.
        cited = self.folder / "cited.docx"
        c_plan = build_plan(
            self.source,
            cited,
            {
                "operations": [
                    {
                        "action": "citation.insert",
                        "target": {"type": "document"},
                        "params": {
                            "placements": [
                                {
                                    "text_contains": "正文说明",
                                    "reference_ids": ["ref-1"],
                                    "citation_text": "[1]",
                                    "display": "superscript",
                                    "insert": "before-terminal-punctuation",
                                }
                            ]
                        },
                    }
                ]
            },
        )
        c_rep = apply_plan(c_plan)
        self.assertEqual(verify_plan(c_plan, c_rep)["status"], "passed")

        template, authority = load_template("thesis-standard")
        plan = template_plan(cited, self.output, template, template_source=authority)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        result = Document(self.output)
        bib_style = result.styles["Bibliography"]
        self.assertTrue(bib_style.builtin)
        self.assertIsNone(bib_style.element.get(qn("w:customStyle")))

        # Numbering levels for lists and references use space suffix (no tab jump).
        numbering_root = result.part.numbering_part.element
        self.assertEqual(sequence_warnings(numbering_root), [])
        for p in result.paragraphs:
            if "Author A" in p.text or "记录输入" in p.text:
                nid, ilvl, abstract, _ = effective_numbering(result, p._p)
                self.assertGreater(int(nid), 0)
                lvl = next(
                    n for n in abstract.findall(qn("w:lvl")) if n.get(qn("w:ilvl")) == str(ilvl)
                )
                suff = lvl.find(qn("w:suff"))
                self.assertIsNotNone(suff)
                self.assertEqual(suff.get(qn("w:val")), "space")

        list_p = next(p for p in result.paragraphs if "记录输入" in p.text)
        self.assertEqual(read_font(result, list_p._p, list_p.runs[0]._r, "font_east_asia"), "宋体")
        self.assertEqual(
            read_font(result, list_p._p, list_p.runs[0]._r, "font_latin"), "Times New Roman"
        )
        self.assertEqual(read_font(result, list_p._p, list_p.runs[0]._r, "font_size_pt"), 12.0)
        self.assertLess(read_paragraph(result, list_p._p, "first_line_indent")["value"], 0)

        ref_p = next(p for p in result.paragraphs if "Author A" in p.text)
        self.assertEqual(ref_p.style.name, "Bibliography")
        self.assertEqual(read_font(result, ref_p._p, ref_p.runs[0]._r, "font_east_asia"), "宋体")
        self.assertEqual(
            read_font(result, ref_p._p, ref_p.runs[0]._r, "font_latin"), "Times New Roman"
        )
        self.assertEqual(read_font(result, ref_p._p, ref_p.runs[0]._r, "font_size_pt"), 10.5)
        self.assertEqual(read_paragraph(result, ref_p._p, "alignment"), "justify")
        self.assertEqual(
            read_paragraph(result, ref_p._p, "first_line_indent"),
            {"value": -2.0, "unit": "chars"},
        )
        self.assertEqual(
            read_paragraph(result, ref_p._p, "line_spacing")["value"],
            1.0,
        )

    def test_paragraph_font_writes_format_field_control_runs_for_mergeformat(self):
        from word_format.properties import read_font

        doc = Document()
        doc.add_heading("绪论", 1)
        doc.add_paragraph("正文说明。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] Author A. Book[M]. 2025.")
        cited_doc = self.execute(
            doc,
            [
                {
                    "action": "citation.insert",
                    "target": {"type": "document"},
                    "params": {
                        "placements": [
                            {
                                "text_contains": "正文说明",
                                "reference_ids": ["ref-1"],
                                "citation_text": "[1]",
                                "display": "superscript",
                                "insert": "before-terminal-punctuation",
                            }
                        ]
                    },
                }
            ],
        )
        self.output.unlink()
        formatted = self.execute(
            cited_doc,
            [
                {
                    "action": "font.east_asia.set",
                    "target": {"type": "paragraph", "id": "p1"},
                    "params": {"value": "宋体"},
                },
                {
                    "action": "font.latin.set",
                    "target": {"type": "paragraph", "id": "p1"},
                    "params": {"value": "Times New Roman"},
                },
                {
                    "action": "font.size.set",
                    "target": {"type": "paragraph", "id": "p1"},
                    "params": {"value": 12},
                },
            ],
        )
        p1 = formatted.paragraphs[1]._p
        ctrl_runs = [
            r
            for r in p1.iter(qn("w:r"))
            if any(c.tag in {qn("w:fldChar"), qn("w:instrText")} for c in r)
        ]
        self.assertGreaterEqual(len(ctrl_runs), 3)
        for run in ctrl_runs:
            self.assertEqual(read_font(formatted, p1, run, "font_east_asia"), "宋体")
            self.assertEqual(read_font(formatted, p1, run, "font_latin"), "Times New Roman")
            self.assertEqual(read_font(formatted, p1, run, "font_size_pt"), 12.0)
            vert = run.find("./" + qn("w:rPr") + "/" + qn("w:vertAlign"))
            self.assertIsNotNone(vert)
            self.assertEqual(vert.get(qn("w:val")), "superscript")

    def test_toc_creates_indented_builtin_styles_and_ordered_outline_levels(self):
        from word_format.inspect import _style_paragraph
        from word_format.properties import read_font, read_paragraph
        from word_format.templates import load_template, template_plan

        doc = Document()
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要内容。")
        doc.add_paragraph("关键词：目录；样式")
        h1 = doc.add_paragraph("第1章 绪论")
        element(element(h1._p, "pPr"), "rPr")
        doc.add_paragraph("1.1 研究目标")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] Ref.")
        doc.save(self.source)

        template, authority = load_template("thesis-standard")
        plan = template_plan(self.source, self.output, template, template_source=authority)
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["execution_status"], "passed")

        out = Document(self.output)
        self.assertEqual(sequence_warnings(out.element.body), [])
        self.assertEqual(sequence_warnings(out.styles.element), [])
        for level, expected_left in ((1, 0.0), (2, 2.0), (3, 4.0)):
            toc_st = out.styles[f"toc {level}"]
            self.assertTrue(toc_st.builtin)
            probe = _style_paragraph(f"toc {level}", out)
            self.assertEqual(
                read_paragraph(out, probe, "first_line_indent"),
                {"value": 0.0, "unit": "chars"},
            )
            self.assertEqual(
                read_paragraph(out, probe, "left_indent"),
                {"value": expected_left, "unit": "chars"},
            )
            self.assertEqual(read_font(out, probe, None, "font_east_asia"), "宋体")
            self.assertFalse(read_font(out, probe, None, "bold"))
        wfm_ref_probe = _style_paragraph("WFM references Heading", out)
        self.assertEqual(read_font(out, wfm_ref_probe, None, "font_east_asia"), "黑体")
        self.assertTrue(read_font(out, wfm_ref_probe, None, "bold"))

    def test_three_line_table_and_empty_numbering_obey_ooxml_sequences(self):
        doc = Document()
        doc.add_paragraph("第一项", "List Number")
        doc.add_paragraph("第二项", "List Number")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "指标"
        tbl.cell(0, 1).text = "数值"
        tbl.cell(1, 0).text = "准确率"
        tbl.cell(1, 1).text = "99%"
        for row in tbl.rows:
            for cell in row.cells:
                element(cell._tc.get_or_add_tcPr(), "vAlign").set(qn("w:val"), "center")
        numbering = doc.part.numbering_part.element
        for child in list(numbering):
            numbering.remove(child)

        out = self.execute(
            doc,
            [
                {
                    "action": "list.numbering.configure",
                    "target": {"type": "document"},
                    "params": {"style": "decimal-parenthesis", "start": 1},
                },
                {
                    "action": "table.three_line.apply",
                    "target": {"type": "table", "ordinal": 1},
                    "params": {"value": True},
                },
                {
                    "action": "table.alignment.set",
                    "target": {"type": "table", "ordinal": 1},
                    "params": {"value": "center"},
                },
            ],
        )
        self.assertEqual(sequence_warnings(out.element.body), [])
        self.assertEqual(sequence_warnings(out.part.numbering_part.element), [])
        num_id = int(out.paragraphs[0]._p.xpath("./w:pPr/w:numPr/w:numId")[0].get(qn("w:val")))
        self.assertGreater(num_id, 0)

    def test_nameless_word_style_does_not_crash_document_index(self):
        from word_format.inspect import inspect_document

        doc = Document()
        nameless = OxmlElement("w:style")
        nameless.set(qn("w:type"), "paragraph")
        nameless.set(qn("w:customStyle"), "1")
        nameless.set(qn("w:styleId"), "af1")
        doc.styles.element.append(nameless)
        p = doc.add_paragraph("目录")
        element(element(p._p, "pPr"), "pStyle").set(qn("w:val"), "af1")
        doc.add_heading("绪论", 1)
        doc.save(self.source)

        inspected = inspect_document(self.source)
        self.assertEqual(inspected["objects"][1]["style_id"], "af1")
        self.assertEqual(inspected["objects"][1]["style_name"], "")

    def test_cover_cloning_matches_research_titles_strips_callouts_and_orders_ppr(self):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        from word_format.cover import clone_template_cover

        tpl_path = self.folder / "cover_tpl.docx"
        tpl = Document()
        callout_p = tpl.add_paragraph("中石油课题研究报告")
        callout_p._p.append(
            parse_xml(
                f'<w:r {nsdecls("w")} xmlns:v="urn:schemas-microsoft-com:vml">'
                "<w:pict><v:shape id='_x0000_s9'><v:textbox><w:txbxContent>"
                "<w:p><w:r><w:t>楷体，二号，居中</w:t></w:r></w:p>"
                "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>"
            )
        )
        title_p = tpl.add_paragraph("课题题目：   XXXXXX   ")
        element(element(title_p._p, "pPr"), "ind").set(qn("w:left"), "240")
        cm_start = OxmlElement("w:commentRangeStart")
        cm_start.set(qn("w:id"), "0")
        title_p._p.insert(0, cm_start)
        cm_run = OxmlElement("w:r")
        cm_ref = OxmlElement("w:commentReference")
        cm_ref.set(qn("w:id"), "0")
        cm_run.append(cm_ref)
        title_p._p.append(cm_run)
        tpl.save(tpl_path)

        cloned = clone_template_cover(tpl_path, fields={"title": "高精度排版引擎研究"})
        self.assertEqual(cloned._wfm_cover_fields["matched"], ["title"])
        self.assertEqual(sequence_warnings(cloned.element.body), [])
        self.assertEqual(list(cloned.element.body.iter(qn("w:commentReference"))), [])
        self.assertEqual(list(cloned.element.body.iter(qn("w:pict"))), [])

    def test_office_validation_catches_undefined_citation_fonts_and_toc_reference_drift(self):
        from word_format.templates import load_template, template_plan

        doc = Document()
        doc.add_paragraph("摘要")
        doc.add_paragraph("摘要内容。")
        doc.add_paragraph("关键词：校验；兼容")
        doc.add_heading("绪论", 1)
        doc.add_paragraph("正文内容。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] Author. Reference[M]. 2025.")
        doc.save(self.source)
        template, authority = load_template("thesis-standard")
        plan = template_plan(self.source, self.output, template, template_source=authority)
        apply_plan(plan)
        formatted = Document(self.output)

        bad_evidence = {
            "paragraphs": [
                {"text": "摘要\r", "font_east_asia": "黑体", "font_size_pt": 16, "alignment": 1},
                {
                    "text": "摘要内容。\r",
                    "font_east_asia": "宋体",
                    "font_latin": "Times New Roman",
                    "font_size_pt": 12,
                    "alignment": 3,
                    "first_line_indent_chars": 2,
                    "line_spacing_rule": 1,
                    "line_spacing_pt": 18,
                },
                {
                    "text": "关键词：校验；兼容\r",
                    "font_east_asia": "宋体",
                    "font_latin": "Times New Roman",
                    "font_size_pt": 12,
                    "alignment": 0,
                    "first_line_indent_chars": 0,
                    "line_spacing_rule": 1,
                    "line_spacing_pt": 18,
                },
                {"text": "目　　录\r", "font_east_asia": "+中文正文", "font_size_pt": 14, "alignment": 1},
                {
                    "text": "第一章 绪论\t1\r",
                    "style": "TOC 1",
                    "font_east_asia": "",
                    "font_latin": "Times New Roman",
                    "font_size_pt": 12,
                },
                {
                    "text": "绪论\r",
                    "list_string": "第一章",
                    "page": 1,
                    "physical_page": 2,
                    "font_east_asia": "黑体",
                    "font_latin": "Times New Roman",
                    "font_size_pt": 16,
                    "alignment": 1,
                    "first_line_indent_chars": 0,
                    "line_spacing_rule": 1,
                    "line_spacing_pt": 18,
                },
                {
                    "text": "正文内容。\r",
                    "font_east_asia": "",
                    "font_latin": "",
                    "font_size_pt": 9999999,
                    "alignment": 3,
                    "first_line_indent_chars": 2,
                    "line_spacing_rule": 1,
                    "line_spacing_pt": 18,
                },
                {"text": "参考文献\r", "font_east_asia": "黑体", "font_size_pt": 16, "alignment": 1},
                {
                    "text": "Author. Reference[M]. 2025.\r",
                    "list_string": "[1]",
                    "font_east_asia": "+中文正文",
                    "font_latin": "+西文正文",
                    "font_size_pt": 11,
                    "alignment": 0,
                    "first_line_indent_chars": 0,
                    "line_spacing_rule": 5,
                    "line_spacing_pt": 13.8,
                },
            ],
            "sections": [
                {"index": 1, "adjusted_page": 1, "physical_page": 1},
                {"index": 2, "adjusted_page": 1, "physical_page": 2},
            ],
            "fields": [],
            "pagination_stable": True,
            "application": {"name": "Microsoft Word"},
        }
        check = validate_office(formatted, bad_evidence)
        self.assertEqual(check["status"], "failed")
        props = {(e.get("paragraph"), e.get("property")) for e in check["errors"]}
        self.assertIn(("toc:1:第一章 绪论\t1", "font_east_asia"), props)
        body_id = next(
            oid
            for oid, rec in DocumentIndex(formatted).records.items()
            if rec.get("text") == "正文内容。"
        )
        ref_id = next(
            oid
            for oid, rec in DocumentIndex(formatted).records.items()
            if "Author. Reference" in rec.get("text", "")
        )
        self.assertIn((body_id, "font_size_pt"), props)
        self.assertIn((ref_id, "font_size_pt"), props)
        self.assertIn((ref_id, "first_line_indent"), props)
        self.assertIn((ref_id, "line_spacing_pt"), props)

    def test_numbering_order_repair_handles_document_without_numbering_part(self):
        doc = Document()
        for rel_id, rel in list(doc.part.rels.items()):
            if "numbering" in rel.reltype:
                doc.part.rels.pop(rel_id)
        doc.add_paragraph("无编号部件文档")
        result = self.execute(
            doc,
            [{"action": "numbering.order.repair", "target": {"type": "document"}, "params": {}}],
        )
        self.assertEqual(result.paragraphs[0].text, "无编号部件文档")

    def test_caption_position_set_allows_id_shift_during_verify_plan(self):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        doc = Document()
        doc.add_paragraph("图 1 架构示意图")
        fig_p = doc.add_paragraph()
        fig_p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:drawing/></w:r>'))
        doc.add_paragraph("后续正文")
        result = self.execute(
            doc,
            [
                {
                    "action": "caption.position.set",
                    "target": {"type": "document"},
                    "params": {"role": "figure", "position": "below"},
                },
                {
                    "action": "font.size.set",
                    "target": {"type": "paragraph", "contains": "图 1"},
                    "params": {"value": 10.5},
                },
            ],
        )
        self.assertIn("图 1", result.paragraphs[1].text)

    def test_wps_backend_isolation_for_detection_conversion_and_refresh(self):
        import runtime_detection
        from word_format.doc_loader import _convert_legacy_word_file

        with (
            patch.object(
                runtime_detection,
                "_find_word",
                return_value=Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"),
            ),
            patch.object(
                runtime_detection,
                "_word_application_com_server",
                return_value=Path(r"C:\Users\demo\AppData\Local\Kingsoft\WPS Office\wps.exe"),
            ),
            patch.object(
                runtime_detection,
                "_find_wps",
                return_value=(
                    "KWPS.Application",
                    Path(r"C:\Users\demo\AppData\Local\Kingsoft\WPS Office\wps.exe"),
                ),
            ),
            patch.object(runtime_detection, "_find_libreoffice", return_value=None),
        ):
            detected = runtime_detection.detect_renderers(self.folder)
            self.assertFalse(detected["word"]["available"])
            self.assertTrue(detected["wps"]["available"])
            self.assertEqual(detected["preferred"], "wps")

        staged_wps = self.folder / "legacy.wps"
        staged_wps.write_bytes(b"\xd0\xcf\x11\xe0LEGACY-WPS")
        converted_out = self.folder / "converted.docx"
        sample_doc = Document()
        sample_doc.add_paragraph("WPS converted")
        sample_docx_bytes = self.folder / "backing.docx"
        sample_doc.save(sample_docx_bytes)

        calls = []

        def fake_wps_convert(src, dst):
            calls.append(("wps", src.suffix.lower()))
            dst.write_bytes(sample_docx_bytes.read_bytes())

        with (
            patch(
                "word_format.doc_loader.detect_renderers",
                return_value={
                    "word": {"available": True},
                    "wps": {"available": True},
                    "libreoffice": {"available": False},
                },
            ),
            patch("word_format.doc_loader._wps_com_to_docx", side_effect=fake_wps_convert),
            patch(
                "word_format.doc_loader._word_com_to_docx",
                side_effect=AssertionError("word should not precede wps for .wps files"),
            ),
        ):
            _convert_legacy_word_file(staged_wps, converted_out)
        self.assertEqual(calls, [("wps", ".wps")])

        # Auto refresh falls back to WPS when Word is unavailable.
        sample_doc.save(self.source)
        with (
            patch(
                "runtime_detection.detect_renderers",
                return_value={"word": {"available": False}, "wps": {"available": True}},
            ),
            patch(
                "office_automation.automate_wps",
                return_value={
                    "application": {"name": "WPS Office", "version": "12.0"},
                    "fields": [],
                    "field_update_errors": [],
                    "pagination_stable": True,
                    "update_passes": 1,
                    "coverage": {"story_types": [1]},
                },
            ),
        ):
            refreshed = refresh_document_fields(self.source, self.output, overwrite=True)
        self.assertEqual(refreshed["method"], "wps_com")
        self.assertEqual(refreshed["refresh_status"], "wps_refreshed")

    def test_shipped_course_report_template_is_clean_of_personal_metadata_and_nameless_styles(self):
        import zipfile
        from xml.etree import ElementTree as ET

        tpl = ROOT / "templates/reference/general-course-report/通用课程报告模板.docx"
        with zipfile.ZipFile(tpl) as zf:
            core = ET.fromstring(zf.read("docProps/core.xml"))
            for elem in core.iter():
                if elem.tag.endswith(("lastModifiedBy", "creator")):
                    self.assertFalse((elem.text or "").strip(), f"Metadata leaked in {elem.tag}")
            styles = ET.fromstring(zf.read("word/styles.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for st in styles.findall("w:style", ns):
                self.assertIsNotNone(st.find("w:name", ns))
        doc = Document(tpl)
        self.assertEqual(sequence_warnings(doc.styles.element), [])
        self.assertEqual(sequence_warnings(doc.part.numbering_part.element), [])

    def test_universal_chinese_base_respects_custom_config_and_dynamic_section_width(self):
        from docx.shared import Mm
        from word_format.constants import UniversalChineseBaseConfig, section_content_width_twips
        from word_format.properties import read_font, read_paragraph

        doc = Document()
        sec = doc.sections[0]
        sec.page_width = Mm(297)
        sec.left_margin = Mm(25)
        sec.right_margin = Mm(25)
        expected_tab = section_content_width_twips(sec)

        cfg = UniversalChineseBaseConfig(
            east_asia_body_font="仿宋",
            latin_font="Calibri",
            body_font_size_pt=14.0,
            toc_indent_step_chars=1.5,
        )
        normalize_universal_chinese_base(doc, config=cfg)

        p = doc.add_paragraph("配置化基线测试")
        self.assertEqual(read_font(doc, p._p, None, "font_east_asia"), "仿宋")
        self.assertEqual(read_font(doc, p._p, None, "font_latin"), "Calibri")
        self.assertEqual(read_font(doc, p._p, None, "font_size_pt"), 14.0)

        toc2 = doc.styles["toc 2"]
        probe = doc.add_paragraph("目录项", style=toc2)
        self.assertEqual(
            read_paragraph(doc, probe._p, "left_indent"),
            {"value": 1.5, "unit": "chars"},
        )
        tabs = read_paragraph(doc, probe._p, "tabs")
        self.assertEqual(tabs[0]["position"], expected_tab)

    def test_validate_office_cursor_window_prevents_distant_duplicate_jump(self):
        doc = Document()
        doc.add_heading("第一章 绪论", 1)
        doc.add_heading("实验设置", 2)
        for idx in range(1, 12):
            doc.add_paragraph(f"第一章正文第{idx}段。")
        doc.add_heading("第二章 方法", 1)
        doc.add_heading("实验设置", 2)
        doc.add_paragraph("第二章正文。")
        normalize_universal_chinese_base(doc)

        # Simulate first "实验设置" text rewritten by a dynamic field while second "实验设置"
        # appears >8 paragraphs later; cursor must stay in Chapter 1 and validate Chapter 1 prose.
        paragraphs = []
        for idx, p in enumerate(doc.paragraphs):
            text = "实验设置（动态域已刷新）" if idx == 1 else p.text
            paragraphs.append(
                {
                    "text": text,
                    "style": p.style.name,
                    "font_east_asia": "宋体" if "正文" in p.text else "黑体",
                    # Deliberate mismatch on Chapter 1 body paragraph #5 to prove it was NOT skipped!
                    "font_latin": "Arial" if p.text == "第一章正文第5段。" else "Times New Roman",
                    "font_size_pt": 12.0
                    if "正文" in p.text
                    else (15.0 if "章" in p.text else 14.0),
                }
            )
        evidence = {
            "application": {"name": "Microsoft Word", "version": "16.0"},
            "paragraphs": paragraphs,
            "fields": [],
            "toc_entries": [],
            "pagination_stable": True,
            "update_passes": 1,
            "field_update_errors": [],
        }
        result = validate_office(doc, evidence)
        self.assertEqual(result["status"], "failed")
        self.assertTrue(
            any(
                err.get("property") == "font_latin" and err.get("actual") == "Arial"
                for err in result["errors"]
            ),
            result["errors"],
        )


if __name__ == "__main__":
    unittest.main()
