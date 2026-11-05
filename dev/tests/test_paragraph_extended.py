"""Lossless paragraph settings across inspection, profiles and registered writes."""

import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from analyze_docx import analyze_docx, parse_paragraph_properties
from word_format import apply_plan, build_plan, inspect_document, verify_plan
from word_format.contracts import FormatError, sha256
from word_format.profiles import profile_plan, template_profile
from word_format.properties import read_paragraph
from word_format.registry import REGISTRY
from word_format.spec import legacy_paths


def add_setting(pr, tag, value=None):
    node = OxmlElement("w:" + tag)
    if value is not None:
        node.set(qn("w:val"), value)
    pr.append(node)
    return node


def add_tabs(pr, *stops):
    tabs = OxmlElement("w:tabs")
    for position, alignment, leader in stops:
        node = add_setting(tabs, "tab", alignment)
        node.set(qn("w:pos"), str(position))
        if leader is not None:
            node.set(qn("w:leader"), leader)
    pr.append(tabs)


def tab(position, alignment="right", leader="dot"):
    return {"alignment": alignment, "position": position, "unit": "twips", "leader": leader}


class ExtendedParagraphTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "source.docx"
        self.output = self.root / "output.docx"

    def document(self):
        doc = Document()
        doc.add_heading("第1章 格式检查", 1)
        doc.add_paragraph("第一段正文")
        doc.add_paragraph("第二段正文")
        return doc

    def operation(self, name, value, target=None):
        return {"action": "paragraph." + name + ".set", "params": {"value": value},
                "target": target or {"type": "paragraph", "id": "p1"}}

    def test_inspect_and_independent_analysis_keep_boolean_and_tab_semantics(self):
        doc = self.document()
        pr = doc.paragraphs[1]._p.get_or_add_pPr()
        for tag_name, value in (("contextualSpacing", None), ("snapToGrid", "0"), ("mirrorIndents", "1")):
            add_setting(pr, tag_name, value)
        add_tabs(pr, (1441, "right", "dot"), (2883, "decimal", "heavy"))
        doc.save(self.source)
        expected = {"contextual_spacing": True, "snap_to_grid": False, "mirror_indents": True,
                    "tabs": [tab(1441), tab(2883, "decimal", "heavy")]}
        record = next(r for r in inspect_document(self.source)["objects"] if r["id"] == "p1")
        for name, value in expected.items():
            self.assertEqual(record["format"].get(name), value, name)
        analysis = parse_paragraph_properties(ElementTree.fromstring(pr.xml))
        for name, value in expected.items():
            self.assertEqual(analysis.get(name), value, name)

    def test_effective_tabs_merge_inheritance_and_clear_positions(self):
        doc = self.document()
        base = doc.styles.add_style("Tab base", WD_STYLE_TYPE.PARAGRAPH)
        add_setting(base.element.get_or_add_pPr(), "contextualSpacing", "1")
        add_tabs(base.element.get_or_add_pPr(), (720, "left", None), (1441, "right", "dot"))
        child = doc.styles.add_style("Tab child", WD_STYLE_TYPE.PARAGRAPH)
        child.base_style = base
        add_setting(child.element.get_or_add_pPr(), "contextualSpacing", "0")
        add_tabs(child.element.get_or_add_pPr(), (720, "clear", None), (2161, "center", "hyphen"))
        doc.paragraphs[1].style = child
        add_tabs(doc.paragraphs[1]._p.get_or_add_pPr(), (1441, "clear", None), (2883, "decimal", "heavy"))
        self.assertEqual(read_paragraph(doc, doc.paragraphs[1]._p, "contextual_spacing"), False)
        self.assertEqual(read_paragraph(doc, doc.paragraphs[1]._p, "snap_to_grid"), True)
        self.assertEqual(read_paragraph(doc, doc.paragraphs[1]._p, "mirror_indents"), False)
        self.assertEqual(read_paragraph(doc, doc.paragraphs[1]._p, "tabs"),
                         [tab(2161, "center", "hyphen"), tab(2883, "decimal", "heavy")])

    def test_registered_writes_replace_inherited_tabs_and_preserve_other_xml(self):
        doc = self.document()
        add_tabs(doc.styles["Normal"].element.get_or_add_pPr(), (720, "left", None), (1441, "right", "dot"))
        add_setting(doc.paragraphs[1]._p.get_or_add_pPr(), "textDirection", "lrTb")
        doc.save(self.source)
        source_hash = sha256(self.source)
        before_unselected = doc.paragraphs[2]._p.xml
        expected_tabs = [tab(2883, "decimal", "heavy")]
        operations = [self.operation("contextual_spacing", True), self.operation("snap_to_grid", False),
                      self.operation("mirror_indents", True), self.operation("tabs", expected_tabs)]
        plan = build_plan(self.source, self.output, {"operations": operations})
        report = apply_plan(plan)
        self.assertEqual(verify_plan(plan, report)["status"], "passed")
        out = Document(self.output)
        pr = out.paragraphs[1]._p.pPr
        for tag_name, value in (("contextualSpacing", "1"), ("snapToGrid", "0"), ("mirrorIndents", "1")):
            self.assertEqual(pr.find(qn("w:" + tag_name)).get(qn("w:val")), value)
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "tabs"), expected_tabs)
        raw_tabs = {(n.get(qn("w:pos")), n.get(qn("w:val")), n.get(qn("w:leader")))
                    for n in pr.find(qn("w:tabs"))}
        self.assertEqual(raw_tabs, {("720", "clear", None), ("1441", "clear", None), ("2883", "decimal", "heavy")})
        self.assertEqual(pr.find(qn("w:textDirection")).get(qn("w:val")), "lrTb")
        self.assertEqual(out.paragraphs[2]._p.xml, before_unselected)
        self.assertEqual(sha256(self.source), source_hash)

    def test_style_target_can_clear_all_inherited_tabs(self):
        doc = self.document()
        add_tabs(doc.styles["Normal"].element.get_or_add_pPr(), (1441, "right", "dot"))
        doc.save(self.source)
        target = {"type": "style", "name": "Heading 1"}
        plan = build_plan(self.source, self.output, {"operations": [self.operation("tabs", [], target)]})
        self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
        out = Document(self.output)
        self.assertEqual(read_paragraph(out, out.paragraphs[0]._p, "tabs"), [])
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "tabs"), [tab(1441)])

    def test_replacing_tabs_retains_clears_that_already_hide_inherited_stops(self):
        doc = self.document()
        add_tabs(doc.styles["Normal"].element.get_or_add_pPr(), (720, "left", None), (1441, "right", "dot"))
        add_tabs(doc.paragraphs[1]._p.get_or_add_pPr(), (720, "clear", None))
        doc.save(self.source)
        plan = build_plan(self.source, self.output, {"operations": [self.operation("tabs", [tab(2883)])]})
        self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
        out = Document(self.output)
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "tabs"), [tab(2883)])

    def test_template_profile_migrates_exact_values_and_reports_unknown_nodes(self):
        sample = self.root / "sample.docx"
        doc = self.document()
        for paragraph in doc.paragraphs[1:]:
            pr = paragraph._p.get_or_add_pPr()
            for tag_name, value in (("contextualSpacing", "1"), ("snapToGrid", "0"), ("mirrorIndents", "1")):
                add_setting(pr, tag_name, value)
            add_tabs(pr, (1441, "right", "dot"))
            add_setting(pr, "textDirection", "lrTb")
        doc.save(sample)
        self.document().save(self.source)
        profile = template_profile(sample)
        for name in ("contextual_spacing", "snap_to_grid", "mirror_indents", "tabs"):
            rules = [r for r in profile["rules"] if r["action"] == "paragraph." + name + ".set"
                     and r["target"].get("role") == "body"]
            self.assertEqual(len(rules), 1, name)
        self.assertTrue(any("textDirection" in i.get("properties", []) for i in profile["issues"]))
        self.assertFalse(any(any(n in i.get("properties", []) for n in ("contextualSpacing", "snapToGrid", "mirrorIndents", "tabs"))
                             for i in profile["issues"]))
        decisions = {i["id"]: {"decision": "ignore", "reason": "Preserve unsupported or unspecified sample settings"}
                     for i in profile["issues"]}
        plan = profile_plan(self.source, self.output, profile, resolutions=decisions)
        self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
        out = Document(self.output)
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "tabs"), [tab(1441)])
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "snap_to_grid"), False)

    def test_malformed_parameters_fail_without_publishing_even_after_a_valid_operation(self):
        self.document().save(self.source)
        invalid = [self.operation("contextual_spacing", "true"), self.operation("snap_to_grid", 0),
                   self.operation("mirror_indents", None), self.operation("tabs", False),
                   self.operation("tabs", [tab(1441.5)]), self.operation("tabs", [tab(True)]),
                   self.operation("tabs", [tab(1441, "unknown")]),
                   self.operation("tabs", [tab(1441, leader="unknown")]),
                   self.operation("tabs", [{**tab(1441), "unit": "mm"}]),
                   self.operation("tabs", [tab(1441), tab(1441, "left")]),
                   self.operation("tabs", [{**tab(1441), "extra": "discard me"}])]
        for operation in invalid:
            with self.subTest(operation=operation), self.assertRaises(FormatError):
                build_plan(self.source, self.output, {"operations": [self.operation("contextual_spacing", True), operation]})
            self.assertFalse(self.output.exists())

    def test_broken_tab_writer_is_rejected_before_output_is_published(self):
        self.document().save(self.source)
        cap = REGISTRY["paragraph.tabs.set"]

        def broken(index, ref, params):
            cap.handler(index, ref, params)
            index.elements[ref["id"]].pPr.find(qn("w:tabs"))[0].set(qn("w:leader"), "none")

        with patch.dict(REGISTRY, {cap.action: replace(cap, handler=broken)}):
            plan = build_plan(self.source, self.output, {"operations": [self.operation("tabs", [tab(1441)])]})
            with self.assertRaises(FormatError):
                apply_plan(plan)
        self.assertFalse(self.output.exists())

    def test_spec_accepts_new_paragraph_tokens(self):
        self.document().save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {"body": {
            "contextual_spacing": True, "snap_to_grid": False, "mirror_indents": True,
            "tabs": [tab(1441)]}}})
        self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
        out = Document(self.output)
        self.assertEqual(read_paragraph(out, out.paragraphs[1]._p, "tabs"), [tab(1441)])

    def test_header_spec_uses_registered_extended_tokens(self):
        self.document().save(self.source)
        plan = build_plan(self.source, self.output, {"spec": {"headers_footers": {
            "header": {"enabled": True, "text": "Header", "tabs": [tab(1441)],
                       "contextual_spacing": True, "snap_to_grid": False, "mirror_indents": True}}}})
        self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
        for name in ("tabs", "contextual_spacing", "snap_to_grid", "mirror_indents"):
            self.assertIn("header." + name + ".set", [op["action"] for op in plan["operations"]])
        out = Document(self.output)
        paragraph = out.sections[0].header.paragraphs[0]
        self.assertEqual(read_paragraph(out, paragraph._p, "tabs"), [tab(1441)])
        self.assertEqual(read_paragraph(out, paragraph._p, "snap_to_grid"), False)

    def test_profile_reports_malformed_and_unknown_expanded_setting_evidence(self):
        doc = self.document()
        for paragraph in doc.paragraphs[1:]:
            pr = paragraph._p.get_or_add_pPr()
            add_setting(pr, "contextualSpacing", "sometimes")
            add_tabs(pr, (1441, "right", "dot"), (720, "clear", "invalid"))
            pr.find(qn("w:tabs"))[0].set(qn("w:customLeader"), "unsupported")
        doc.save(self.source)
        profile = template_profile(self.source)
        self.assertTrue(any(i.get("property") == "contextual_spacing" for i in profile["issues"]))
        self.assertTrue(any("tabs/tab@customLeader" in i.get("properties", []) for i in profile["issues"]))
        self.assertTrue(any("tabs" in i.get("properties", []) and "Malformed" in i["reason"] for i in profile["issues"]))
        self.assertEqual(profile_plan(self.source, self.output, profile)["status"], "needs_clarification")
        self.assertFalse(self.output.exists())

    def test_malformed_tab_position_is_preserved_as_unresolved_evidence(self):
        doc = self.document()
        pr = doc.paragraphs[1]._p.get_or_add_pPr()
        add_tabs(pr, ("--1", "right", "dot"))
        parsed = parse_paragraph_properties(ElementTree.fromstring(pr.xml))
        self.assertEqual(parsed["tabs"][0]["position"], "--1")
        doc.save(self.source)
        profile = template_profile(self.source)
        self.assertTrue(any("tabs" in i.get("properties", []) for i in profile["issues"]))
        self.assertEqual(profile_plan(self.source, self.output, profile)["status"], "needs_clarification")
        self.assertFalse(self.output.exists())

    def test_inferred_style_tokens_include_extended_effective_paragraph_properties(self):
        doc = self.document()
        defaults = doc.styles.element.find("./" + qn("w:docDefaults") + "/" + qn("w:pPrDefault") + "/" + qn("w:pPr"))
        add_tabs(defaults, (720, "left", None))
        normal_pr = doc.styles["Normal"].element.get_or_add_pPr()
        add_tabs(normal_pr, (720, "clear", None), (1441, "right", "dot"))
        for tag_name, value in (("contextualSpacing", "1"), ("snapToGrid", "0"), ("mirrorIndents", "1")):
            add_setting(normal_pr, tag_name, value)
        body = doc.styles["Body Text"]
        body.base_style = doc.styles["Normal"]
        add_tabs(body.element.get_or_add_pPr(), (2883, "decimal", "heavy"))
        doc.paragraphs[1].style = body
        doc.save(self.source)
        tokens = analyze_docx(self.source)["inferred_spec"]["body"]
        self.assertEqual(tokens.get("contextual_spacing"), True)
        self.assertEqual(tokens.get("snap_to_grid"), False)
        self.assertEqual(tokens.get("mirror_indents"), True)
        self.assertEqual(tokens.get("tabs"), [tab(1441), tab(2883, "decimal", "heavy")])

    def test_header_visual_alignment_counts_visible_tabs_and_not_tab_definitions(self):
        doc = self.document()
        header = doc.sections[0].header.paragraphs[0]
        header.alignment = 0
        add_tabs(header._p.get_or_add_pPr(), (4680, "center", None))
        header.add_run().add_tab()
        header.add_run("Centered template header")
        doc.save(self.source)
        config = analyze_docx(self.source)["inferred_spec"]["headers_footers"]["header"]
        self.assertEqual(config["alignment"], "center")

    def test_capability_paths_never_advertise_unregistered_header_actions(self):
        paths = legacy_paths()
        for path, action in paths.items():
            if path.startswith("headers_footers.") and action.startswith("header."):
                self.assertIn(action, REGISTRY, path)
        self.assertEqual(paths["body.tabs"], "paragraph.tabs.set")
        self.assertEqual(paths["headers_footers.header.tabs"], "header.tabs.set")

    def test_managed_toc_title_replaces_inherited_tab_set(self):
        for value in ([], [tab(2883)]):
            with self.subTest(value=value):
                doc = self.document()
                add_tabs(doc.styles["Normal"].element.get_or_add_pPr(), (720, "left", None), (1441, "right", "dot"))
                doc.save(self.source)
                output = self.root / ("empty.docx" if not value else "replaced.docx")
                request = {"spec": {"table_of_contents": {"enabled": True, "title_style": {
                    "tabs": value, "contextual_spacing": True, "snap_to_grid": False, "mirror_indents": True}}}}
                plan = build_plan(self.source, output, request)
                self.assertEqual(verify_plan(plan, apply_plan(plan))["status"], "passed")
                out = Document(output)
                title = next(p for p in out.paragraphs if p.style.name == "WFM TOC Heading")
                self.assertEqual(read_paragraph(out, title._p, "tabs"), value)
                self.assertEqual(read_paragraph(out, title._p, "snap_to_grid"), False)


if __name__ == "__main__":
    unittest.main()
