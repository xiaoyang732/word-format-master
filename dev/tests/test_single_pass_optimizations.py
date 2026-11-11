"""Regression tests for single-pass formatting optimizations and acceptance remediation."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

import run_format_pipeline as pipeline
from word_format import apply_plan, build_plan, verify_plan
from word_format.cover import clone_template_cover
from word_format.fields import strip_last_rendered_page_breaks_in_doc
from word_format.numbering_audit import audit_numbering
from word_format.profiles import check_template_structure, profile_plan, template_profile
from word_format.remediation import (
    VISUAL_DEFECT_PLAYBOOK,
    build_visual_report_template,
    diagnose_numbering_remediation,
    diagnose_structure_remediation,
)
from word_format.template_structure import compare_template_structure
from word_format.visual import verify_visual_evidence


class SinglePassOptimizationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_toc_appendix_entry_cover_banner_and_typed_references_need_zero_manual_rule_mappings(
        self,
    ):
        sample_path = self.root / "sample.docx"
        draft_path = self.root / "draft.docx"
        output_path = self.root / "output.docx"

        def build_doc(path: Path, *, is_sample: bool) -> None:
            doc = Document()
            normal = doc.styles["Normal"]
            normal.font.name = "Times New Roman"
            normal.font.size = Pt(12)
            normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

            # Cover spacer (10.5pt), cover title (48pt), cover banner "文献综述" (48pt), metadata, date
            spacer = doc.add_paragraph(" ")
            spacer.runs[0].font.size = Pt(10.5)
            p_title = doc.add_paragraph("学术前沿专题")
            p_title.runs[0].font.size = Pt(48)
            p_title.runs[0].bold = True
            p_title.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")

            p_banner = doc.add_paragraph("文献综述")
            p_banner.runs[0].font.size = Pt(48)
            p_banner.runs[0].bold = True
            p_banner.runs[0]._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "黑体")

            p_subj = doc.add_paragraph(
                "题    目：" + ("样例题目" if is_sample else "大语言模型综述")
            )
            p_subj.runs[0].font.size = Pt(14)
            p_date = doc.add_paragraph("二〇二六年十一月")
            p_date.runs[0].font.size = Pt(18)
            doc.add_page_break()

            # Abstract & TOC (including a cached TOC entry starting with "附录1：...")
            p_abs_title = doc.add_paragraph("样例题目" if is_sample else "大语言模型综述")
            p_abs_title.runs[0].font.size = Pt(16)
            doc.add_paragraph("摘  要")
            doc.add_paragraph("这是摘要正文内容。")
            doc.add_paragraph("关键词：大模型；综述")

            doc.add_paragraph("目 录")
            toc_style = doc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
            toc_style.font.size = Pt(12)
            doc.add_paragraph("第1章 绪论\t1", style=toc_style)
            doc.add_paragraph("附录1：系统原理图39", style=toc_style)

            # Chapter & References
            doc.add_heading("第1章 绪论", level=1)
            doc.add_paragraph("这是第一章正文内容。")
            doc.add_paragraph("参考文献")
            if is_sample:
                doc.add_paragraph("[1] 张三. 样例文献[J]. 计算机学报, 2024.")
            else:
                ref_p = doc.add_paragraph("李四. 目标文献[J]. 软件学报, 2025.", style="List Number")
                ref_p.runs[0].font.size = Pt(10.5)
            doc.save(path)

        build_doc(sample_path, is_sample=True)
        build_doc(draft_path, is_sample=False)

        profile = template_profile(sample_path)
        # 1. TOC entry "附录1：系统原理图39" must NOT produce any appendix_heading rules
        self.assertFalse(
            any(r["target"].get("semantic_role") == "appendix_heading" for r in profile["rules"])
        )
        # 2. Cover banner "文献综述" rule must include text="文献综述" so spacer paragraphs stay 10.5pt
        banner_rules = [
            r
            for r in profile["rules"]
            if r["target"].get("module") == "cover"
            and r["target"].get("semantic_role") == "cover_body"
            and r["action"] == "font.size.set"
        ]
        self.assertTrue(banner_rules)
        self.assertEqual(banner_rules[0]["target"].get("text"), "文献综述")

        # 3. Override abstract_title font size via resolution "replace" WITHOUT needing mapping.rules preserve
        resolutions = {
            issue["id"]: {"decision": "ignore", "reason": "忽略测试样例未定义项"}
            for issue in profile["issues"]
        }
        if profile["issues"]:
            first_issue_id = profile["issues"][0]["id"]
            resolutions[first_issue_id] = {
                "decision": "replace",
                "reason": "按考核要求将中文摘要标题字号覆盖为18pt",
                "rules": [
                    {
                        "action": "font.size.set",
                        "target": {
                            "type": "paragraph",
                            "semantic_role": "abstract_title",
                            "language": "zh",
                            "all": True,
                        },
                        "params": {"value": 18.0, "unit": "pt"},
                    }
                ],
            }

        plan = profile_plan(
            draft_path,
            output_path,
            profile,
            resolutions=resolutions,
            match_structure=True,
        )
        self.assertNotIn("status", plan, f"Expected executable plan, got: {plan.get('issues')}")
        report = apply_plan(plan)
        verified = verify_plan(plan, report)
        self.assertEqual(verified["status"], "passed")

        out_doc = Document(output_path)
        # Spacer paragraph on cover was NOT inflated to 48pt
        self.assertEqual(out_doc.paragraphs[0].runs[0].font.size.pt, 10.5)
        # Cover banner "文献综述" kept 48pt
        self.assertEqual(out_doc.paragraphs[2].runs[0].font.size.pt, 48.0)

    def test_cover_hygiene_strips_red_format_hints_resets_date_indent_and_compacts_spacers(self):
        from PIL import Image

        tpl_path = self.root / "cover_tpl.docx"
        img_path = self.root / "logo.png"
        Image.new("RGB", (16, 16), "white").save(img_path)

        tpl = Document()
        # Paragraph 0: image-only paragraph (empty text, contains <w:drawing>)
        p_img = tpl.add_paragraph()
        p_img.add_run().add_picture(str(img_path))
        # Paragraph 1: cover title
        tpl.add_paragraph("学术前沿专题")
        # Paragraph 2: subject
        tpl.add_paragraph("题    目：旧题目")
        # Paragraphs 3, 4, 5, 6: 4 pure-empty spacer paragraphs above date
        for _ in range(4):
            tpl.add_paragraph("")
        # Paragraph 7: centered date with compensating firstLine indent and trailing red format hint
        p_date = tpl.add_paragraph()
        p_date._p.get_or_add_pPr().append(
            parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="2520" w:firstLineChars="700"/>')
        )
        p_date._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
        r_date = p_date.add_run("二 〇 二 六 年 十 一 月")
        r_date.bold = True
        r_date.font.size = Pt(18)
        r_hint = p_date.add_run("（小二号居中")
        r_hint.bold = False
        r_hint._r.get_or_add_rPr().append(parse_xml(f'<w:color {nsdecls("w")} w:val="FF0000"/>'))
        tpl.add_page_break()
        tpl.add_paragraph("摘  要")
        tpl.add_paragraph("摘要内容")
        tpl.save(tpl_path)

        cloned = clone_template_cover(
            tpl_path,
            fields={"题目": "新综述题目", "date": "二 〇 二 六 年 十 一 月"},
            compact_spacers=3,
        )
        # Image paragraph must be preserved!
        self.assertTrue(
            any(node.tag == qn("w:drawing") for node in cloned.paragraphs[0]._p.iter()),
            "Image paragraph must never be deleted by spacer compaction",
        )
        # 3 of the 4 empty spacer paragraphs were compacted
        date_p = next(p for p in cloned.paragraphs if "二 〇 二 六 年" in p.text)
        self.assertNotIn("小二号居中", date_p.text)
        # Red run and compensating firstLine indent must be gone
        self.assertFalse(
            any(
                c.get(qn("w:val")) == "FF0000"
                for c in date_p._p.iter(qn("w:color"))
            )
        )
        ind = date_p._p.find("./" + qn("w:pPr") + "/" + qn("w:ind"))
        self.assertTrue(ind is None or qn("w:firstLine") not in ind.attrib)

    def test_citation_insert_supports_text_contains_shifted_indices_and_closed_ref_fields(self):
        src_path = self.root / "cite_source.docx"
        out_path = self.root / "cite_output.docx"

        doc = Document()
        doc.add_heading("第1章 绪论", level=1)
        doc.add_paragraph("新增的第一章引导段。")
        # Paragraph with an embedded closed REF field (e.g. "如表 1.1 所示")
        p_table_ref = doc.add_paragraph()
        p_table_ref.add_run("如表 ")
        for tag, val in [
            ("fldChar", "begin"),
            ("instrText", " REF Ref_Table_1_1 \\h "),
            ("fldChar", "separate"),
            ("t", "1.1"),
            ("fldChar", "end"),
        ]:
            r = OxmlElement("w:r")
            el = OxmlElement(f"w:{tag}")
            if tag == "fldChar":
                el.set(qn("w:fldCharType"), val)
            else:
                el.text = val
            r.append(el)
            p_table_ref._p.append(r)
        p_table_ref.add_run(" 所示，注意力机制显著提升了序列建模能力。")

        p_second = doc.add_paragraph("混合专家架构降低了推理激活参数量。")
        doc.add_heading("参考文献", level=1)
        doc.add_paragraph("[1] Vaswani A. Attention is all you need[C]. 2017.")
        doc.add_paragraph("[2] Fedus W. Switch transformers[J]. 2022.")
        doc.save(src_path)

        from citation_workflow import paragraph_sha256

        # Placement 1 uses text_contains (on a paragraph containing a closed REF field!).
        # Placement 2 passes a stale paragraph_index (1 instead of 3) but valid paragraph_sha256.
        request = {
            "operations": [
                {
                    "action": "citation.insert",
                    "target": {"type": "document"},
                    "params": {
                        "placements": [
                            {
                                "text_contains": "所示，注意力机制显著提升了序列建模能力",
                                "reference_ids": ["ref-1"],
                                "citation_text": "[1]",
                                "insert": "before-terminal-punctuation",
                                "display": "superscript",
                            },
                            {
                                "paragraph_index": 1,
                                "paragraph_sha256": paragraph_sha256(p_second.text),
                                "reference_ids": ["ref-2"],
                                "citation_text": "[2]",
                                "insert": "before-terminal-punctuation",
                                "display": "superscript",
                            },
                        ]
                    },
                }
            ]
        }
        plan = build_plan(src_path, out_path, request)
        report = apply_plan(plan)
        verified = verify_plan(plan, report)
        self.assertEqual(verified["status"], "passed")

        out_doc = Document(out_path)
        self.assertIn("序列建模能力[1]。", out_doc.paragraphs[2].text)
        self.assertIn("激活参数量[2]。", out_doc.paragraphs[3].text)

    def test_last_rendered_page_break_does_not_fail_template_structure(self):
        sample_path = self.root / "sample_struct.docx"
        draft_path = self.root / "draft_struct.docx"

        sample = Document()
        p = sample.add_paragraph("封面标题")
        # Continuous section break on cover (separate_page=False in raw sample)
        p._p.get_or_add_pPr().append(
            parse_xml(
                f'<w:sectPr {nsdecls("w")}><w:type w:val="continuous"/></w:sectPr>'
            )
        )
        sample.add_paragraph("摘  要")
        sample.add_paragraph("摘要内容")
        sample.add_heading("第1章 绪论", level=1)
        sample.add_paragraph("正文")
        sample.save(sample_path)

        profile = template_profile(sample_path)

        draft = Document(sample_path)
        # Simulate Word COM field update injecting <w:lastRenderedPageBreak/> on cover
        draft.paragraphs[0].runs[0]._r.insert(0, OxmlElement("w:lastRenderedPageBreak"))
        draft.save(draft_path)

        # Structure check passes even with <w:lastRenderedPageBreak/>
        res = check_template_structure(draft_path, profile)
        self.assertEqual(res["status"], "passed")

        # And strip_last_rendered_page_breaks_in_doc removes the cached node cleanly
        reopened = Document(draft_path)
        removed = strip_last_rendered_page_breaks_in_doc(reopened)
        self.assertEqual(removed, 1)

    def test_acceptance_remediation_visual_template_and_resume_parameter_inheritance(self):
        source = self.root / "src.docx"
        output = self.root / "out.docx"
        reports = self.root / "reports"

        doc = Document()
        doc.styles["Normal"].font.name = "Times New Roman"
        doc.styles["Normal"].font.size = Pt(12)
        doc.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        doc.add_heading("绪论", level=1)
        doc.add_paragraph("测试正文。")
        doc.save(source)

        # Verify numbering remediation diagnostics
        bad_doc = Document()
        bad_doc.add_paragraph("表 1.1 测试表格题注")
        bad_doc.add_paragraph("正文包含手打引用[1]。")
        num_audit = audit_numbering(bad_doc)
        self.assertEqual(num_audit["status"], "failed")
        self.assertTrue(num_audit.get("remediation"))
        rem_kinds = {item["kind"] for item in num_audit["remediation"]}
        self.assertIn("caption", rem_kinds)
        self.assertIn("reference_or_citation", rem_kinds)

        # Verify structure remediation diagnostics
        struct_cmp = compare_template_structure(
            doc,
            {
                "schema_version": "1.0",
                "match": "exact",
                "ordered_modules": [{"kind": "abstract", "variant": "zh"}, {"kind": "chapters", "outline_levels": [1]}],
            },
        )
        self.assertEqual(struct_cmp["status"], "failed")
        self.assertTrue(struct_cmp.get("remediation"))

        # Run pipeline with render=True, then resume with ONLY visual_report (omitting profile/resolutions/render/overwrite)
        profile = template_profile(source)
        profile_path = self.root / "profile.json"
        decisions_path = self.root / "decisions.json"
        profile_path.write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")
        decisions_path.write_text(
            json.dumps(
                {i["id"]: {"decision": "ignore", "reason": "测试保留"} for i in profile["issues"]},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        real_run = subprocess.run

        def mock_render_run(command, **kwargs):
            if str(command[1]).endswith("render_docx.py"):
                from pipeline_state import digest
                from PIL import Image

                render_dir = reports / "render"
                render_dir.mkdir(parents=True, exist_ok=True)
                page1 = render_dir / "page-1.png"
                Image.new("RGB", (32, 32), "white").save(page1)
                out_sha = digest(output)
                manifest = {
                    "schema_version": "1.0",
                    "source": str(output),
                    "source_sha256": out_sha,
                    "selected_method": "word",
                    "page_count": 1,
                    "pages": [{"number": 1, "path": str(page1), "sha256": digest(page1)}],
                }
                manifest_path = render_dir / "render-manifest.json"
                manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
                tpl_visual = build_visual_report_template(
                    source_sha256=out_sha,
                    manifest_path=str(manifest_path),
                    manifest_sha256=digest(manifest_path),
                    selected_method="word",
                    page_count=1,
                )
                (reports / "visual-report.template.json").write_text(
                    json.dumps(tpl_visual, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                return subprocess.CompletedProcess(command, 0)
            return real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=mock_render_run):
            pipeline.run_pipeline(
                source,
                output,
                profile=profile_path,
                resolutions=decisions_path,
                match_structure=True,
                report_dir=reports,
                overwrite=True,
                render=True,
            )

        summary_pending = json.loads((reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary_pending["status"], "pending")
        self.assertIn("visual_defect_playbook", summary_pending["remediation"])
        self.assertIn("cover_overflow_two_pages", VISUAL_DEFECT_PLAYBOOK)

        # 1) Unedited visual-report.template.json (status=pending_visual_review) MUST be rejected!
        visual_tpl_path = reports / "visual-report.template.json"
        self.assertTrue(visual_tpl_path.is_file())
        tpl_data = json.loads(visual_tpl_path.read_text(encoding="utf-8"))
        self.assertEqual(tpl_data["status"], "pending_visual_review")
        self.assertIn("checklist", tpl_data)
        with self.assertRaises(subprocess.CalledProcessError):
            pipeline.run_pipeline(
                source,
                output,
                report_dir=reports,
                visual_report=visual_tpl_path,
                resume=True,
            )

        # 2) Changing status to passed without updating placeholder summary MUST also be rejected!
        tpl_data["status"] = "passed"
        visual_tpl_path.write_text(json.dumps(tpl_data, ensure_ascii=False, indent=2), encoding="utf-8")
        with self.assertRaises(subprocess.CalledProcessError):
            pipeline.run_pipeline(
                source,
                output,
                report_dir=reports,
                visual_report=visual_tpl_path,
                resume=True,
            )

        # 3) After inspecting pages and filling a concrete summary + status=passed, sign-off succeeds!
        tpl_data["summary"] = "已逐页检查全部 1 页 PNG 渲染图，封面、对齐、页眉页脚与编号均正常。"
        visual_tpl_path.write_text(json.dumps(tpl_data, ensure_ascii=False, indent=2), encoding="utf-8")
        pipeline.run_pipeline(
            source,
            output,
            report_dir=reports,
            visual_report=visual_tpl_path,
            resume=True,
        )
        summary_passed = json.loads((reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary_passed["status"], "passed")
        self.assertEqual(summary_passed["checks"]["visual"], "passed")

    def test_skill_root_write_protection_and_pipeline_quality_gate(self):
        from json_output import SKILL_ROOT, ensure_not_in_skill_root, ensure_protected_paths
        from word_format.profiles import requirements_profile

        with self.assertRaisesRegex(ValueError, "技能安装目录"):
            ensure_not_in_skill_root(SKILL_ROOT / "templates/compatibility/heading-numbering.json")
        with self.assertRaisesRegex(ValueError, "技能安装目录"):
            ensure_protected_paths(SKILL_ROOT / "SKILL.md")

        # Test requirement_rules enhancements (论文标题, 摘要内容, 正文内容, 首行空两格)
        req_file = self.root / "req.txt"
        req_file.write_text(
            "论文标题：使用小二号黑体，居中放置。\n"
            "摘要内容楷体小四号。\n"
            "正文内容首行空两格，小四号字体，不加粗。\n",
            encoding="utf-8",
        )
        req_prof = requirements_profile(req_file)
        self.assertEqual(req_prof["issues"], [])
        actions_roles = {
            (r["action"], r["target"].get("semantic_role") or r["target"].get("role"))
            for r in req_prof["rules"]
        }
        self.assertIn(("font.size.set", "abstract_title"), actions_roles)
        self.assertIn(("font.east_asia.set", "abstract_body"), actions_roles)
        self.assertIn(("paragraph.first_line_indent.set", "body"), actions_roles)

        # Test pipeline quality gate blocking uncited references and passing once cited
        raw_source = self.root / "q_raw.docx"
        source = self.root / "q_src.docx"
        output = self.root / "q_out.docx"
        reports = self.root / "q_reports"
        doc = Document()
        doc.styles["Normal"].font.name = "Times New Roman"
        doc.styles["Normal"].font.size = Pt(12)
        doc.styles["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        doc.add_paragraph("摘  要")
        doc.add_paragraph("这是一段满足字数下限要求的中文摘要内容，用于验证流水线内置质量门禁。")
        doc.add_heading("绪 论", level=1)
        p_body = doc.add_paragraph("这是第一章正文内容，仅引用了第一篇文献")
        for tag, val in [
            ("t", "["),
            ("fldChar", "begin"),
            ("instrText", " REF WFM_Ref_1 \\h "),
            ("fldChar", "separate"),
            ("t", "1"),
            ("fldChar", "end"),
            ("t", "]"),
            ("t", "。"),
        ]:
            r = OxmlElement("w:r")
            el = OxmlElement(f"w:{tag}")
            if tag == "fldChar":
                el.set(qn("w:fldCharType"), val)
            else:
                el.text = val
            r.append(el)
            p_body._p.append(r)
        doc.add_heading("参考文献", level=1)
        doc.add_paragraph("[1] 张三. 大语言模型综述[J]. 计算机学报, 2024, 47(1): 1-15.")
        doc.add_paragraph("[2] 李四. 未被正文引用的悬空文献[J]. 软件学报, 2023, 34(2): 20-30.")
        doc.save(raw_source)
        apply_plan(
            build_plan(
                raw_source,
                source,
                {"spec": {"references": {"numbering_mode": "word-numbering"}}},
            )
        )

        profile = template_profile(source)
        profile_path = self.root / "q_profile.json"
        decisions_path = self.root / "q_decisions.json"
        profile_path.write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")
        decisions_path.write_text(
            json.dumps(
                {i["id"]: {"decision": "ignore", "reason": "测试保留"} for i in profile["issues"]},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        # Reference [2] is uncited -> require_all_references_cited must fail the pipeline!
        with self.assertRaises(subprocess.CalledProcessError):
            pipeline.run_pipeline(
                source,
                output,
                profile=profile_path,
                resolutions=decisions_path,
                report_dir=reports,
                overwrite=True,
                check_quality=True,
                min_chinese_characters=10,
                min_abstract_chinese_characters=10,
                max_abstract_chinese_characters=100,
                min_references=2,
                min_recent_references=2,
                recent_start_year=2022,
                require_all_references_cited=True,
            )
        failed_summary = json.loads((reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(failed_summary["status"], "failed")
        self.assertEqual(failed_summary["checks"]["content"], "failed")

    def test_toc_levels_tab_runs_rpr_order_multi_table_flexible_chapters_and_universal_base(self):
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from word_format import normalize_universal_chinese_base
        from word_format.fields import _sync_toc_entry_runs_from_styles
        from word_format.inspect import DocumentIndex
        from word_format.profiles import requirements_profile
        from word_format.toc import result_status
        from word_format.verification import read_complex_state

        # 1) Universal Chinese base normalization cleans ja-JP theme fonts and #365F91/#4F81BD colors
        dirty_doc = Document()
        h1_rpr = dirty_doc.styles["Heading 1"].element.get_or_add_rPr()
        c_el = OxmlElement("w:color")
        c_el.set(qn("w:val"), "365F91")
        h1_rpr.append(c_el)
        cleaned_doc = normalize_universal_chinese_base(dirty_doc)
        self.assertIs(cleaned_doc, dirty_doc)
        self.assertEqual(
            dirty_doc.styles["Heading 1"].element.find(qn("w:rPr")).find(qn("w:color")).get(qn("w:val")),
            "000000",
        )
        self.assertEqual(
            dirty_doc.styles["Normal"].element.find(qn("w:rPr")).find(qn("w:rFonts")).get(qn("w:eastAsia")),
            "宋体",
        )

        # 2) Requirement rules support TOC entries per level
        req_path = self.root / "toc_req.txt"
        req_path.write_text(
            "一级目录：小四号黑体，1.5倍行距。\n"
            "二级目录：小四号宋体。\n",
            encoding="utf-8",
        )
        req_prof = requirements_profile(req_path)
        self.assertEqual(req_prof["issues"], [])
        toc_targets = [
            (r["action"], r["target"].get("role"), r["target"].get("level"), r["params"].get("value"))
            for r in req_prof["rules"]
            if r["target"].get("role") == "toc"
        ]
        self.assertIn(("font.east_asia.set", "toc", 1, "黑体"), toc_targets)
        self.assertIn(("font.east_asia.set", "toc", 2, "宋体"), toc_targets)

        # 3) Build a sample with toc 1..2 (with pure <w:tab/> runs), 1 chapter, and 1 three-line table
        def add_toc_entry(doc, style_obj, text_left, page_num, left_indent_pt, east_asia_font):
            p = doc.add_paragraph(style=style_obj)
            p.paragraph_format.left_indent = Pt(left_indent_pt)
            p.paragraph_format.line_spacing = 1.5
            r1 = p.add_run(text_left)
            r1.font.name = "Times New Roman"
            r1.font.size = Pt(12)
            r1._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia_font)
            # Pure <w:tab/> run without rPr (or with misplaced rPr after w:tab)
            r_tab = OxmlElement("w:r")
            r_tab.append(OxmlElement("w:tab"))
            p._p.append(r_tab)
            r2 = p.add_run(str(page_num))
            r2.font.name = "Times New Roman"
            r2.font.size = Pt(12)
            r2._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia_font)
            return p

        def make_three_line_table(doc, caption_text):
            cap = doc.add_paragraph(caption_text)
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl = doc.add_table(rows=2, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.cell(0, 0).text = "指标"
            tbl.cell(0, 1).text = "数值"
            tbl.cell(1, 0).text = "准确率"
            tbl.cell(1, 1).text = "99.1%"
            tbl_pr = tbl._tbl.tblPr
            borders = OxmlElement("w:tblBorders")
            for edge, val, sz in [
                ("top", "single", "12"),
                ("bottom", "single", "12"),
                ("left", "none", "0"),
                ("right", "none", "0"),
                ("insideH", "none", "0"),
                ("insideV", "none", "0"),
            ]:
                el = OxmlElement(f"w:{edge}")
                el.set(qn("w:val"), val)
                el.set(qn("w:sz"), sz)
                borders.append(el)
            tbl_pr.append(borders)
            return tbl

        sample_path = self.root / "toc_multi_sample.docx"
        sdoc = Document()
        normalize_universal_chinese_base(sdoc)
        st_toc1 = sdoc.styles["toc 1"] if "toc 1" in sdoc.styles else sdoc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
        st_toc2 = sdoc.styles["toc 2"] if "toc 2" in sdoc.styles else sdoc.styles.add_style("toc 2", WD_STYLE_TYPE.PARAGRAPH)
        p_toc_head = sdoc.add_paragraph("目  录")
        p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_toc_entry(sdoc, st_toc1, "1 绪论", 1, 0, "黑体")
        add_toc_entry(sdoc, st_toc2, "1.1 研究背景", 1, 12, "宋体")
        add_toc_entry(sdoc, st_toc1, "参考文献", 5, 0, "黑体")
        sdoc.add_heading("1 绪论", level=1)
        sdoc.add_heading("1.1 研究背景", level=2)
        sdoc.add_paragraph("正文内容段落。")
        make_three_line_table(sdoc, "表1.1 样张基准实验对比表")
        sdoc.add_heading("参考文献", level=1)
        sdoc.add_paragraph("[1] 张三. 深度学习综述[J]. 计算机学报, 2025, 48(1): 1-10.")
        sdoc.save(sample_path)

        prof = template_profile(sample_path)
        # Ensure toc 1 and toc 2 did not trigger a false left_indent conflict
        toc_conflict_ids = [
            i["id"] for i in prof["issues"] if "toc" in str(i.get("scope", ""))
        ]
        self.assertEqual(toc_conflict_ids, [])
        # Ensure table rules use all: True rather than ordinal: 1
        table_rules = [r for r in prof["rules"] if r["target"].get("type") == "table"]
        self.assertTrue(table_rules)
        for tr in table_rules:
            self.assertTrue(tr["target"].get("all"))
            self.assertNotIn("ordinal", tr["target"])

        # 4) Build a draft with 2 chapters, 2 unformatted tables, and unformatted toc 1..2 with pure <w:tab/> runs
        draft_path = self.root / "toc_multi_draft.docx"
        out_path = self.root / "toc_multi_out.docx"
        ddoc = Document()
        normalize_universal_chinese_base(ddoc)
        dt_toc1 = ddoc.styles["toc 1"] if "toc 1" in ddoc.styles else ddoc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
        dt_toc2 = ddoc.styles["toc 2"] if "toc 2" in ddoc.styles else ddoc.styles.add_style("toc 2", WD_STYLE_TYPE.PARAGRAPH)
        dp_toc_head = ddoc.add_paragraph("目  录")
        dp_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_toc_entry(ddoc, dt_toc1, "1 绪论", 1, 0, "楷体")
        add_toc_entry(ddoc, dt_toc2, "1.1 研究背景", 1, 12, "楷体")
        add_toc_entry(ddoc, dt_toc1, "2 关键技术分析", 2, 0, "楷体")
        add_toc_entry(ddoc, dt_toc2, "2.1 核心架构", 2, 12, "楷体")
        add_toc_entry(ddoc, dt_toc1, "参考文献", 5, 0, "楷体")
        ddoc.add_heading("1 绪论", level=1)
        ddoc.add_heading("1.1 研究背景", level=2)
        ddoc.add_paragraph("第一章正文内容。")
        ddoc.add_paragraph("表1.1 第一章对比表")
        t1 = ddoc.add_table(rows=2, cols=2)
        t1.cell(0, 0).text = "A"
        t1.cell(0, 1).text = "B"
        t1.cell(1, 0).text = "1"
        t1.cell(1, 1).text = "2"
        ddoc.add_heading("2 关键技术分析", level=1)
        ddoc.add_heading("2.1 核心架构", level=2)
        ddoc.add_paragraph("第二章正文内容。")
        ddoc.add_paragraph("表2.1 第二章对比表")
        t2 = ddoc.add_table(rows=2, cols=2)
        t2.cell(0, 0).text = "C"
        t2.cell(0, 1).text = "D"
        t2.cell(1, 0).text = "3"
        t2.cell(1, 1).text = "4"
        ddoc.add_heading("参考文献", level=1)
        ddoc.add_paragraph("[1] 李四. 架构演进研究[J]. 软件学报, 2025, 36(2): 11-22.")
        ddoc.save(draft_path)

        # Strict mode rejects 2 chapters vs 1 chapter sample; flexible_chapters=True passes!
        strict_check = check_template_structure(draft_path, prof, flexible_chapters=False)
        self.assertEqual(strict_check["status"], "failed")
        flex_check = check_template_structure(draft_path, prof, flexible_chapters=True)
        self.assertEqual(flex_check["status"], "passed")

        # Apply profile_plan with flexible_chapters=True
        resolutions = {
            i["id"]: {"decision": "ignore", "reason": "test"}
            for i in prof["issues"]
        }
        plan = profile_plan(
            draft_path,
            out_path,
            prof,
            resolutions=resolutions,
            match_structure=True,
            flexible_chapters=True,
        )
        self.assertNotIn("status", plan, f"Expected executable plan, got: {plan.get('issues')}")
        applied = apply_plan(plan, overwrite=True)
        verified = verify_plan(plan, applied)
        self.assertEqual(verified["status"], "passed")

        # Verify output:
        # - TOC level 1 uses 黑体, TOC level 2 uses 宋体 (including pure <w:tab/> runs)
        # - All <w:rPr> are at child index 0 of <w:r>
        # - Both tables (t1 and t2) are centered three-line tables
        # - styles.xml toc 1 and toc 2 styles are synchronized
        out_index = DocumentIndex(Document(out_path))
        toc_state = read_complex_state(out_index, "table_of_contents")
        self.assertEqual(len(toc_state["entry_formats"]), 5)
        self.assertEqual(toc_state["entry_formats"][0]["font"][0]["font_east_asia"], "黑体")
        self.assertEqual(toc_state["entry_formats"][1]["font"][0]["font_east_asia"], "宋体")
        self.assertTrue(toc_state["results_by_level"]["2"]["valid"])
        toc_res = result_status(out_index.document, 2)
        self.assertTrue(toc_res["valid"])

        for r in out_index.records.values():
            if r.get("kind") == "paragraph" and r.get("role") == "toc":
                self.assertIsNone(r.get("heading_level"))
                expected_ea = "黑体" if r.get("toc_level") == 1 else "宋体"
                self.assertEqual(out_index.font_values(r["id"], "font_east_asia"), [expected_ea])
                p_el = out_index.elements[r["id"]]
                for run in p_el.findall(qn("w:r")):
                    rpr = run.find(qn("w:rPr"))
                    if rpr is not None:
                        self.assertEqual(list(run).index(rpr), 0)

        from word_format.registry import REGISTRY

        tbl_recs = [r for r in out_index.records.values() if r.get("kind") == "table"]
        self.assertEqual(len(tbl_recs), 2)
        for tbl_rec in tbl_recs:
            self.assertEqual(
                REGISTRY["table.alignment.set"].reader(
                    out_index, tbl_rec, REGISTRY["table.alignment.set"]
                ),
                "center",
            )
            self.assertTrue(
                REGISTRY["table.three_line.apply"].reader(
                    out_index, tbl_rec, REGISTRY["table.three_line.apply"]
                )
            )

        # Verify _sync_toc_entry_runs_from_styles re-applies toc style fonts after simulated COM update
        sim_doc = Document(out_path)
        for p in sim_doc.paragraphs:
            if p.style and p.style.name.lower() in {"toc 1", "toc 2"}:
                for run in p.runs:
                    rpr = run._r.find(qn("w:rPr"))
                    if rpr is not None:
                        rfonts = rpr.find(qn("w:rFonts"))
                        if rfonts is not None:
                            rfonts.attrib.pop(qn("w:eastAsia"), None)
        _sync_toc_entry_runs_from_styles(sim_doc)
        sim_out = self.root / "toc_multi_sim.docx"
        sim_doc.save(sim_out)
        sim_index = DocumentIndex(Document(sim_out))
        for r in sim_index.records.values():
            if r.get("kind") == "paragraph" and r.get("role") == "toc":
                expected_ea = "黑体" if r.get("toc_level") == 1 else "宋体"
                self.assertEqual(sim_index.font_values(r["id"], "font_east_asia"), [expected_ea])


if __name__ == "__main__":
    unittest.main()
