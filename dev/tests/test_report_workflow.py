"""Default report selection must reach the existing evidence/format pipeline."""

import json
import subprocess
import sys
import tempfile
import unittest
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
from word_format.profiles import template_profile
from word_format.contracts import FormatError


class ReportWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)

    def cli(self, *args):
        return subprocess.run(
            [sys.executable, str(ROOT / "scripts/format_cli.py"), *map(str, args)],
            capture_output=True, encoding="utf-8", check=False,
        )

    def test_default_is_selected_without_a_template_argument(self):
        result = self.cli("report", "list")
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertEqual(data["default"], "general-course-report")
        self.assertTrue(Path(data["templates"][0]["path"]).is_file())

    def test_prepare_clones_selected_template_and_profiles_the_source(self):
        destination = self.work / "报告任务"
        result = self.cli("report", "prepare", destination, "--topic", "计算机网络")
        self.assertEqual(result.returncode, 0, result.stderr)
        manifest = json.loads((destination / "workflow.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["template_id"], "general-course-report")
        self.assertEqual(manifest["topic"], "计算机网络")
        self.assertEqual(manifest["status"], "draft")
        source = Path(manifest["source"]["path"])
        draft = Path(manifest["draft"])
        self.assertEqual(source.read_bytes(), draft.read_bytes())
        profile = json.loads(Path(manifest["profile"]).read_text(encoding="utf-8"))
        self.assertEqual(profile["source"]["sha256"], sha256(source.read_bytes()).hexdigest())
        self.assertEqual(profile["source"]["path"], str(source))
        self.assertEqual(manifest["required_plan_options"], ["--match-structure"])
        with ZipFile(draft) as archive:
            self.assertIn('第%1章', archive.read("word/numbering.xml").decode("utf-8"))
            xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("TOC", xml)
            self.assertIn("tblBorders", xml)

    def test_explicit_sample_takes_priority_over_default(self):
        sample = self.work / "教师模板.docx"
        doc = Document()
        doc.add_heading("自定报告章节", 1)
        doc.add_paragraph("教师要求")
        doc.save(sample)
        result = self.cli("report", "prepare", self.work / "custom", "--sample", sample)
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertIsNone(data["template_id"])
        self.assertEqual(Path(data["source"]["path"]), sample)
        self.assertEqual(Path(data["draft"]).read_bytes(), sample.read_bytes())

    def test_existing_workspace_and_unknown_template_are_rejected_without_writes(self):
        workspace = self.work / "existing"
        workspace.mkdir()
        original = workspace / "draft.docx"
        original.write_bytes(b"user original")
        result = self.cli("report", "prepare", workspace)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(original.read_bytes(), b"user original")
        self.assertEqual(list(workspace.iterdir()), [original])
        unknown = self.work / "unknown"
        result = self.cli("report", "prepare", unknown, "--template", "missing-report")
        self.assertEqual(result.returncode, 2)
        self.assertFalse(unknown.exists())

    def test_catalog_rejects_changed_templates_and_paths_outside_library(self):
        from word_format.reports import load_report_catalog

        library = self.work / "templates"
        library.mkdir()
        sample = library / "sample.docx"
        Document().save(sample)
        notes = library / "notes.md"
        notes.write_text("来源和使用说明", encoding="utf-8")
        entry = {"id": "test-report", "path": "sample.docx",
                 "sha256": sha256(sample.read_bytes()).hexdigest(),
                 "source_notes": "notes.md", "usage_notes": "notes.md"}
        data = {"schema_version": "1.0", "default": "test-report", "templates": [entry]}
        catalog = library / "reports.json"
        catalog.write_text(json.dumps(data), encoding="utf-8")
        self.assertEqual(load_report_catalog(catalog)["default"], "test-report")
        entry["sha256"] = "0" * 64
        catalog.write_text(json.dumps(data), encoding="utf-8")
        with self.assertRaisesRegex(FormatError, "hash mismatch"):
            load_report_catalog(catalog)
        outside = self.work / "outside.docx"
        Document().save(outside)
        entry.update(path="../outside.docx", sha256=sha256(outside.read_bytes()).hexdigest())
        catalog.write_text(json.dumps(data), encoding="utf-8")
        with self.assertRaisesRegex(FormatError, "Invalid or missing"):
            load_report_catalog(catalog)

    def test_prepare_cannot_write_inside_template_library(self):
        forbidden = ROOT / "templates" / "not-a-report-task"
        result = self.cli("report", "prepare", forbidden)
        self.assertEqual(result.returncode, 2)
        self.assertFalse(forbidden.exists())

    def test_report_json_output_cannot_overwrite_catalog(self):
        catalog = ROOT / "templates/reports.json"
        original = catalog.read_bytes()
        result = self.cli("report", "list", "--output", catalog)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(catalog.read_bytes(), original)

    def test_profile_pipeline_checks_structure_and_preserves_native_numbering(self):
        sample = self.work / "sample.docx"
        doc = Document()
        doc.styles["Normal"].font.name = "宋体"
        doc.styles["Normal"].font.size = Pt(12)
        doc.add_heading("引言", 1)
        doc.add_paragraph("计算机课程报告正文。")
        doc.save(sample)
        profile = template_profile(sample)
        profile_path = self.work / "profile.json"
        profile_path.write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")
        decisions = {i["id"]: {"decision": "ignore", "reason": "合成测试样例未定义的属性保持原值"}
                     for i in profile["issues"]}
        resolutions = self.work / "decisions.json"
        resolutions.write_text(json.dumps(decisions, ensure_ascii=False), encoding="utf-8")
        output = self.work / "report.docx"
        reports = self.work / "checks"
        result = subprocess.run([
            sys.executable, str(ROOT / "scripts/run_format_pipeline.py"), str(sample), str(output),
            "--profile", str(profile_path), "--match-structure", "--resolutions", str(resolutions),
            "--report-dir", str(reports),
        ], capture_output=True, encoding="utf-8", check=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(output.is_file())
        verification = json.loads((reports / "verification.json").read_text(encoding="utf-8"))
        self.assertEqual(verification["status"], "passed")
        self.assertEqual(verification["visual"]["status"], "skipped")
        self.assertEqual(json.loads((reports / "structure.json").read_text(encoding="utf-8"))["status"], "passed")
        self.assertEqual(json.loads((reports / "numbering.json").read_text(encoding="utf-8"))["status"], "passed")
        # A new chapter must block publication even when format decisions are supplied.
        doc.add_heading("新增章节", 1)
        doc.save(self.work / "extra.docx")
        result = subprocess.run([
            sys.executable, str(ROOT / "scripts/run_format_pipeline.py"),
            str(self.work / "extra.docx"), str(self.work / "invalid.docx"),
            "--profile", str(profile_path), "--match-structure", "--resolutions", str(resolutions),
            "--report-dir", str(self.work / "bad-checks"),
        ], capture_output=True, encoding="utf-8", check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse((self.work / "invalid.docx").exists())

    def test_pipeline_report_artifacts_cannot_overwrite_profile_inputs(self):
        sample = self.work / "sample.docx"
        doc = Document()
        doc.add_heading("引言", 1)
        doc.add_paragraph("正文")
        doc.save(sample)
        profile = template_profile(sample)
        decisions = {i["id"]: {"decision": "ignore", "reason": "合成样例保持未定义格式"}
                     for i in profile["issues"]}
        for kind in ("profile", "resolutions"):
            with self.subTest(input_kind=kind):
                reports = self.work / kind
                reports.mkdir()
                profile_path = reports / ("numbering.json" if kind == "profile" else "profile.json")
                resolutions = reports / ("verification.json" if kind == "resolutions" else "decisions.json")
                profile_path.write_text(json.dumps(profile), encoding="utf-8")
                resolutions.write_text(json.dumps(decisions), encoding="utf-8")
                original_profile, original_decisions = profile_path.read_bytes(), resolutions.read_bytes()
                output = self.work / (kind + ".docx")
                result = subprocess.run([
                    sys.executable, str(ROOT / "scripts/run_format_pipeline.py"), str(sample), str(output),
                    "--profile", str(profile_path), "--match-structure", "--resolutions", str(resolutions),
                    "--report-dir", str(reports),
                ], capture_output=True, encoding="utf-8", check=False)
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(profile_path.read_bytes(), original_profile)
                self.assertEqual(resolutions.read_bytes(), original_decisions)
                self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
