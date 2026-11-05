"""Recover from late failures without writing the formatted document twice."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))
import run_format_pipeline as pipeline
from word_format.profiles import template_profile


class RecoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.source, self.output = self.work / "source.docx", self.work / "output.docx"
        self.reports = self.work / "reports"
        doc = Document()
        doc.styles["Normal"].font.name = "宋体"
        doc.styles["Normal"].font.size = Pt(12)
        doc.add_heading("引言", 1)
        doc.add_paragraph("恢复测试正文。")
        doc.save(self.source)
        profile = template_profile(self.source)
        self.profile, self.decisions = self.work / "profile.json", self.work / "decisions.json"
        self.profile.write_text(json.dumps(profile), encoding="utf-8")
        self.decisions.write_text(json.dumps({i["id"]: {"decision": "ignore", "reason": "测试样例未定义属性保留"}
                                              for i in profile["issues"]}), encoding="utf-8")

    def run_pipeline(self, **kwargs):
        return pipeline.run_pipeline(self.source, self.output, profile=self.profile,
                                     resolutions=self.decisions, match_structure=True,
                                     report_dir=self.reports, **kwargs)

    def fail_render(self, command, **kwargs):
        if str(command[1]).endswith("render_docx.py"):
            raise subprocess.CalledProcessError(2, command)
        return self.real_run(command, **kwargs)

    def test_render_failure_resumes_without_reapplying(self):
        self.real_run = subprocess.run
        with patch.object(pipeline.subprocess, "run", side_effect=self.fail_render):
            with self.assertRaises(subprocess.CalledProcessError):
                self.run_pipeline(render=True)
        state_path = self.reports / "pipeline-state.json"
        self.assertTrue(state_path.is_file(), "pipeline must preserve stage progress")
        before = self.output.read_bytes()
        calls = []

        def resume_run(command, **kwargs):
            calls.append(command)
            if str(command[1]).endswith("render_docx.py"):
                target = self.reports / "render"
                target.mkdir(exist_ok=True)
                from pipeline_state import digest
                (target / "render-manifest.json").write_text(json.dumps({"source": str(self.output),
                    "source_sha256": digest(self.output), "pages": []}), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0)
            return self.real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=resume_run):
            self.run_pipeline(render=True, resume=True)
        self.assertFalse(any("apply" in c for c in calls))
        self.assertEqual(before, self.output.read_bytes())
        summary = json.loads((self.reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["checks"]["visual"], "pending")
        self.assertEqual(summary["checks"]["render"], "passed")

    def test_resume_rejects_modified_output_and_missing_apply_evidence(self):
        self.run_pipeline()
        original = self.output.read_bytes()
        (self.reports / "apply.json").unlink()
        with self.assertRaisesRegex(ValueError, "证据|摘要|产物"):
            self.run_pipeline(resume=True)
        self.assertEqual(self.output.read_bytes(), original)

    def test_summary_does_not_claim_visual_or_content_acceptance(self):
        self.run_pipeline()
        path = self.reports / "summary.json"
        self.assertTrue(path.is_file())
        summary = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(summary["checks"]["visual"], "skipped")
        self.assertEqual(summary["checks"]["content"], "not_performed")

    def test_resume_without_state_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "状态"):
            self.run_pipeline(resume=True)
        self.assertFalse(self.output.exists())

    def test_render_directory_cannot_contain_an_input(self):
        target = self.reports / "render"
        target.mkdir(parents=True)
        moved = target / "page-1.png"
        moved.write_bytes(self.decisions.read_bytes())
        self.decisions = moved
        original = moved.read_bytes()
        with self.assertRaisesRegex(ValueError, "渲染|输入"):
            self.run_pipeline(render=True)
        self.assertEqual(moved.read_bytes(), original)
        self.assertFalse(self.output.exists())

    def test_failed_visual_stage_is_reported_as_failed(self):
        # Force the verifier to reject a review, without depending on office software.
        review = self.work / "review.json"
        review.write_text(json.dumps({"status": "failed", "model": {"id": "synthetic", "supports_image_input": True},
                                      "reviewed_pages": [1], "findings": [{"page": 1, "issue": "test"}]}), encoding="utf-8")
        real_run = subprocess.run

        def command_run(command, **kwargs):
            if str(command[1]).endswith("render_docx.py"):
                from pipeline_state import digest
                target = self.reports / "render"
                target.mkdir(exist_ok=True)
                (target / "render-manifest.json").write_text(json.dumps({"source": str(self.output),
                    "source_sha256": digest(self.output), "pages": []}), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0)
            return real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=command_run):
            with self.assertRaises(subprocess.CalledProcessError):
                self.run_pipeline(render=True, visual_report=review)
        summary = json.loads((self.reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["checks"]["visual"], "failed")

    def test_saved_visual_evidence_is_checked_when_review_argument_is_omitted(self):
        from pipeline_state import digest
        review = self.work / "review.json"
        review.write_text(json.dumps({"status": "skipped", "model": {
            "id": "text-model", "supports_image_input": False}}), encoding="utf-8")
        real_run = subprocess.run

        def command_run(command, **kwargs):
            if str(command[1]).endswith("render_docx.py"):
                target = self.reports / "render"
                target.mkdir(exist_ok=True)
                (target / "render-manifest.json").write_text(json.dumps({"source": str(self.output),
                    "source_sha256": digest(self.output), "pages": []}), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0)
            return real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=command_run):
            self.run_pipeline(render=True, visual_report=review)
        visual_path = self.reports / "visual-verification.json"
        data = json.loads(visual_path.read_text(encoding="utf-8"))
        data["visual"]["status"] = "passed"
        visual_path.write_text(json.dumps(data), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "摘要|证据"):
            self.run_pipeline(render=True, resume=True)
        summary = json.loads((self.reports / "summary.json").read_text(encoding="utf-8"))
        self.assertNotEqual(summary["checks"]["visual"], "passed")

    def test_changed_output_invalidates_downstream_summary(self):
        self.run_pipeline()
        self.output.write_bytes(self.output.read_bytes() + b"changed")
        with self.assertRaisesRegex(ValueError, "摘要|证据"):
            self.run_pipeline(resume=True)
        summary = json.loads((self.reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["failure"]["code"], "EVIDENCE_STALE")
        self.assertNotEqual(summary["checks"]["format"], "passed")

    def test_timeout_with_binary_stderr_can_retry_and_saves_json_state(self):
        from pipeline_state import digest
        real_run = subprocess.run
        attempts = []

        def command_run(command, **kwargs):
            if str(command[1]).endswith("render_docx.py"):
                attempts.append(1)
                if len(attempts) == 1:
                    raise subprocess.TimeoutExpired(command, 1, stderr=b"renderer timed out")
                target = self.reports / "render"
                target.mkdir(exist_ok=True)
                (target / "render-manifest.json").write_text(json.dumps({"source": str(self.output),
                    "source_sha256": digest(self.output), "pages": []}), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0)
            return real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=command_run):
            self.run_pipeline(render=True, render_retries=1)
        state = json.loads((self.reports / "pipeline-state.json").read_text(encoding="utf-8"))
        self.assertEqual(state["stages"]["render"]["attempts"], 2)

    def test_corrupt_render_manifest_invalidates_render_and_visual(self):
        from pipeline_state import digest
        real_run = subprocess.run

        def command_run(command, **kwargs):
            if str(command[1]).endswith("render_docx.py"):
                target = self.reports / "render"
                target.mkdir(exist_ok=True)
                (target / "render-manifest.json").write_text(json.dumps({"source": str(self.output),
                    "source_sha256": digest(self.output), "pages": []}), encoding="utf-8")
                return subprocess.CompletedProcess(command, 0)
            return real_run(command, **kwargs)

        with patch.object(pipeline.subprocess, "run", side_effect=command_run):
            self.run_pipeline(render=True)
        (self.reports / "render/render-manifest.json").write_text("{bad json", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "证据|摘要"):
            self.run_pipeline(render=True, resume=True)
        summary = json.loads((self.reports / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["failure"]["code"], "EVIDENCE_STALE")
        self.assertNotEqual(summary["checks"]["render"], "passed")


if __name__ == "__main__":
    unittest.main()
