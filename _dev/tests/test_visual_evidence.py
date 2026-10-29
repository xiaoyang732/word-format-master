"""Visual acceptance must bind the output, manifest and complete page inventory."""

import copy
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from word_format import apply_plan, build_plan, verify_plan
from word_format.contracts import FormatError, sha256
from word_format.visual import verify_visual_evidence


class VisualEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "原稿.docx"
        self.output = self.folder / "排版后.docx"
        doc = Document()
        doc.add_paragraph("测试正文")
        doc.save(self.source)
        self.plan = build_plan(self.source, self.output, {"operations": [
            {"action": "font.size.set", "target": {"type": "paragraph", "all": True}, "params": {"value": 12}}
        ]})
        self.applied = apply_plan(self.plan)
        # These are synthetic hash fixtures, not a claim of actual rendering.
        self.pages = []
        for i in (1, 2):
            page = self.folder / f"page-{i}.png"
            Image.new("RGB", (8, 8), (i, i, i)).save(page)
            self.pages.append({"number": i, "path": str(page), "sha256": sha256(page)})
        self.manifest_path = self.folder / "render-manifest.json"
        self.manifest = {"schema_version": "1.0", "source": str(self.output),
                         "source_sha256": sha256(self.output), "selected_method": "libreoffice",
                         "page_count": 2, "pages": self.pages}
        self.save_manifest()
        self.visual = {"status": "passed", "model": {"id": "synthetic-test", "supports_image_input": True},
                       "output_sha256": sha256(self.output), "manifest_sha256": sha256(self.manifest_path),
                       "renderer": {"method": "libreoffice"}, "reviewed_pages": [1, 2], "findings": []}

    def save_manifest(self):
        self.manifest_path.write_text(json.dumps(self.manifest), encoding="utf-8")

    def test_bound_complete_report_and_explicit_skipped_default(self):
        result = verify_plan(self.plan, self.applied, visual_report=self.visual, render_manifest=self.manifest_path)
        self.assertEqual(result["status"], "passed")
        self.assertEqual(result["visual"]["page_count"], 2)
        self.assertEqual(verify_plan(self.plan, self.applied)["visual"]["status"], "skipped")

    def test_changed_or_missing_pages_and_stale_manifest_are_rejected(self):
        original = Path(self.pages[0]["path"]).read_bytes()
        for change in ("changed", "missing", "stale"):
            with self.subTest(change=change):
                page = Path(self.pages[0]["path"])
                page.write_bytes(original)
                self.save_manifest()
                if change == "changed":
                    page.write_bytes(b"changed")
                elif change == "missing":
                    page.unlink()
                else:
                    self.manifest_path.write_text("{}", encoding="utf-8")
                with self.assertRaises(FormatError):
                    verify_visual_evidence(self.output, self.visual, self.manifest_path)

    def test_wrong_source_output_hash_and_renderer_are_rejected(self):
        for update in ({"source": str(self.source)}, {"source_sha256": "0" * 64},
                       {"selected_method": "word"}):
            with self.subTest(update=update):
                changed = {**self.manifest, **update}
                self.manifest_path.write_text(json.dumps(changed), encoding="utf-8")
                visual = {**self.visual, "manifest_sha256": sha256(self.manifest_path)}
                with self.assertRaises(FormatError):
                    verify_visual_evidence(self.output, visual, self.manifest_path)
        self.save_manifest()
        with self.assertRaises(FormatError):
            verify_visual_evidence(self.output, {**self.visual, "output_sha256": "0" * 64}, self.manifest_path)

    def test_every_page_must_be_reviewed_once_as_integer(self):
        for pages in ([1], [1, 1], [True, 2], [1.0, 2], []):
            with self.subTest(pages=pages):
                visual = {**self.visual, "reviewed_pages": pages}
                result = verify_plan(self.plan, self.applied, visual_report=visual, render_manifest=self.manifest_path)
                self.assertEqual(result["status"], "failed")
                self.assertEqual(result["visual"]["status"], "failed")

    def test_manifest_inventory_cannot_duplicate_pages_or_use_boolean_numbers(self):
        for pages in ([self.pages[0], self.pages[0]], [{**self.pages[0], "number": True}, self.pages[1]],
                      [{**self.pages[0], "number": 1.0}, self.pages[1]]):
            with self.subTest(pages=pages):
                changed = {**self.manifest, "pages": pages}
                self.manifest_path.write_text(json.dumps(changed), encoding="utf-8")
                visual = {**self.visual, "manifest_sha256": sha256(self.manifest_path)}
                with self.assertRaises(FormatError):
                    verify_visual_evidence(self.output, visual, self.manifest_path)

    def test_missing_image_capability_is_skipped_and_renderer_preference_is_honored(self):
        for support in (False, None):
            visual = {**self.visual, "model": {"id": "text-only", "supports_image_input": support}}
            self.assertEqual(verify_visual_evidence(self.output, visual)["status"], "skipped")
        request = copy.deepcopy(self.plan["request"])
        request["verification"] = {"render_method": "word"}
        plan = build_plan(self.source, self.output, request)
        applied = apply_plan(plan, overwrite=True)
        self.manifest["source_sha256"] = sha256(self.output)
        self.save_manifest()
        visual = {**self.visual, "output_sha256": sha256(self.output), "manifest_sha256": sha256(self.manifest_path)}
        with self.assertRaises(FormatError):
            verify_plan(plan, applied, visual_report=visual, render_manifest=self.manifest_path)
        self.assertEqual(verify_plan(plan, applied, visual_report=visual, render_manifest=self.manifest_path,
                                     render_method="libreoffice")["visual"]["status"], "passed")

    def test_cli_checks_visual_evidence_and_protects_page_files(self):
        plan_path, report_path, visual_path = [self.folder / name for name in ("plan.json", "report.json", "visual.json")]
        for path, value in ((plan_path, self.plan), (report_path, self.applied), (visual_path, self.visual)):
            path.write_text(json.dumps(value), encoding="utf-8")
        command = [sys.executable, str(ROOT / "scripts/format_cli.py"), "verify", "--plan", str(plan_path),
                   "--report", str(report_path), "--visual-report", str(visual_path), "--render-manifest", str(self.manifest_path)]
        result = subprocess.run(command, capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode("utf-8"))
        self.assertEqual(json.loads(result.stdout)["visual"]["status"], "passed")
        image = Path(self.pages[0]["path"])
        before = image.read_bytes()
        result = subprocess.run(command + ["--output", str(image)], capture_output=True, timeout=30)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(image.read_bytes(), before)
