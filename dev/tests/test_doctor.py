"""Environment diagnosis works even before third-party packages can import."""
import json
from pathlib import Path
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[2]


class DoctorTests(unittest.TestCase):
    def test_no_site_packages_returns_actionable_dependency_state(self):
        result = subprocess.run([sys.executable, "-S", "-B", str(ROOT / "scripts/format_cli.py"), "doctor"],
                                capture_output=True, encoding="utf-8")
        self.assertEqual(result.returncode, 2, result.stderr)
        data = json.loads(result.stdout)
        self.assertFalse(data["capabilities"]["core_formatting"])
        self.assertIn("python-docx", data["missing_dependencies"])
        self.assertTrue(data["next_actions"])

    def test_doctor_separates_runtime_readiness_from_visual_review(self):
        result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/format_cli.py"), "doctor"],
                                capture_output=True, encoding="utf-8")
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(result.stdout)
        self.assertTrue(data["capabilities"]["core_formatting"])
        self.assertTrue(data["capabilities"]["default_report"])
        self.assertEqual(data["visual_review"], "not_performed")

    def test_missing_report_template_has_a_stable_code(self):
        sys.path.insert(0, str(ROOT / "scripts"))
        from word_format.reports import load_report_catalog
        import tempfile
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "reports.json"
            path.write_text(json.dumps({"schema_version": "1.0", "default": "missing", "templates": [
                {"id": "missing", "path": "missing.docx", "sha256": "0" * 64,
                 "source_notes": "SOURCE.md", "usage_notes": "usage.md"}]}), encoding="utf-8")
            try:
                load_report_catalog(path)
            except ValueError as exc:
                self.assertEqual(getattr(exc, "code", None), "REPORT_TEMPLATE_MISSING")
            else:
                self.fail("missing template was accepted")


if __name__ == "__main__":
    unittest.main()
