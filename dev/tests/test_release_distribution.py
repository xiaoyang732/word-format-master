"""The installed archive must work without source files or a warm cache."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]


class ReleaseTests(unittest.TestCase):
    def test_native_and_encoded_archives_prepare_reports_in_isolation(self):
        for mode in ("native", "json"):
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as folder:
                work = Path(folder)
                archive = work / "release.zip"
                result = subprocess.run([sys.executable, str(ROOT / "dev/tools/build_release.py"),
                                         "--output", str(archive), "--template-mode", mode],
                                        capture_output=True, encoding="utf-8")
                self.assertEqual(result.returncode, 0, result.stderr)
                target = work / "中文 安装目录"
                with ZipFile(archive) as package:
                    names = package.namelist()
                    self.assertIn("templates/reference/general-course-report/格式与使用说明.md", names)
                    self.assertFalse(any(n.startswith("dev/") for n in names))
                    package.extractall(target)
                env = dict(os.environ, TMP=str(work), TEMP=str(work), TMPDIR=str(work), PYTHONPATH="")
                task = work / "报告任务"
                result = subprocess.run([sys.executable, "-E", str(target / "scripts/format_cli.py"),
                                         "report", "prepare", str(task)], cwd=work, env=env,
                                        capture_output=True, encoding="utf-8")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(json.loads((task / "workflow.json").read_text(encoding="utf-8"))["status"], "draft")
                self.assertTrue((task / "draft.docx").is_file())

    def test_validator_rejects_an_archive_with_missing_usage_notes(self):
        with tempfile.TemporaryDirectory() as folder:
            good, bad = Path(folder) / "good.zip", Path(folder) / "bad.zip"
            result = subprocess.run([sys.executable, str(ROOT / "dev/tools/build_release.py"), "--output", str(good)],
                                    capture_output=True, encoding="utf-8")
            self.assertEqual(result.returncode, 0, result.stderr)
            with ZipFile(good) as source, ZipFile(bad, "w") as dest:
                for name in source.namelist():
                    if not name.endswith("格式与使用说明.md"):
                        dest.writestr(name, source.read(name))
            result = subprocess.run([sys.executable, str(ROOT / "dev/tools/validate_release.py"), str(bad)],
                                    capture_output=True, encoding="utf-8")
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("格式与使用说明", result.stderr)


if __name__ == "__main__":
    unittest.main()
