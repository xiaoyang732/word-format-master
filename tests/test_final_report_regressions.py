"""Independent regressions for the cross-checked October review findings."""

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from json_output import write_json
from word_format import apply_plan, build_plan, verify_plan
from word_format.contracts import FormatError
from word_format.registry import REGISTRY
from word_format.properties import xml_value


class JsonSafetyTests(unittest.TestCase):
    def test_cli_collisions_preserve_inputs_and_do_not_create_outputs(self):
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            source = folder / "source.docx"
            doc = Document()
            doc.add_paragraph("Source must survive")
            doc.save(source)
            original = source.read_bytes()
            alias = folder / "alias.json"
            os.link(source, alias)
            handoff = folder / "handoff.json"
            handoff.write_text("{}", encoding="utf-8")
            output = folder / "result.docx"
            for report in (source, alias, folder / "sub" / ".." / "source.docx"):
                for script, args in (
                    ("analyze_docx.py", [source, "--output", report]),
                    ("analyze_citations.py", [source, "--output", report]),
                    ("apply_spec.py", [source, output, "--handoff", handoff, "--report", report]),
                    ("format_cli.py", ["inspect", source, "--output", report]),
                ):
                    with self.subTest(script=script, report=report):
                        result = subprocess.run([sys.executable, str(ROOT / "scripts" / script), *map(str, args)],
                                                capture_output=True, timeout=30)
                        self.assertNotEqual(result.returncode, 0)
                        self.assertEqual(hashlib.sha256(source.read_bytes()).digest(), hashlib.sha256(original).digest())
                        self.assertTrue(zipfile.is_zipfile(source))
                        self.assertFalse(output.exists())
            requirements = folder / "requirements.txt"
            requirements.write_text("正文宋体", encoding="utf-8")
            result = subprocess.run([sys.executable, str(ROOT / "scripts/parse_requirements.py"),
                                     str(requirements), "--output", str(requirements)], capture_output=True, timeout=30)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(requirements.read_text(encoding="utf-8"), "正文宋体")
            result = subprocess.run([sys.executable, str(ROOT / "scripts/apply_spec.py"), str(source), str(output),
                                     "--handoff", str(handoff), "--report", str(handoff)], capture_output=True, timeout=30)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(handoff.read_text(encoding="utf-8"), "{}")

    def test_atomic_json_write_preserves_previous_report_on_failure(self):
        with tempfile.TemporaryDirectory() as folder:
            report = Path(folder) / "report.json"
            report.write_text("old", encoding="utf-8")
            with patch("json_output.os.replace", side_effect=OSError("publish failed")):
                with self.assertRaises(OSError):
                    write_json(report, {"new": "中文"})
            self.assertEqual(report.read_text(encoding="utf-8"), "old")
            self.assertEqual(list(Path(folder).iterdir()), [report])
            write_json(report, {"new": "中文"})
            self.assertEqual(json.loads(report.read_text(encoding="utf-8")), {"new": "中文"})


class ReferenceSafetyTests(unittest.TestCase):
    def test_references_stop_at_plain_module_titles_and_protect_their_content(self):
        with tempfile.TemporaryDirectory() as folder:
            source, output = Path(folder) / "source.docx", Path(folder) / "output.docx"
            doc = Document()
            for text in ("参考文献", "[1] Reference entry", "致谢", "Thanks", "附录", "Appendix contents"):
                doc.add_paragraph(text)
            protected = [xml_value(p._p) for p in doc.paragraphs[2:]]
            doc.save(source)
            plan = build_plan(source, output, {"operations": [{"action": "reference.numbering.convert",
                "target": {"type": "document"}, "params": {"numbering_mode": "word-numbering"}}]})
            report = apply_plan(plan)
            self.assertEqual(verify_plan(plan, report)["status"], "passed")
            out = Document(output)
            self.assertEqual(out.paragraphs[1].style.name, "Bibliography")
            self.assertEqual(out.paragraphs[1].text, "Reference entry")
            self.assertEqual([xml_value(p._p) for p in out.paragraphs[2:]], protected)
            output.unlink()
            cap = REGISTRY["reference.numbering.convert"]
            def bad(index, ref, params):
                cap.handler(index, ref, params)
                index.document.paragraphs[2].style = "Bibliography"
            with patch.dict(REGISTRY, {cap.action: type(cap)(**{**cap.__dict__, "handler": bad})}):
                with self.assertRaises(FormatError):
                    apply_plan(plan)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
