#!/usr/bin/env python3
"""Validate direct-AI assets, presets, operation contracts and removed entry points."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from apply_spec import validate_spec
from word_format.registry import REGISTRY
from word_format.templates import list_templates, load_template
REQUIRED = [
    "SKILL.md", "README.md", "agents/openai.yaml", "assets/presets.json",
    "references/format-profile.md", "references/direct-format-api.md",
    "scripts/format_cli.py", "scripts/word_format/profiles.py",
    "scripts/word_format/requirement_sources.py", "scripts/word_format/requirement_rules.py",
    "scripts/render_docx.py", "scripts/verify_output.py",
    "scripts/word_format/visual.py",
    "scripts/word_format/template_structure.py",
    "references/thesis-generation.md",
]
REMOVED = [
    "assets/ui/index.html", "assets/ui/app.js", "assets/ui/styles.css",
    "scripts/serve_dashboard.py", "scripts/dashboard_session.py",
    "scripts/wait_for_dashboard.py", "scripts/local_ai_analysis.py",
    "package.json", "package-lock.json", "tests/test_dashboard_ui.cjs",
    "references/dashboard-workflow.md",
]


def main():
    errors = []
    for template in list_templates():
        try:
            load_template(template["id"])
        except (ValueError, OSError) as exc:
            errors.append(f"Invalid builtin template {template['id']}: {exc}")
    for name in REQUIRED:
        if not (ROOT / name).is_file():
            errors.append(f"Missing required file: {name}")
    for name in REMOVED:
        if (ROOT / name).exists():
            errors.append(f"Removed web entry point remains: {name}")
    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    if not re.match(r"^---\nname: word-format-master\ndescription: [^\n]+(?:\n[a-zA-Z0-9_-]+: [^\n]+)*\n---\n", skill):
        errors.append("Invalid skill frontmatter")
    for cap in REGISTRY.values():
        for attribute in ("handler", "validator", "reader", "verifier"):
            if not callable(getattr(cap, attribute)):
                errors.append(f"{cap.action}: missing {attribute}")
    presets_path = ROOT / "assets" / "presets.json"
    if presets_path.is_file():
        try:
            data = json.loads(presets_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            errors.append(f"presets.json is invalid: {exc}")
            data = {"presets": []}
        ids: set[str] = set()
        for preset in data.get("presets", []):
            preset_id = preset.get("id")
            if not preset_id:
                errors.append("preset without id")
                continue
            if preset_id in ids:
                errors.append(f"duplicate preset id: {preset_id}")
            ids.add(preset_id)
            if preset.get("template_required"):
                errors.append(f"{preset_id}: numeric presets cannot require a template")
            validation = validate_spec(preset)
            errors.extend(f"{preset_id}: {message}" for message in validation)
            body = preset.get("body", {})
            if not body.get("font_latin") or not body.get("font_east_asia") or not body.get("font_size_pt"):
                errors.append(f"{preset_id}: body font role is incomplete")
            heading_levels = {int(item.get("level", 0)) for item in preset.get("headings", [])}
            if not {1, 2, 3}.issubset(heading_levels):
                errors.append(f"{preset_id}: H1-H3 presets are incomplete")
            toc = preset.get("table_of_contents", {})
            if not {"enabled", "title", "max_heading_level", "page_break_after"}.issubset(toc):
                errors.append(f"{preset_id}: table-of-contents preset is incomplete")
        expected = {
            "thesis-standard",
            "journal-paper-cn",
            "official-report-cn",
        }
        missing = expected - ids
        if missing:
            errors.append(f"missing expected presets: {', '.join(sorted(missing))}")

        workflow_ids: set[str] = set()
        for workflow in data.get("template_workflows", []):
            workflow_id = workflow.get("id")
            if not workflow_id:
                errors.append("template workflow without id")
                continue
            if workflow_id in ids or workflow_id in workflow_ids:
                errors.append(f"duplicate preset or workflow id: {workflow_id}")
            workflow_ids.add(workflow_id)
            if workflow.get("mode") != "official-template":
                errors.append(f"{workflow_id}: mode must be official-template")
            if workflow.get("template_required") is not True:
                errors.append(f"{workflow_id}: template_required must be true")
            if not str(workflow.get("official_url", "")).startswith("https://"):
                errors.append(f"{workflow_id}: official_url must use HTTPS")
            if workflow.get("page") or workflow.get("body") or workflow.get("headings"):
                errors.append(f"{workflow_id}: official workflows cannot contain synthetic format tokens")
            if len(workflow.get("workflow", [])) < 2:
                errors.append(f"{workflow_id}: workflow steps are missing")
            for template in workflow.get("local_templates", []):
                relative = str(template.get("path") or template.get("filename", "")).strip()
                filename = Path(relative).name
                template_root = ROOT / "assets" / "templates"
                candidate = (template_root / relative).resolve() if relative else None
                if candidate is not None and template_root.resolve() not in candidate.parents:
                    errors.append(f"{workflow_id}: local template path escapes assets/templates")
                    continue
                matches = [candidate] if candidate and candidate.is_file() else []
                if not template.get("path") and filename:
                    matches = list(template_root.rglob(filename)) if template_root.is_dir() else []
                if len(matches) > 1:
                    errors.append(f"{workflow_id}: ambiguous local template filename {filename}; use path")
                if template.get("required") is True and not matches:
                    errors.append(f"{workflow_id}: missing required local template {filename}")
                if matches and template.get("sha256"):
                    digest = hashlib.sha256(matches[0].read_bytes()).hexdigest()
                    if digest.lower() != str(template["sha256"]).lower():
                        errors.append(f"{workflow_id}: local template hash mismatch {filename}")
        expected_workflows = set()
        missing_workflows = expected_workflows - workflow_ids
        if missing_workflows:
            errors.append(f"missing expected workflows: {', '.join(sorted(missing_workflows))}")


    for name in ("README.md", "SKILL.md", "SECURITY.md", "agents/openai.yaml",
                 "references/architecture.md", "references/format-profile.md",
                 "references/direct-format-api.md", "references/format-spec.md",
                 "references/header-requirements.md", "references/thesis-generation.md"):
        text = (ROOT / name).read_text(encoding="utf-8").lower()
        if any(token in text for token in ("dashboard", "--handoff", "npm test", "playwright")):
            errors.append(f"Obsolete web instructions in {name}")
    if errors:
        for error in errors:
            print(error)
        return 1
    print(f"Project validation passed: {len(REGISTRY)} registered operations, direct AI profiles and presets")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
