#!/usr/bin/env python3
"""Validate the Word Format Master skill assets and data contracts."""

from __future__ import annotations

import json
import hashlib
import re
import sys
from pathlib import Path

from apply_spec import SUPPORTED_APPLICATION_PATHS, validate_spec
from analyze_docx import analyze_docx
from parse_requirements import parse_requirements


ROOT = Path(__file__).resolve().parent.parent
REQUIRED = [
    "SKILL.md",
    "agents/openai.yaml",
    "assets/presets.json",
    "assets/ui/index.html",
    "assets/ui/styles.css",
    "assets/ui/app.js",
    "references/format-spec.md",
    "references/architecture.md",
    "references/project-research.md",
    "references/standards-registry.md",
    "references/direct-format-api.md",
    "references/header-requirements.md",
    "scripts/format_cli.py",
    "scripts/word_format/__init__.py",
    "scripts/word_format/api.py",
    "scripts/word_format/contracts.py",
    "scripts/word_format/headers.py",
    "scripts/word_format/section_breaks.py",
    "scripts/word_format/inspect.py",
    "scripts/word_format/operations.py",
    "scripts/word_format/properties.py",
    "scripts/word_format/registry.py",
    "scripts/word_format/selectors.py",
    "scripts/word_format/spec.py",
    "scripts/word_format/verification.py",
    "tests/test_format_core.py",
    "tests/test_headers.py",
    "tests/test_dashboard_ui.cjs",
    "scripts/analyze_docx.py",
    "scripts/analyze_citations.py",
    "scripts/apply_spec.py",
    "scripts/citation_workflow.py",
    "scripts/parse_requirements.py",
    "scripts/serve_dashboard.py",
    "scripts/smoke_test.py",
    "scripts/verify_output.py",
    "scripts/dashboard_session.py",
    "scripts/wait_for_dashboard.py",
    "scripts/runtime_detection.py",
    "scripts/render_docx.py",
    "scripts/libreoffice_render.py",
    "scripts/config.py",
    "scripts/document_structure.py",
    "scripts/local_ai_analysis.py",
]


def main() -> int:
    errors: list[str] = []
    for relative in REQUIRED:
        if not (ROOT / relative).is_file():
            errors.append(f"missing required file: {relative}")

    skill_path = ROOT / "SKILL.md"
    if skill_path.is_file():
        skill = skill_path.read_text(encoding="utf-8")
        match = re.match(r"^---\n(.*?)\n---\n", skill, re.S)
        if not match:
            errors.append("SKILL.md frontmatter is missing")
        else:
            keys = [
                line.split(":", 1)[0].strip()
                for line in match.group(1).splitlines()
                if ":" in line
            ]
            if keys != ["name", "description"]:
                errors.append("SKILL.md frontmatter must contain only name and description")
        if "TODO" in skill:
            errors.append("SKILL.md contains TODO placeholders")
        required_interaction_rules = (
            "call `format_cli.py` without launching a browser",
            "AI must not click confirmation",
            "same planner, registry and executor",
            "Direct requests do not require a handoff",
        )
        for rule in required_interaction_rules:
            if rule not in skill:
                errors.append(f"SKILL.md is missing two-entrance contract: {rule}")

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

    ui_files = [
        ROOT / "assets" / "ui" / "index.html",
        ROOT / "assets" / "ui" / "styles.css",
        ROOT / "assets" / "ui" / "app.js",
    ]
    for path in ui_files:
        if path.is_file() and not path.read_text(encoding="utf-8").strip():
            errors.append(f"empty UI asset: {path.name}")

    html_path, css_path, js_path = ui_files
    if html_path.is_file() and js_path.is_file():
        html = html_path.read_text(encoding="utf-8")
        javascript = js_path.read_text(encoding="utf-8")
        html_ids = set(re.findall(r'id="([A-Za-z][A-Za-z0-9_-]*)"', html))
        javascript_ids = set(re.findall(r'\$\("#([A-Za-z][A-Za-z0-9_-]*)"\)', javascript))
        for element_id in sorted(javascript_ids - html_ids):
            errors.append(f"app.js references missing HTML id: {element_id}")
        required_heading_ids = {
            "headingLevelTabs",
            "headingDefinitionBadge",
            "headingEastAsiaInput",
            "headingLatinInput",
            "headingSizeInput",
            "headingAlignInput",
            "headingSpaceBeforeInput",
            "headingSpaceAfterInput",
            "headingBoldInput",
            "headingItalicInput",
        }
        for element_id in sorted(required_heading_ids - html_ids):
            errors.append(f"missing heading editor control: {element_id}")
        heading_tabs = {int(value) for value in re.findall(r'data-heading-level="([1-3])"', html)}
        preview_headings = {int(value) for value in re.findall(r'data-preview-heading="([1-3])"', html)}
        if heading_tabs != {1, 2, 3}:
            errors.append("heading editor must expose only H1-H3 tabs")
        if preview_headings != {1, 2, 3}:
            errors.append("paper preview must render only H1-H3 samples")
        required_workflow_ids = {
            "citationParagraphCount",
            "citationReferenceCount",
            "citationReadyBadge",
            "citationWorkflowStatus",
            "citationRequestButton",
            "visualReviewToggle",
            "visualModelSelect",
            "customVisualModelInput",
            "visualReviewStatus",
            "renderRuntimePanel",
            "renderRuntimeBadge",
            "renderRuntimeStatus",
            "renderMethodSelect",
            "wordRuntimeBadge",
            "libreOfficeRuntimeBadge",
            "renderRuntimeProgress",
            "renderRuntimeProgressText",
            "renderRuntimeProgressBytes",
            "renderRuntimeActions",
            "libreOfficeOfficialLink",
            "aiDownloadRuntimeButton",
            "templateScopeNote",
        }
        for element_id in sorted(required_workflow_ids - html_ids):
            errors.append(f"missing citation or verification control: {element_id}")
        if not re.search(r'id="aiDownloadRuntimeButton"\s+class="button primary"', html):
            errors.append("Agent-managed LibreOffice install must use a high-contrast primary button")
        if not re.search(r'id="visualReviewToggle"\s+type="checkbox"\s+disabled', html):
            errors.append("visual verification toggle must be disabled until renderer detection completes")
        if 'data-settings-tab="toc"' not in html or 'data-setting-panel="toc"' not in html:
            errors.append("missing table-of-contents dashboard panel")
        list_styles = set(re.findall(r'<option value="([a-z-]+)">[^<]*(?:1|一|A|a|I)', html))
        expected_list_styles = {
            "decimal-period",
            "decimal-parenthesis",
            "decimal-parentheses",
            "decimal-fullwidth-parentheses",
            "chinese-period",
            "chinese-parentheses",
            "upper-alpha-period",
            "lower-alpha-parenthesis",
            "upper-roman-period",
        }
        if not expected_list_styles.issubset(list_styles):
            errors.append("numbered-list editor must expose every supported marker style")
        if "function headingForLevel(level, create = false)" not in javascript:
            errors.append("heading editor must defer creating inherited levels")
        if "function listMarker(style, value)" not in javascript:
            errors.append("list marker preview formatter is missing")
        if (
            'fetch("/api/session"' not in javascript
            or 'fetch("/api/session/submit"' not in javascript
            or 'fetch("/api/session/status"' not in javascript
        ):
            errors.append("AI dashboard session workflow is missing")
        if "function submitToAI()" not in javascript:
            errors.append("AI dashboard submit handler is missing")
        if "function visualReviewLogClass(visual)" not in javascript or 'pending: "log-pending"' not in javascript:
            errors.append("visual verification status styling must distinguish pending from passed")
        if (
            'fetch("/api/runtime/status"' not in javascript
            or 'fetch("/api/runtime/download"' not in javascript
            or "function downloadRuntime()" not in javascript
            or "function selectedRuntime()" not in javascript
            or 'render_method: "auto"' not in javascript
        ):
            errors.append("render runtime selection, detection, or background download workflow is missing")
        if (
            "function visualReviewRuntimeReady()" not in javascript
            or "function visualReviewUnavailableMessage()" not in javascript
            or "visualToggle.disabled = !runtimeReady || state.sessionSubmitted" not in javascript
            or "if (!visualReviewRuntimeReady() || state.sessionSubmitted)" not in javascript
        ):
            errors.append("visual verification toggle is not gated by renderer availability and submission state")
        dashboard_source = (ROOT / "scripts" / "serve_dashboard.py").read_text(encoding="utf-8")
        for obsolete in ("exportCitationTaskButton", "citationResultInput", "visualReportInput", "/api/citations/export"):
            if obsolete in html or obsolete in javascript or obsolete in dashboard_source:
                errors.append(f"obsolete user-facing citation or visual import control remains: {obsolete}")
        ui_paths = set(re.findall(r'data-spec-path="([^"]+)"', html))
        ui_paths.update({
            "body.line_spacing",
            "references.line_spacing",
            "headings.*.font_east_asia",
            "headings.*.font_latin",
            "headings.*.font_size_pt",
            "headings.*.alignment",
            "headings.*.space_before_pt",
            "headings.*.space_after_pt",
            "headings.*.bold",
            "headings.*.italic",
            "clear_direct_font_formatting",
        })
        missing_application_paths = ui_paths - set(SUPPORTED_APPLICATION_PATHS)
        for path in sorted(missing_application_paths):
            errors.append(f"UI setting has no DOCX application method: {path}")
        if 'data-spec-path="references.citation_system"' in html:
            errors.append("references.citation_system must remain a read-only identifier, not an editable setting")
    if css_path.is_file():
        css = css_path.read_text(encoding="utf-8")
        for selector in (
            ".heading-level-tabs",
            ".preview-main-body.two-column",
            ".preview-reference",
            ".log-pending",
            ".log-neutral",
            "#customVisualModelControl[hidden]",
            ".runtime-panel",
            ".runtime-progress",
            ".runtime-actions .button.primary:disabled",
        ):
            if selector not in css:
                errors.append(f"missing UI style: {selector}")

    session_path = ROOT / "scripts" / "dashboard_session.py"
    apply_path = ROOT / "scripts" / "apply_spec.py"
    renderer_path = ROOT / "scripts" / "render_docx.py"
    dashboard_path = ROOT / "scripts" / "serve_dashboard.py"
    if session_path.is_file() and "def record_visual_verification(" not in session_path.read_text(encoding="utf-8"):
        errors.append("AI session must record post-render visual verification")
    if apply_path.is_file():
        apply_source = apply_path.read_text(encoding="utf-8")
        for contract in (
            'parser.add_argument("--handoff", required=True',
            "def load_confirmed_handoff(",
            "Input DOCX changed after Dashboard confirmation; reopen the Dashboard",
            "Task-local specification may only add validated citations",
        ):
            if contract not in apply_source:
                errors.append(f"apply_spec.py is missing confirmed-handoff guard: {contract}")
        if 'selection.add_argument("--preset-id"' in apply_source or 'selection.add_argument("--spec"' in apply_source:
            errors.append("apply_spec.py must not expose direct preset or standalone-spec application")
        if "SUPPORTED_APPLICATION_PATHS = legacy_paths()" not in apply_source:
            errors.append("apply_spec.py must register the document module structure application method")
    if session_path.is_file():
        session_source = session_path.read_text(encoding="utf-8")
        for contract in ("resolve_renderer_plan", '"no-fallback"', '"scripts" / "render_docx.py"', "VISUAL_MODEL_SELECTIONS"):
            if contract not in session_source:
                errors.append(f"AI handoff is missing renderer contract: {contract}")
        if "无法启用视觉验收" not in session_source or 'session.get("status") == "submitted"' not in session_source:
            errors.append("AI handoff must reject missing renderers and repeated submissions")
        if '"--handoff",' not in session_source or '"<dashboard-handoff.json>"' not in session_source:
            errors.append("AI handoff must require its confirmed Dashboard handoff when applying formatting")
    if renderer_path.is_file():
        renderer_source = renderer_path.read_text(encoding="utf-8")
        for implementation in ("_word_to_pdf", "_libreoffice_to_pdf", "_pdf_to_pngs", "render-manifest.json", "resolve_renderer", "detect_pdf_rasterizer", "pypdfium2", "libreoffice_render.py"):
            if implementation not in renderer_source:
                errors.append(f"renderer implementation is missing: {implementation}")
    if dashboard_path.is_file():
        dashboard = dashboard_path.read_text(encoding="utf-8")
        for endpoint in ('"/api/session/status"', '"/api/session/verification"', '"/api/runtime/status"', '"/api/runtime/download"'):
            if endpoint not in dashboard:
                errors.append(f"dashboard missing visual verification endpoint: {endpoint}")
        if 'path == "/api/apply"' in dashboard or "def _apply_upload" in dashboard:
            errors.append("dashboard must not expose direct formatting application outside a confirmed session")
    if session_path.is_file() and '"data": base64.b64encode(data)' in session_path.read_text(encoding="utf-8"):
        errors.append("public Dashboard sessions must not serialize source DOCX data")

    sample = parse_requirements(
        "A4 纵向。正文宋体小四，1.5 倍行距，首行缩进 2 字符，两端对齐。"
        "上边距 2.5 cm，下边距 2.5 cm，左边距 3 cm，右边距 2.5 cm。"
        "一级标题黑体三号加粗居中。图题五号居中放在图下并自动编号。"
        "表题五号居中放在表上并自动编号。页码放在页脚居中，从 1 开始，首页不显示页码。"
        "参考文献采用 GB/T 7714-2025，五号，单倍行距，悬挂缩进 7.4 mm，自动编号。"
        "目录包含到三级标题，目录后分页。"
    )
    patch = sample["spec_patch"]
    toc_only_patch = parse_requirements("目录包含到三级标题，目录后分页。 ")["spec_patch"]
    checks = {
        "page.size": patch.get("page", {}).get("size") == "A4",
        "page.margin_left_mm": patch.get("page", {}).get("margin_left_mm") == 30,
        "body.font_east_asia": patch.get("body", {}).get("font_east_asia") == "宋体",
        "body.font_size_pt": patch.get("body", {}).get("font_size_pt") == 12,
        "body.line_spacing": patch.get("body", {}).get("line_spacing", {}).get("value") == 1.5,
        "references.citation_system": patch.get("references", {}).get("citation_system") == "GB/T 7714-2025",
        "captions.figure.position": patch.get("captions", {}).get("figure", {}).get("position") == "below",
        "captions.table.position": patch.get("captions", {}).get("table", {}).get("position") == "above",
        "page_numbers.location": patch.get("page_numbers", {}).get("location") == "footer",
        "page_numbers.show_on_first_page": patch.get("page_numbers", {}).get("show_on_first_page") is False,
        "table_of_contents.enabled": patch.get("table_of_contents", {}).get("enabled") is True,
        "table_of_contents.max_heading_level": patch.get("table_of_contents", {}).get("max_heading_level") == 3,
        "table_of_contents.no_heading_side_effect": not toc_only_patch.get("headings"),
        "references.hanging_indent_mm": patch.get("references", {}).get("hanging_indent_mm") == 7.4,
    }
    errors.extend(f"parser smoke check failed: {name}" for name, passed in checks.items() if not passed)

    if errors:
        print("Project validation failed:")
        for error in errors:
            print(f"- {error}")
        return 1
    print(
        f"Project validation passed: {len(REQUIRED)} required files, "
        f"{len(data.get('presets', []))} presets, "
        f"{len(data.get('template_workflows', []))} template workflows"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
