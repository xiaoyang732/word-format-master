#!/usr/bin/env python3
"""Prepare and hand off a local Word formatting dashboard session."""

from __future__ import annotations

import copy
import hashlib
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx import Document

from analyze_docx import analyze_docx
from citation_workflow import build_citation_task
from config import (
    MAX_DOCUMENT_BYTES,
    SESSION_TTL_SECONDS,
    RENDER_METHODS,
    RENDER_ORDER,
    VISUAL_MODEL_SELECTIONS,
    utc_now as _utc_now,
)
from verify_output import validate_visual_report


def _default_output_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}-formatted.docx")


def session_expired(session: dict[str, Any] | None) -> bool:
    if not session or session.get("status") == "expired":
        return True
    created_at = str(session.get("created_at") or "")
    try:
        created = datetime.fromisoformat(created_at)
    except ValueError:
        return True
    if created.tzinfo is None:
        created = created.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - created).total_seconds() > SESSION_TTL_SECONDS


def expire_session_if_needed(session: dict[str, Any] | None) -> bool:
    if not session_expired(session):
        return False
    if session is not None:
        session["status"] = "expired"
        session.get("input", {}).pop("data", None)
        session["analysis"] = {}
        session["citation_context"] = {"status": "expired", "task": None}
        session["result"] = None
    return True


def _citation_context(source: Path, data: bytes) -> dict[str, Any]:
    if source.suffix.lower() != ".docx":
        return {"status": "unavailable", "reason": "Citation analysis requires DOCX input", "task": None}
    try:
        document = Document(str(source))
    except (KeyError, ValueError):
        return {
            "status": "unavailable",
            "reason": "Strict OOXML publisher templates stay in the official-template workflow",
            "task": None,
        }
    task = build_citation_task(
        document,
        filename=source.name,
        source_sha256=hashlib.sha256(data).hexdigest(),
    )
    return {
        "status": "ready",
        "paragraph_count": len(task["paragraphs"]),
        "reference_count": len(task["references"]),
        "task": task,
    }


def prepare_dashboard_session(input_path: str | Path, output_path: str | Path | None = None) -> dict[str, Any]:
    source = Path(input_path).resolve()
    if not source.is_file() or source.suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("Dashboard session input must be an existing DOCX or DOTX file")
    data = source.read_bytes()
    if not data or len(data) > MAX_DOCUMENT_BYTES:
        raise ValueError("Dashboard session input is empty or exceeds 30 MB")
    output = Path(output_path).resolve() if output_path else _default_output_path(source)
    if output == source:
        raise ValueError("Dashboard session output must differ from the input")
    analysis = analyze_docx(source)
    citations = _citation_context(source, data)
    return {
        "id": uuid.uuid4().hex,
        "mode": "ai-session",
        "status": "configuring",
        "created_at": _utc_now(),
        "input": {
            "path": str(source),
            "name": source.name,
            "size": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
        },
        "output_path": str(output),
        "analysis": analysis,
        "citation_context": citations,
        "verification": {},
        "visual_verification": None,
        "result": None,
    }


def public_dashboard_session(session: dict[str, Any] | None) -> dict[str, Any]:
    if not session:
        return {"mode": "standalone", "status": "ready"}
    expire_session_if_needed(session)
    return {
        "id": session["id"],
        "mode": session["mode"],
        "status": session["status"],
        "input": {
            key: value
            for key, value in session["input"].items()
            if key != "data"
        },
        "output_path": session["output_path"],
        "analysis": copy.deepcopy(session["analysis"]),
        "citation_context": copy.deepcopy(session["citation_context"]),
        "citation_request": bool(session.get("citation_request", False)),
        "verification": copy.deepcopy(session.get("verification") or {}),
        "visual_verification": copy.deepcopy(session.get("visual_verification")),
    }


def public_dashboard_session_status(session: dict[str, Any] | None) -> dict[str, Any]:
    """Return the small, poll-safe portion of a dashboard session."""
    if not session:
        return {"mode": "standalone", "status": "ready"}
    expire_session_if_needed(session)
    return {
        "id": session["id"],
        "mode": session["mode"],
        "status": session["status"],
        "verification": copy.deepcopy(session.get("verification") or {}),
        "visual_verification": copy.deepcopy(session.get("visual_verification")),
    }


def _value_at_path(spec: dict[str, Any], path: str) -> tuple[bool, Any]:
    cursor: Any = spec
    for part in path.split("."):
        if not isinstance(cursor, dict) or part not in cursor:
            return False, None
        cursor = cursor[part]
    return True, cursor


def resolve_application_methods(spec: dict[str, Any], registry: dict[str, str]) -> list[dict[str, Any]]:
    methods: list[dict[str, Any]] = []
    for path, method in registry.items():
        if path == "clear_direct_font_formatting":
            continue
        if "*" in path:
            prefix, field = path.split("*.", 1)
            for heading in spec.get("headings", []):
                if field in heading:
                    methods.append({
                        "path": f"{prefix}{heading.get('level')}.{field}",
                        "value": copy.deepcopy(heading[field]),
                        "method": method,
                    })
            continue
        found, value = _value_at_path(spec, path)
        if found:
            methods.append({"path": path, "value": copy.deepcopy(value), "method": method})
    return methods


def resolve_renderer_plan(verification: dict[str, Any], renderers: dict[str, Any] | None) -> dict[str, Any]:
    requested = str(verification.get("render_method") or "auto")
    if requested not in RENDER_METHODS:
        raise ValueError(f"不支持的视觉渲染方式：{requested}")
    verification["render_method"] = requested
    order = list(RENDER_ORDER) if requested == "auto" else [requested]
    candidates = []
    for method in order:
        entry = (renderers or {}).get(method) or {}
        candidates.append({
            "id": method,
            "name": entry.get("name") or ("Microsoft Word" if method == "word" else "LibreOffice"),
            "available": bool(entry.get("available")) if renderers is not None else None,
            "path": entry.get("path"),
            "method": entry.get("method"),
        })
    selected = next((item for item in candidates if item["available"] is True), None)
    availability_known = renderers is not None
    reason = None
    if availability_known and selected is None:
        reason = (
            "未检测到 Microsoft Word 或 LibreOffice"
            if requested == "auto"
            else f"用户指定的 {candidates[0]['name']} 当前不可用；禁止自动改用其他渲染器"
        )
    return {
        "requested_method": requested,
        "fallback_policy": "word-then-libreoffice" if requested == "auto" else "no-fallback",
        "allowed_order": order,
        "selected_method": selected["id"] if selected else None,
        "available": bool(selected) if availability_known else None,
        "reason": reason,
        "candidates": candidates,
    }


def build_ai_handoff(
    session: dict[str, Any],
    payload: dict[str, Any],
    *,
    skill_dir: Path,
    registry: dict[str, str],
    dashboard_url: str | None = None,
    renderers: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if expire_session_if_needed(session):
        raise ValueError("Dashboard session has expired; reopen the Dashboard")
    if session.get("status") == "submitted":
        raise ValueError("本次设置已经交给 AI，不能重复提交或覆盖交接方案")
    spec = copy.deepcopy(payload.get("spec") or {})
    verification = copy.deepcopy(payload.get("verification") or {})
    citation_requested = bool(payload.get("citation_request", False))
    if not isinstance(verification, dict):
        raise ValueError("验收设置必须是对象")
    visual_model = str(verification.get("visual_model") or "auto")
    if visual_model not in VISUAL_MODEL_SELECTIONS:
        raise ValueError(f"不支持的视觉模型选择：{visual_model}")
    verification["visual_model"] = visual_model
    custom_model_id = str(verification.get("custom_model_id") or "").strip()
    if len(custom_model_id) > 200:
        raise ValueError("自定义视觉模型 ID 不能超过 200 个字符")
    verification["custom_model_id"] = custom_model_id
    if verification.get("visual_enabled") and visual_model == "custom" and not custom_model_id:
        raise ValueError("启用视觉验收并选择自定义模型时，必须填写模型 ID")
    renderer_available = any(
        bool((renderers or {}).get(method, {}).get("available"))
        for method in RENDER_ORDER
    )
    visual_pipeline_available = bool(renderers.get("visual_ready")) if renderers and "visual_ready" in renderers else renderer_available
    if verification.get("visual_enabled") and renderers is not None and not visual_pipeline_available:
        if not renderer_available:
            raise ValueError("未检测到 Microsoft Word 或 LibreOffice，无法启用视觉验收")
        raise ValueError("缺少 PDF 逐页转图组件，无法启用视觉验收")
    clear_direct = bool(payload.get("clear_direct_font_formatting", False))
    official_template = spec.get("mode") == "official-template"
    if spec.get("template_required") and not official_template:
        raise ValueError("template_required is reserved for official-template workflows")
    methods = resolve_application_methods(spec, registry)
    if clear_direct:
        methods.append({
            "path": "clear_direct_font_formatting",
            "value": True,
            "method": registry["clear_direct_font_formatting"],
        })
    input_path = session["input"]["path"]
    output_path = session["output_path"]
    execution_mode = "official-template" if official_template else "parameterized"
    apply_script = str((skill_dir / "scripts" / "apply_spec.py").resolve())
    analyze_script = str((skill_dir / "scripts" / "analyze_docx.py").resolve())
    verify_script = str((skill_dir / "scripts" / "verify_output.py").resolve())
    apply_arguments = [
        input_path,
        output_path,
        "--handoff",
        "<dashboard-handoff.json>",
    ]
    renderer_plan = resolve_renderer_plan(verification, renderers)
    renderer_unavailable = verification.get("visual_enabled") and renderer_plan["available"] is False
    visual_verification = (
        {
            "status": "skipped" if renderer_unavailable else "pending",
            "reason": renderer_plan["reason"] if renderer_unavailable else "等待 AI 渲染所有页面并回传视觉验收结果",
            "model_selection": verification.get("visual_model", "auto"),
            "render_method": renderer_plan["requested_method"],
        }
        if verification.get("visual_enabled")
        else None
    )
    render_script = str((skill_dir / "scripts" / "render_docx.py").resolve())
    visual_tool = {
        "tool": "Word Format Master multi-renderer",
        "path": render_script,
        "purpose": "Render every DOCX page to PNG with the user-selected local renderer",
        "arguments": [
            output_path,
            "--output-dir",
            "<rendered-pages-dir>",
            "--method",
            renderer_plan["requested_method"],
        ],
        "renderer_selection": renderer_plan,
        "on_unavailable": "Report skipped with the detection reason. Never replace an explicit renderer selection.",
        "report_contract": {
            "session_id": session["id"],
            "output_sha256": "SHA-256 of the completed output DOCX",
            "page_count": "positive integer count of rendered pages",
            "report": {
                "model": {"id": "selected model identifier", "supports_image_input": True},
                "renderer": {
                    "method": "word or libreoffice",
                    "manifest_path": "<rendered-pages-dir>/render-manifest.json",
                    "page_files": ["page-1.png"],
                },
                "status": "passed, failed, or skipped",
                "reviewed_pages": [1],
                "findings": [],
            },
        },
    }
    if dashboard_url:
        visual_tool["report_endpoint"] = f"{dashboard_url.rstrip('/')}/api/session/verification"
    template_contract = None
    if official_template:
        local_templates = []
        for item in spec.get("local_templates", []):
            relative = str(item.get("path") or item.get("filename") or "").strip()
            if not relative:
                continue
            template_root = (skill_dir / "assets" / "templates").resolve()
            requested_path = (template_root / relative).resolve()
            if template_root not in requested_path.parents:
                raise ValueError("Official template path must stay inside assets/templates")
            if item.get("path"):
                local_path = requested_path if requested_path.is_file() else None
            else:
                matches = list(template_root.rglob(Path(relative).name)) if template_root.is_dir() else []
                if len(matches) > 1:
                    raise ValueError(f"Official template filename is ambiguous: {relative}")
                local_path = matches[0] if matches else None
            if item.get("required") is True and not local_path:
                raise ValueError(f"Required official template is unavailable: {relative}")
            filename = Path(relative).name
            local_templates.append({
                "label": item.get("label") or item.get("paper_size") or filename,
                "path": str(local_path.resolve()) if local_path else None,
                "filename": filename,
                "paper_size": item.get("paper_size"),
                "sha256": item.get("sha256"),
                "available": bool(local_path and local_path.is_file()),
            })
        template_contract = {
            "workflow_id": spec.get("id"),
            "name": spec.get("name"),
            "official_url": spec.get("official_url"),
            "local_templates": local_templates,
            "features": copy.deepcopy(spec.get("template_features") or []),
            "processing_scope": copy.deepcopy(spec.get("processing_scope") or {}),
            "format_contract": copy.deepcopy(spec.get("format_contract") or {}),
            "instructions": copy.deepcopy(spec.get("workflow") or []),
        }
    handoff = {
        "schema_version": "1.2",
        "kind": "word-format-ai-handoff",
        "status": "ready-for-ai",
        "session_id": session["id"],
        "submitted_at": _utc_now(),
        "source": {
            "path": input_path,
            "name": session["input"]["name"],
            "sha256": session["input"]["sha256"],
        },
        "output_path": output_path,
        "execution_mode": execution_mode,
        "spec": spec,
        "clear_direct_font_formatting": clear_direct,
        "application_methods": methods,
        "citation_context": copy.deepcopy(session["citation_context"]),
        "citation_request": citation_requested,
        "verification": verification,
        "visual_verification": copy.deepcopy(visual_verification),
        "template": template_contract,
        "tools": {
            "apply": {
                "path": apply_script,
                "purpose": "Apply the user-confirmed Dashboard settings and AI-approved citation placements deterministically",
                "arguments": apply_arguments,
                "optional_arguments": [
                    "--spec",
                    "<task-local-citation-spec.json>",
                ],
            },
            "reanalyze": {
                "path": analyze_script,
                "purpose": "Re-open and structurally inspect the generated DOCX",
                "arguments": [output_path, "--output", "<final-analysis.json>"],
            },
            "verify_structure": {
                "path": verify_script,
                "purpose": "Compare generated DOCX structure with the confirmed specification",
                "arguments": [output_path, "<confirmed-spec.json>"],
            },
            "verify_visual": visual_tool,
        },
        "next_steps": (
            [
                "Use tools.template.local_templates as the authoring mother document and preserve its package structure.",
                "Process only tools.template.processing_scope.start through tools.template.processing_scope.end; preserve every excluded page exactly as supplied.",
                "Follow tools.template.instructions and every applicable rule in tools.template.format_contract; do not pass this workflow to apply_spec.py.",
                "Perform structural verification before delivery. When visual verification is enabled, run tools.verify_visual.path with its exact arguments and post the final result to tools.verify_visual.report_endpoint.",
            ]
            if official_template
            else [
                "Analyze citation_context.task in memory and add only validated citations.placements to spec.citations." if citation_requested else "Do not add citation placements unless the user explicitly requested citation analysis.",
                "Run tools.apply with this handoff. Omit its optional --spec argument unless validated citations.placements were added to a task-local copy of handoff.spec.",
                "Run tools.reanalyze and tools.verify_structure.",
                "When visual verification is enabled, run tools.verify_visual.path with its exact arguments. Auto mode may use Word then LibreOffice; an explicit selection has no fallback.",
                "Use the selected model only if image-input support is established; otherwise report skipped.",
                "Post the final visual report to tools.verify_visual.report_endpoint. A passed or failed report must list every rendered page; a skipped report must state why image input was unavailable.",
                "Return the output DOCX and acceptance result to the user.",
            ]
        ),
    }
    session["status"] = "submitted"
    session["verification"] = verification
    session["visual_verification"] = visual_verification
    session["result"] = handoff
    return copy.deepcopy(handoff)


def record_visual_verification(session: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    """Record a post-render visual acceptance result for the active AI session."""
    if expire_session_if_needed(session):
        raise ValueError("Dashboard session has expired; reopen the Dashboard")
    if not session or session.get("status") != "submitted":
        raise ValueError("AI 会话尚未交接，不能回传视觉验收结果")
    if payload.get("session_id") != session.get("id"):
        raise ValueError("视觉验收会话 ID 不匹配")
    if not session.get("verification", {}).get("visual_enabled"):
        raise ValueError("当前会话没有开启视觉验收")
    output = Path(session["output_path"])
    if not output.is_file():
        raise ValueError("尚未找到 AI 完成后的输出 DOCX，不能记录视觉验收")
    output_sha256 = hashlib.sha256(output.read_bytes()).hexdigest()
    if str(payload.get("output_sha256") or "").lower() != output_sha256:
        raise ValueError("视觉验收报告与当前输出 DOCX 的 SHA-256 不匹配")
    report = payload.get("report")
    model = report.get("model") if isinstance(report, dict) and isinstance(report.get("model"), dict) else {}
    verification = session.get("verification", {})
    if verification.get("visual_model") == "custom":
        expected_model_id = str(verification.get("custom_model_id") or "")
        if str(model.get("id") or "") != expected_model_id:
            raise ValueError("视觉验收回传模型与用户选择的自定义模型 ID 不一致")
    reported_status = str(report.get("status") or "") if isinstance(report, dict) else ""
    page_count = payload.get("page_count")
    renderer = report.get("renderer") if isinstance(report, dict) and isinstance(report.get("renderer"), dict) else {}
    if model.get("supports_image_input") is True and reported_status in {"passed", "failed"}:
        if isinstance(page_count, bool) or not isinstance(page_count, int) or page_count < 1:
            raise ValueError("通过或失败的视觉验收必须提供正整数渲染页数")
        renderer_method = str(renderer.get("method") or "")
        requested_method = str(verification.get("render_method") or "auto")
        allowed = RENDER_ORDER if requested_method == "auto" else [requested_method]
        if renderer_method not in allowed:
            raise ValueError("视觉验收使用的渲染器与用户选择不一致")
    manifest_path = Path(str(renderer.get("manifest_path") or "")).resolve()
    if reported_status in {"passed", "failed"} and model.get("supports_image_input") is True:
        if not manifest_path.is_file():
            raise ValueError("视觉验收必须提供本次渲染生成的 manifest 文件")
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError("渲染 manifest 无法读取") from exc
        if Path(str(manifest.get("source") or "")).resolve() != output.resolve():
            raise ValueError("渲染 manifest 的源文件不是当前输出 DOCX")
        if manifest.get("source_sha256") != output_sha256 or manifest.get("page_count") != page_count:
            raise ValueError("渲染 manifest 与当前输出 DOCX 或页数不匹配")
        if manifest.get("selected_method") != renderer.get("method"):
            raise ValueError("渲染 manifest 与验收报告使用的渲染器不一致")
        pages = manifest.get("pages")
        if not isinstance(pages, list) or len(pages) != page_count:
            raise ValueError("渲染 manifest 的页面清单不完整")
        for index, item in enumerate(pages, start=1):
            if item.get("number") != index:
                raise ValueError("渲染 manifest 页码不连续")
            page_path = Path(str(item.get("path") or "")).resolve()
            if not page_path.is_file() or hashlib.sha256(page_path.read_bytes()).hexdigest() != item.get("sha256"):
                raise ValueError("渲染 manifest 中存在缺失或被修改的页面文件")
        reported_pages = renderer.get("page_files")
        manifest_page_paths = [str(Path(str(item["path"])).resolve()) for item in pages]
        if not isinstance(reported_pages, list) or [str(Path(str(path)).resolve()) for path in reported_pages] != manifest_page_paths:
            raise ValueError("视觉验收报告的页面清单与渲染 manifest 不一致")
    result = validate_visual_report(report, page_count)
    result.update({
        "output_path": str(output),
        "output_sha256": output_sha256,
        "recorded_at": _utc_now(),
    })
    session["visual_verification"] = result
    return copy.deepcopy(result)
