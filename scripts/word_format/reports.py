"""Select an evidence document and prepare a report draft; never invent formats."""

from __future__ import annotations

import json
import base64
import hashlib
import os
import re
import tempfile
from pathlib import Path

from json_output import write_json
from .contracts import FormatError, sha256
from .profiles import template_profile

CATALOG = Path(__file__).resolve().parents[2] / "templates" / "reports.json"


def _distribution_template(path, expected_hash):
    """Restore a verified template when a distribution forbids DOCX uploads."""
    encoded = path.with_suffix(path.suffix + ".json")
    if not encoded.is_file():
        return path
    if not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
        raise FormatError("Invalid encoded report template hash")
    payload = json.loads(encoded.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("format") != "docx-base64-v1":
        raise FormatError("Unsupported encoded report template")
    try:
        data = base64.b64decode(payload["data"], validate=True)
    except (ValueError, KeyError, TypeError) as exc:
        raise FormatError("Invalid encoded report template data") from exc
    if hashlib.sha256(data).hexdigest() != expected_hash:
        raise FormatError("Encoded report template hash mismatch", code="TEMPLATE_HASH_MISMATCH",
                          user_message="编码模板摘要与登记不一致", next_actions=["重新安装完整发布包，勿绕过摘要校验"])
    cache = Path(tempfile.gettempdir()) / "word-format-master" / "templates"
    cache.mkdir(parents=True, exist_ok=True)
    restored = cache / (expected_hash + ".docx")
    if not restored.is_file() or sha256(restored) != expected_hash:
        with tempfile.NamedTemporaryFile(dir=cache, delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(data)
        try:
            os.replace(temporary, restored)
        finally:
            temporary.unlink(missing_ok=True)
    return restored


def load_report_catalog(path=CATALOG):
    """Validate portable project-relative paths and the approved template digest."""
    path = Path(path).resolve()
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict) or data.get("schema_version") != "1.0":
        raise FormatError("Unsupported report catalog schema")
    entries = data.get("templates")
    if not isinstance(entries, list) or not entries:
        raise FormatError("Report catalog needs templates")
    ids = set()
    resolved = []
    for entry in entries:
        if not isinstance(entry, dict):
            raise FormatError("Report template must be an object")
        identifier = entry.get("id", "")
        if not isinstance(identifier, str) or not re.fullmatch(r"[a-z][a-z0-9-]*", identifier):
            raise FormatError("Invalid report template ID")
        if identifier in ids:
            raise FormatError(f"Duplicate report template: {identifier}")
        ids.add(identifier)
        value = dict(entry)
        for key in ("path", "source_notes", "usage_notes"):
            relative = entry.get(key)
            if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
                raise FormatError(f"Report {key} must be project-relative")
            candidate = (path.parent / relative).resolve()
            if path.parent not in candidate.parents:
                raise FormatError(f"Invalid or missing report {key}: {relative}")
            if key == "path" and not candidate.is_file():
                candidate = _distribution_template(candidate, entry.get("sha256"))
            if not candidate.is_file():
                code = {"path": "REPORT_TEMPLATE_MISSING", "source_notes": "REPORT_SOURCE_NOTES_MISSING",
                        "usage_notes": "REPORT_USAGE_NOTES_MISSING"}[key]
                raise FormatError(f"Invalid or missing report {key}: {relative}", code=code,
                                  user_message=f"默认报告所需资源缺失：{relative}",
                                  next_actions=["重新安装完整发布包，或通过 --sample 指定自己的 DOCX 模板"])
            value[key] = str(candidate)
        if Path(value["path"]).suffix.lower() != ".docx":
            raise FormatError("Registered report templates must be DOCX")
        if sha256(value["path"]) != entry.get("sha256"):
            raise FormatError(f"Report template hash mismatch: {identifier}", code="TEMPLATE_HASH_MISMATCH",
                              user_message="报告模板摘要不匹配", next_actions=["核对模板来源并重新安装完整发布包"])
        resolved.append(value)
    if data.get("default") not in ids:
        raise FormatError("Default report template is not registered")
    return {"kind": "word-report-templates", "default": data["default"], "templates": resolved}


def prepare_report(workspace, *, template_id=None, sample=None, topic=""):
    """Clone all package parts and extract evidence into a new task directory.

    Content authoring remains the AI's job. This copy is explicitly a draft;
    publication still requires the shared plan/apply/verify and page review.
    """
    if template_id and sample:
        raise FormatError("Choose either a registered report template or a supplied sample")
    destination = Path(workspace).resolve()
    template_root = CATALOG.parent.resolve()
    if destination == template_root or template_root in destination.parents:
        raise FormatError("Report workspace must not be inside the template library")
    if destination.exists():
        raise FormatError("Report workspace already exists; choose a new directory")
    entry = None
    if sample:
        source = Path(sample).resolve()
        if not source.is_file() or source.suffix.lower() != ".docx":
            raise FormatError("Supplied report sample must be an existing DOCX; convert other formats first")
        template_id = None
    else:
        catalog = load_report_catalog()
        template_id = template_id or catalog["default"]
        entry = next((t for t in catalog["templates"] if t["id"] == template_id), None)
        if entry is None:
            raise FormatError(f"Report template is not registered: {template_id}")
        source = Path(entry["path"])
    source_bytes = source.read_bytes()
    digest = sha256(source)
    if entry and digest != entry["sha256"]:
        raise FormatError("Registered report template changed before preparation")
    profile = template_profile(source)
    if sha256(source) != digest or profile["source"]["sha256"] != digest:
        raise FormatError("Report template changed during preparation")
    manifest = {
        "schema_version": "1.0", "kind": "word-report-workflow", "status": "draft",
        "template_id": template_id, "topic": topic,
        "source": {"path": str(source), "sha256": digest},
        "draft": str(destination / "draft.docx"),
        "profile": str(destination / "profile.json"),
        "required_plan_options": ["--match-structure"],
        "structure_contract": profile["snapshot"]["structure_contract"],
        "instructions": str(CATALOG.parents[1] / "references" / "report-generation.md"),
        "source_notes": entry["source_notes"] if entry else None,
        "usage_notes": entry["usage_notes"] if entry else None,
        "next_steps": [
            "Read instructions and profile issues; replace draft placeholders with the requested report content.",
            "Resolve evidence issues and map targets; preserve cloned cover tables, fields and complex objects explicitly.",
            "Run check-structure, then run_format_pipeline.py with --profile, --match-structure and applicable --mapping/--resolutions.",
            "Refresh fields, check citations/content, render all pages and bind the page review to final verify.",
        ],
    }
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".report-", dir=destination.parent) as temporary:
        staging = Path(temporary) / "task"
        staging.mkdir()
        (staging / "draft.docx").write_bytes(source_bytes)
        if sha256(staging / "draft.docx") != digest:
            raise FormatError("Report template changed before cloning")
        write_json(staging / "profile.json", profile, source)
        write_json(staging / "workflow.json", manifest, source)
        if destination.exists():
            raise FormatError("Report workspace was created by another process")
        staging.rename(destination)
    return manifest
