"""Select an evidence document and prepare a report draft; never invent formats."""

from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path

from json_output import write_json
from .contracts import FormatError, sha256
from .profiles import template_profile

CATALOG = Path(__file__).resolve().parents[2] / "templates" / "reports.json"


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
            if path.parent not in candidate.parents or not candidate.is_file():
                raise FormatError(f"Invalid or missing report {key}: {relative}")
            value[key] = str(candidate)
        if Path(value["path"]).suffix.lower() != ".docx":
            raise FormatError("Registered report templates must be DOCX")
        if sha256(value["path"]) != entry.get("sha256"):
            raise FormatError(f"Report template hash mismatch: {identifier}")
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
