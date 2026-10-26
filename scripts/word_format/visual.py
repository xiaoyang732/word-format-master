"""Bind visual acceptance to the current DOCX and every rendered page."""

from __future__ import annotations

import json
from pathlib import Path

from .contracts import FormatError, number, require_keys, sha256, text_value


def verify_visual_evidence(output, report, manifest_path=None, requested_method="auto"):
    from verify_output import validate_visual_report

    if requested_method not in {"auto", "word", "libreoffice"}:
        raise FormatError("Unknown visual renderer")
    if not isinstance(report, dict):
        raise FormatError("Visual report must be an object")
    basic = validate_visual_report(report)
    if basic["status"] == "skipped":
        return basic
    if report.get("status") not in {"passed", "failed"}:
        raise FormatError("Visual report must be passed, failed or skipped")
    text_value(report.get("model", {}).get("id"), "visual model ID")
    if manifest_path is None:
        raise FormatError("Visual acceptance requires a render manifest")
    path = Path(manifest_path).resolve()
    if not path.is_file():
        raise FormatError("Render manifest is missing")
    original_manifest_hash = sha256(path)
    if report.get("manifest_sha256") != original_manifest_hash:
        raise FormatError("Visual report does not bind the current render manifest")
    manifest = json.loads(path.read_text(encoding="utf-8-sig"))
    require_keys(
        manifest,
        {"schema_version", "source", "source_sha256", "selected_method", "page_count", "pages"},
        {"schema_version", "source", "source_sha256", "selected_method", "page_count", "pages"},
        "render manifest",
    )
    if manifest["schema_version"] != "1.0":
        raise FormatError("Unsupported render manifest version")
    output = Path(output).resolve()
    output_hash = sha256(output)
    source = Path(text_value(manifest["source"], "manifest source"))
    if not source.is_absolute() or source.resolve() != output:
        raise FormatError("Render manifest source is not the current output")
    if manifest["source_sha256"] != output_hash or report.get("output_sha256") != output_hash:
        raise FormatError("Visual evidence does not match the current output SHA-256")
    selected = manifest["selected_method"]
    if selected not in {"word", "libreoffice"} or (
        requested_method != "auto" and requested_method != selected
    ):
        raise FormatError("Visual evidence used a different renderer than requested")
    renderer = report.get("renderer")
    if not isinstance(renderer, dict) or renderer.get("method") != selected:
        raise FormatError("Visual report renderer does not match the manifest")
    count = manifest["page_count"]
    if isinstance(count, bool) or not isinstance(count, int):
        raise FormatError("Rendered page count must be an integer")
    number(count, 1, 100000, "rendered page count", True)
    pages = manifest["pages"]
    if not isinstance(pages, list) or len(pages) != count:
        raise FormatError("Render manifest must list every page")
    page_paths = set()
    for i, page in enumerate(pages, 1):
        require_keys(
            page, {"number", "path", "sha256"}, {"number", "path", "sha256"}, "rendered page"
        )
        if (
            isinstance(page["number"], bool)
            or not isinstance(page["number"], int)
            or page["number"] != i
        ):
            raise FormatError("Render manifest page numbers must be sequential integers")
        image_path = Path(text_value(page["path"], "rendered page path"))
        if not image_path.is_absolute():
            raise FormatError("Rendered page paths must be absolute")
        image_path = image_path.resolve()
        if (
            image_path in page_paths
            or not image_path.is_file()
            or sha256(image_path) != page["sha256"]
        ):
            raise FormatError("Rendered page is duplicated, missing or changed")
        page_paths.add(image_path)
    result = validate_visual_report(report, count)
    if sha256(output) != output_hash or sha256(path) != original_manifest_hash:
        raise FormatError("Visual evidence changed during verification")
    return {
        **result,
        "output_sha256": output_hash,
        "manifest_sha256": original_manifest_hash,
        "manifest_path": str(path),
        "renderer": {"method": selected},
        "page_count": count,
    }
