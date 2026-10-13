#!/usr/bin/env python3
"""Validate confirmed Dashboard handoffs and execute the shared format core."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import sys
from pathlib import Path
from typing import Any

from word_format.spec import legacy_paths
# Compatibility exports for callers of the former monolithic implementation.
from word_format.structures import append_field as append_field, apply_page_numbers as apply_page_numbers

SUPPORTED_APPLICATION_PATHS = legacy_paths()

def load_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _resolved_path(value: str | Path, label: str) -> Path:
    try:
        return Path(value).resolve(strict=False)
    except (OSError, TypeError) as exc:
        raise ValueError(f"{label} path is invalid") from exc


def _without_citations(spec: dict[str, Any]) -> dict[str, Any]:
    """Keep the user-confirmed layout immutable while AI adds citation placements."""
    result = copy.deepcopy(spec)
    result.pop("citations", None)
    return result


def load_confirmed_handoff(
    path: str | Path,
    source: str | Path,
    destination: str | Path,
) -> dict[str, Any]:
    """Load a dashboard handoff and bind it to the exact document operation."""
    handoff = load_json(path)
    if not isinstance(handoff, dict):
        raise ValueError("Dashboard handoff must be a JSON object")
    if handoff.get("kind") != "word-format-ai-handoff" or handoff.get("status") != "ready-for-ai":
        raise ValueError("Formatting requires a confirmed Dashboard handoff")
    if handoff.get("execution_mode") != "parameterized":
        raise ValueError("Official-template handoffs must follow their publisher workflow")
    source_info = handoff.get("source")
    if not isinstance(source_info, dict) or not str(source_info.get("path") or ""):
        raise ValueError("Dashboard handoff is missing its source document")
    actual_source = _resolved_path(source, "Input")
    handoff_source = _resolved_path(str(source_info["path"]), "Handoff source")
    if actual_source != handoff_source:
        raise ValueError("Input DOCX does not match the confirmed Dashboard handoff")
    actual_destination = _resolved_path(destination, "Output")
    handoff_destination = _resolved_path(str(handoff.get("output_path") or ""), "Handoff output")
    if actual_destination != handoff_destination:
        raise ValueError("Output DOCX does not match the confirmed Dashboard handoff")
    if not actual_source.is_file():
        raise ValueError("Input DOCX no longer exists")
    actual_sha256 = hashlib.sha256(actual_source.read_bytes()).hexdigest()
    if actual_sha256.lower() != str(source_info.get("sha256") or "").lower():
        raise ValueError("Input DOCX changed after Dashboard confirmation; reopen the Dashboard")
    if not isinstance(handoff.get("spec"), dict):
        raise ValueError("Dashboard handoff is missing its confirmed specification")
    return handoff


def normalize_spec(data: dict[str, Any]) -> dict[str, Any]:
    if "spec_patch" in data:
        return data["spec_patch"]
    if "spec" in data:
        return data["spec"]
    return data


def validate_spec(spec: dict[str, Any]) -> list[str]:
    from word_format.spec import validate_spec as validate_contract
    try:
        validate_contract(spec)
        return []
    except (ValueError, TypeError, OverflowError, AttributeError) as exc:
        return [str(exc)]


def apply_document(
    input_path: str | Path,
    output_path: str | Path,
    spec: dict[str, Any],
    *,
    clear_direct: bool = False,
    verification: dict[str, Any] | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    from word_format.api import apply_spec as execute_spec
    report = execute_spec(
        input_path, output_path, normalize_spec(spec),
        clear_direct=clear_direct, verification=verification, overwrite=overwrite,
    )
    if verification and verification.get("visual_execution_mode") == "dashboard-direct":
        report["verification"]["visual"] = {
            "status": "skipped",
            "reason": "本地下载模式未连接视觉 AI，请运行显式渲染验收",
        }
    report["clear_direct_font_formatting"] = clear_direct
    return report


def main() -> int:
    from cli_output import configure_cli_output
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source DOCX")
    parser.add_argument("output", help="Destination DOCX; must not equal input")
    parser.add_argument("--handoff", required=True, help="Confirmed Dashboard HANDOFF.json")
    parser.add_argument(
        "--spec",
        help="Task-local spec that may add validated citations to the confirmed handoff",
    )
    parser.add_argument("--report", help="Write application report JSON")
    args = parser.parse_args()
    from json_output import ensure_protected_paths, write_json
    protected = (args.input, args.output, args.handoff, args.spec)
    try:
        ensure_protected_paths(args.report, *protected)
        handoff = load_confirmed_handoff(args.handoff, args.input, args.output)
        if args.spec:
            spec = normalize_spec(load_json(args.spec))
            if not isinstance(spec, dict):
                raise ValueError("Task-local specification must be a JSON object")
            if _without_citations(spec) != _without_citations(handoff["spec"]):
                raise ValueError("Task-local specification may only add validated citations to the confirmed handoff")
        else:
            spec = copy.deepcopy(handoff["spec"])
        report = apply_document(
            args.input,
            args.output,
            spec,
            clear_direct=bool(handoff.get("clear_direct_font_formatting", False)),
            verification=handoff.get("verification") if isinstance(handoff.get("verification"), dict) else None,
        )
        if args.report:
            write_json(args.report, report, *protected)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if not args.report:
        print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
