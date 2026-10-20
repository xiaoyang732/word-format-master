#!/usr/bin/env python3
"""Apply a validated Format Spec directly through the shared formatting core."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from word_format.spec import legacy_paths
from word_format.structures import append_field as append_field
from word_format.structures import apply_page_numbers as apply_page_numbers

SUPPORTED_APPLICATION_PATHS = legacy_paths()


def load_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def normalize_spec(data: dict[str, Any]) -> dict[str, Any]:
    return data.get("spec_patch", data.get("spec", data))


def validate_spec(spec: dict[str, Any]) -> list[str]:
    from word_format.spec import validate_spec as validate_contract

    try:
        validate_contract(spec)
        return []
    except (ValueError, TypeError, OverflowError, AttributeError) as exc:
        return [str(exc)]


def apply_document(
    input_path, output_path, spec, *, clear_direct=False, verification=None, overwrite=False
):
    from word_format.api import apply_spec

    report = apply_spec(
        input_path,
        output_path,
        normalize_spec(spec),
        clear_direct=clear_direct,
        verification=verification,
        overwrite=overwrite,
    )
    report["clear_direct_font_formatting"] = clear_direct
    return report


def main() -> int:
    from cli_output import configure_cli_output
    from json_output import ensure_protected_paths, write_json

    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--spec", required=True, help="Format Spec JSON")
    parser.add_argument("--report")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    protected = (args.input, args.output, args.spec)
    try:
        ensure_protected_paths(args.report, *protected)
        report = apply_document(
            args.input, args.output, load_json(args.spec), overwrite=args.overwrite
        )
        if args.report:
            write_json(args.report, report, *protected)
        else:
            print(json.dumps(report, ensure_ascii=False, indent=2))
    except (OSError, ValueError, TypeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
