#!/usr/bin/env python3
"""Direct AI formatting: inspect, extract profiles, plan, apply and verify."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from json_output import ensure_protected_paths, write_json
from word_format import (
    apply_plan,
    audit_document,
    build_plan,
    capabilities,
    inspect_document,
    verify_plan,
)
from word_format.contracts import FormatError
from word_format.profiles import (
    check_template_structure,
    profile_plan,
    requirements_profile,
    template_profile,
)


def read_json(path):
    value = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise FormatError("JSON input must be an object")
    return value


def main():
    from cli_output import configure_cli_output

    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    p = commands.add_parser("capabilities")
    p.add_argument("--output")
    p = commands.add_parser("inspect")
    p.add_argument("input")
    p.add_argument("--output")
    p = commands.add_parser(
        "profile", help="Extract evidence; unresolved issues are not silently applied"
    )
    p.add_argument("kind", choices=["requirements", "template"])
    p.add_argument("input")
    p.add_argument("--output")
    p = commands.add_parser(
        "check-structure", help="Check a draft against exactly the sample's parts"
    )
    p.add_argument("input")
    p.add_argument("--profile", required=True)
    p.add_argument("--output")
    p = commands.add_parser("audit")
    p.add_argument("input")
    p.add_argument("--request", required=True, help="JSON request file")
    p.add_argument("--output")
    p = commands.add_parser("plan")
    p.add_argument("input")
    p.add_argument("destination")
    authority = p.add_mutually_exclusive_group(required=True)
    authority.add_argument("--request", help="JSON request file (not raw prose)")
    authority.add_argument("--profile", help="Extracted format profile JSON")
    p.add_argument("--resolutions", help="JSON decisions for profile issues")
    p.add_argument("--mapping", help="JSON section and rule target mappings")
    p.add_argument("--request-text", default="", help="Human instruction retained in plan")
    p.add_argument(
        "--match-structure",
        action="store_true",
        help="Required for sample-based drafts: reject extra, missing or reordered parts",
    )
    p.add_argument("--output")
    p = commands.add_parser("apply")
    p.add_argument("--plan", required=True)
    p.add_argument("--report")
    p.add_argument("--overwrite", action="store_true")
    p = commands.add_parser("verify")
    p.add_argument("--plan", required=True)
    p.add_argument("--report", required=True)
    p.add_argument("--visual-report", help="AI page review bound to output and manifest SHA-256")
    p.add_argument("--render-manifest", help="render_docx.py render-manifest.json")
    p.add_argument(
        "--render-method",
        choices=["auto", "word", "libreoffice"],
        help="Defaults to the plan's renderer preference",
    )
    p.add_argument("--output")
    args = parser.parse_args()
    try:
        destination = args.report if args.command == "apply" else args.output
        protected = [
            Path(v).resolve()
            for v in (
                getattr(args, name, None)
                for name in (
                    "input",
                    "destination",
                    "request",
                    "plan",
                    "profile",
                    "resolutions",
                    "mapping",
                    "visual_report",
                    "render_manifest",
                )
            )
            if v
        ]
        if args.command in {"apply", "verify"}:
            plan = read_json(args.plan)
            protected.extend([plan["source"]["path"], plan["output_path"]])
            if "profile_context" in plan.get("request", {}):
                protected.append(plan["request"]["profile_context"]["source"]["path"])
        if args.command == "verify":
            protected.append(args.report)
            if args.render_manifest:
                manifest = read_json(args.render_manifest)
                protected.extend(
                    p["path"]
                    for p in manifest.get("pages", [])
                    if isinstance(p, dict) and "path" in p
                )
        if args.command in {"plan", "check-structure"} and args.profile:
            profile = read_json(args.profile)
            protected.append(profile["source"]["path"])
        if args.command in {"plan", "audit"} and args.request:
            request = read_json(args.request)
            if "profile_context" in request:
                protected.append(request["profile_context"]["source"]["path"])
        ensure_protected_paths(destination, *protected)
        if args.command == "capabilities":
            result = capabilities()
        elif args.command == "inspect":
            result = inspect_document(args.input)
        elif args.command == "profile":
            result = (requirements_profile if args.kind == "requirements" else template_profile)(
                args.input
            )
        elif args.command == "check-structure":
            result = check_template_structure(args.input, profile)
        elif args.command == "audit":
            result = audit_document(args.input, request)
        elif args.command == "plan":
            if args.profile:
                result = profile_plan(
                    args.input,
                    args.destination,
                    profile,
                    resolutions=read_json(args.resolutions) if args.resolutions else None,
                    mapping=read_json(args.mapping) if args.mapping else None,
                    request_text=args.request_text,
                    match_structure=args.match_structure,
                )
            else:
                if args.mapping or args.resolutions or args.match_structure:
                    raise FormatError(
                        "--mapping, --resolutions and --match-structure require --profile"
                    )
                result = build_plan(args.input, args.destination, request)
        elif args.command == "apply":
            result = apply_plan(plan, overwrite=args.overwrite)
        else:
            result = verify_plan(
                plan,
                read_json(args.report),
                visual_report=read_json(args.visual_report) if args.visual_report else None,
                render_manifest=args.render_manifest,
                render_method=args.render_method,
            )
        if destination:
            write_json(destination, result, *protected)
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False))
        return 2 if result.get("status") in {"failed", "unsupported", "needs_clarification"} else 0
    except (FormatError, ValueError, OSError, KeyError, TypeError) as exc:
        print(
            json.dumps(
                {"status": getattr(exc, "status", "failed"), "error": str(exc)}, ensure_ascii=False
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
