#!/usr/bin/env python3
"""Direct AI formatting: inspect, extract profiles, plan, apply and verify."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Diagnosis must run before importing the formatting engine and its dependencies.
if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "doctor":
    from runtime_doctor import main as doctor_main

    raise SystemExit(doctor_main(sys.argv[2:]))

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
from word_format.cover import apply_template_cover_file
from word_format.profiles import (
    check_template_structure,
    profile_plan,
    requirements_profile,
    template_profile,
)
from word_format.reports import CATALOG, load_report_catalog, prepare_report
from word_format.templates import compile_template, list_templates, load_template, template_plan


def read_json(path):
    value = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict):
        raise FormatError("JSON input must be an object")
    return value


def main():
    from cli_output import configure_cli_output, summarize_result

    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("doctor", help="检查依赖、模板、目录权限与页面渲染能力")
    p = commands.add_parser("capabilities")
    p.add_argument("--output")
    p = commands.add_parser("report", help="Select the default report sample and prepare a draft")
    p.add_argument("kind", choices=["list", "prepare"])
    p.add_argument("workspace", nargs="?", help="New directory for draft, profile and workflow")
    report_source = p.add_mutually_exclusive_group()
    report_source.add_argument(
        "--template", help="Registered report ID; default from templates/reports.json"
    )
    report_source.add_argument(
        "--sample", help="User-supplied DOCX sample takes precedence over default"
    )
    p.add_argument("--topic", default="")
    p.add_argument("--output")
    p = commands.add_parser("inspect")
    p.add_argument("input")
    p.add_argument("--output")
    p = commands.add_parser(
        "profile", help="Extract evidence; unresolved issues are not silently applied"
    )
    p.add_argument("kind", choices=["requirements", "template"])
    p.add_argument("--document-kind", choices=["thesis", "summary", "all"], default="thesis")
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
    authority = p.add_mutually_exclusive_group(required=True)
    authority.add_argument("--request", help="JSON request file")
    authority.add_argument("--template", help="Builtin ID or compiled template JSON")
    p.add_argument("--mapping")
    p.add_argument("--output")
    p = commands.add_parser("template", help="Compile reusable templates without preset fallback")
    p.add_argument("kind", choices=["list", "compile", "validate"])
    p.add_argument("input", nargs="?")
    p.add_argument(
        "--profile", action="append", help="Evidence profile; repeat for multiple sources"
    )
    p.add_argument("--resolutions")
    p.add_argument("--id", default="user-template")
    p.add_argument("--output")
    p = commands.add_parser("plan")
    p.add_argument("input")
    p.add_argument("destination")
    authority = p.add_mutually_exclusive_group(required=True)
    authority.add_argument("--request", help="JSON request file (not raw prose)")
    authority.add_argument("--profile", help="Extracted format profile JSON")
    authority.add_argument("--template", help="Builtin ID or compiled format template JSON")
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
    p = commands.add_parser(
        "cover", help="Clone an official template cover into a new or existing DOCX safely"
    )
    p.add_argument("template", help="Template DOCX/DOTX/DOC/DOT/WPS containing the cover")
    p.add_argument("destination", help="Output DOCX path (never overwrites template or input)")
    p.add_argument("--into", help="Existing draft DOCX whose body content is preserved")
    p.add_argument(
        "--replace-existing-cover",
        action="store_true",
        help="Replace any existing front cover in --into before inserting the template cover",
    )
    p.add_argument(
        "--no-section-break",
        action="store_true",
        help="Separate the cloned cover with a page break instead of a section break",
    )
    p.add_argument(
        "--no-declaration",
        dest="include_declaration",
        action="store_false",
        default=True,
        help="Do not include declaration section from template",
    )
    p.add_argument("--fields", help="JSON object mapping cover field labels/keys to values")
    p.add_argument("--replacements", help="JSON object mapping literal placeholder strings")
    p.add_argument("--report", help="Optional JSON report path")
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
    for name in ("inspect", "profile", "check-structure", "audit", "template", "plan", "apply", "cover", "verify"):
        path_option = "--report" if name in {"apply", "cover"} else "--output"
        commands.choices[name].add_argument(
            "--summary", action="store_true",
            help=f"Print a compact summary after writing the full result to {path_option}",
        )
    args = parser.parse_args()
    try:
        destination = args.report if args.command in {"apply", "cover"} else args.output
        summary = getattr(args, "summary", False)
        if summary and not destination:
            path_option = "--report" if args.command in {"apply", "cover"} else "--output"
            raise FormatError(f"--summary requires {path_option}")
        protected = [
            Path(v).resolve()
            for v in (
                getattr(args, name, None)
                for name in (
                    "input",
                    "destination",
                    "into",
                    "fields",
                    "replacements",
                    "request",
                    "plan",
                    "profile",
                    "resolutions",
                    "mapping",
                    "visual_report",
                    "render_manifest",
                )
            )
            if v and not isinstance(v, list)
        ]
        if args.command == "cover":
            protected.append(Path(args.template).resolve())
        if args.command == "template" and args.profile:
            protected.extend(args.profile)
            protected.extend(read_json(p)["source"]["path"] for p in args.profile)
        if args.command in {"plan", "audit"} and args.template:
            template, template_source = load_template(args.template)
            protected.extend([template_source["path"], *[s["path"] for s in template["sources"]]])
        if args.command == "template" and args.kind == "validate":
            if not args.input:
                raise FormatError("template validate requires a template ID or file")
            template, template_source = load_template(args.input)
            protected.extend([template_source["path"], *[s["path"] for s in template["sources"]]])
        if args.command in {"apply", "verify"}:
            # Protect the selected template and every evidence source as well.
            preview = read_json(args.plan)
            context = preview.get("request", {}).get("template_context")
            if context:
                protected.extend(s["path"] for s in context["definition"]["sources"])
                if context["source"]:
                    protected.append(context["source"]["path"])
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
        if args.command == "report":
            protected.extend([CATALOG, CATALOG.parents[1] / "references/report-generation.md"])
            if args.sample:
                protected.append(args.sample)
            else:
                catalog = load_report_catalog()
                protected.extend(t[key] for t in catalog["templates"]
                                 for key in ("path", "source_notes", "usage_notes"))
            if args.workspace:
                protected.extend(Path(args.workspace) / name for name in
                                 ("draft.docx", "profile.json", "workflow.json"))
            ensure_protected_paths(destination, *protected)
            if args.kind == "list":
                result = load_report_catalog()
            else:
                if not args.workspace:
                    raise FormatError("report prepare requires a new workspace directory")
                result = prepare_report(
                    args.workspace, template_id=args.template, sample=args.sample, topic=args.topic
                )
        elif args.command == "capabilities":
            result = capabilities()
        elif args.command == "inspect":
            result = inspect_document(args.input)
        elif args.command == "template":
            if args.kind == "list":
                result = {"kind": "word-format-templates", "templates": list_templates()}
            elif args.kind == "validate":
                value = template
                result = {
                    "status": "needs_clarification" if value["issues"] else "passed",
                    "template_id": value["id"],
                    "issues": value["issues"],
                }
            else:
                result = compile_template(
                    [read_json(p) for p in args.profile or []],
                    identifier=args.id,
                    resolutions=read_json(args.resolutions) if args.resolutions else None,
                )
        elif args.command == "profile":
            result = (
                requirements_profile(args.input, document_kind=args.document_kind)
                if args.kind == "requirements"
                else template_profile(args.input)
            )
        elif args.command == "check-structure":
            result = check_template_structure(args.input, profile)
        elif args.command == "audit":
            if args.template:
                draft = Path(args.input).resolve()
                bound = template_plan(
                    draft,
                    draft.with_name(f".{draft.stem}.audit-only.docx"),
                    template,
                    mapping=read_json(args.mapping) if args.mapping else None,
                    template_source=template_source,
                )
                result = (
                    bound
                    if bound.get("status") == "needs_clarification"
                    else audit_document(args.input, bound["request"])
                )
            else:
                result = audit_document(args.input, request)
        elif args.command == "plan":
            if args.template:
                if args.resolutions or args.match_structure:
                    raise FormatError(
                        "Compile template decisions first; --match-structure belongs to profile mode"
                    )
                result = template_plan(
                    args.input,
                    args.destination,
                    template,
                    mapping=read_json(args.mapping) if args.mapping else None,
                    template_source=template_source,
                )
            elif args.profile:
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
        elif args.command == "cover":
            result = apply_template_cover_file(
                args.template,
                args.destination,
                into_path=args.into,
                fields=read_json(args.fields) if args.fields else None,
                replacements=read_json(args.replacements) if args.replacements else None,
                replace_existing_cover=args.replace_existing_cover,
                use_section_break=not args.no_section_break,
                include_declaration=args.include_declaration,
                overwrite=args.overwrite,
            )
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
            if summary:
                print(json.dumps(summarize_result(args.command, result, destination),
                                 ensure_ascii=False, allow_nan=False))
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False))
        return 2 if result.get("status") in {"failed", "unsupported", "needs_clarification"} else 0
    except (FormatError, ValueError, OSError, KeyError, TypeError) as exc:
        print(
            json.dumps(
                {"status": getattr(exc, "status", "failed"), "error": str(exc),
                 "code": getattr(exc, "code", "INVALID_INPUT"),
                 "message": getattr(exc, "message", "操作未完成，请检查输入与文件权限"),
                 "retryable": getattr(exc, "retryable", False),
                 "next_actions": getattr(exc, "next_actions", ["检查错误详情后修正输入，再重新执行"])}, ensure_ascii=False
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
