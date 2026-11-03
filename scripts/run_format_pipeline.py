"""Run the shared template engine for an existing document or a generated draft."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from json_output import ensure_protected_paths
from word_format.templates import load_template

SKILL_DIR = Path(__file__).resolve().parent.parent


def run_pipeline(
    input_path,
    output_path,
    *,
    template=None,
    profile=None,
    resolutions=None,
    match_structure=False,
    mapping=None,
    report_dir=None,
    overwrite=False,
    render=False,
):
    if bool(template) == bool(profile):
        raise ValueError("Choose exactly one of template or profile")
    if (resolutions or match_structure) and not profile:
        raise ValueError("resolutions and match_structure require a profile")
    source, output = Path(input_path).resolve(), Path(output_path).resolve()
    reports = Path(report_dir).resolve() if report_dir else output.parent / f"{output.stem}-reports"
    protected = [source, output, profile, mapping, resolutions]
    if profile:
        evidence = json.loads(Path(profile).read_text(encoding="utf-8-sig"))
        protected.append(evidence["source"]["path"])
    else:
        definition, authority = load_template(template)
        protected.extend([authority["path"], *[s["path"] for s in definition["sources"]]])
    # Later stages consume the original input files too. Validate every fixed
    # artifact before publishing anything, not just the first plan destination.
    artifacts = [reports / name for name in
                 ("plan.json", "apply.json", "verification.json", "numbering.json")]
    if profile and match_structure:
        artifacts.append(reports / "structure.json")
    if render:
        artifacts.append(reports / "render" / "render-manifest.json")
    for artifact in artifacts:
        ensure_protected_paths(artifact, *protected)
    reports.mkdir(parents=True, exist_ok=True)
    cli = str(SKILL_DIR / "scripts/format_cli.py")

    def run(*args):
        subprocess.run([sys.executable, cli, *map(str, args)], check=True)

    plan, report, verification = (
        reports / name for name in ("plan.json", "apply.json", "verification.json")
    )
    if profile and match_structure:
        run("check-structure", source, "--profile", profile, "--output", reports / "structure.json")
    authority = ["--profile", profile] if profile else ["--template", template]
    args = ["plan", source, output, *authority, "--output", plan]
    if mapping:
        args.extend(["--mapping", mapping])
    if resolutions:
        args.extend(["--resolutions", resolutions])
    if match_structure:
        args.append("--match-structure")
    run(*args)
    run("apply", "--plan", plan, "--report", report, *(["--overwrite"] if overwrite else []))
    run("verify", "--plan", plan, "--report", report, "--output", verification)
    # A format-only pass cannot certify a generated document's numbering.
    subprocess.run(
        [
            sys.executable,
            str(SKILL_DIR / "scripts/check_numbering.py"),
            str(output),
            "--output",
            str(reports / "numbering.json"),
        ],
        check=True,
    )
    if render:
        subprocess.run(
            [
                sys.executable,
                str(SKILL_DIR / "scripts/render_docx.py"),
                str(output),
                "--output-dir",
                str(reports / "render"),
                "--method",
                "auto",
            ],
            check=True,
        )
    result = json.loads(verification.read_text(encoding="utf-8"))
    print(f"输出：{output}\n验收：{result['status']}\n报告：{verification}")
    return result


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("output")
    authority = parser.add_mutually_exclusive_group(required=True)
    authority.add_argument("--template", help="Builtin ID or compiled user template JSON")
    authority.add_argument("--profile", help="Evidence profile, including a prepared report sample")
    parser.add_argument("--resolutions")
    parser.add_argument("--match-structure", action="store_true")
    parser.add_argument("--mapping")
    parser.add_argument("--report-dir")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--render", action="store_true")
    args = parser.parse_args()
    run_pipeline(
        args.input,
        args.output,
        template=args.template,
        profile=args.profile,
        resolutions=args.resolutions,
        match_structure=args.match_structure,
        mapping=args.mapping,
        report_dir=args.report_dir,
        overwrite=args.overwrite,
        render=args.render,
    )


if __name__ == "__main__":
    main()
