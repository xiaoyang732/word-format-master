"""Run the shared template engine for an existing document or a generated draft."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent


def run_pipeline(
    input_path,
    output_path,
    *,
    template,
    mapping=None,
    report_dir=None,
    overwrite=False,
    render=False,
):
    source, output = Path(input_path).resolve(), Path(output_path).resolve()
    reports = Path(report_dir).resolve() if report_dir else output.parent / f"{output.stem}-reports"
    reports.mkdir(parents=True, exist_ok=True)
    cli = str(SKILL_DIR / "scripts/format_cli.py")

    def run(*args):
        subprocess.run([sys.executable, cli, *map(str, args)], check=True)

    plan, report, verification = (
        reports / name for name in ("plan.json", "apply.json", "verification.json")
    )
    args = ["plan", source, output, "--template", template, "--output", plan]
    if mapping:
        args.extend(["--mapping", mapping])
    run(*args)
    run("apply", "--plan", plan, "--report", report, *(["--overwrite"] if overwrite else []))
    run("verify", "--plan", plan, "--report", report, "--output", verification)
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
    parser.add_argument(
        "--template", required=True, help="Builtin ID or compiled user template JSON"
    )
    parser.add_argument("--mapping")
    parser.add_argument("--report-dir")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--render", action="store_true")
    args = parser.parse_args()
    run_pipeline(
        args.input,
        args.output,
        template=args.template,
        mapping=args.mapping,
        report_dir=args.report_dir,
        overwrite=args.overwrite,
        render=args.render,
    )


if __name__ == "__main__":
    main()
