"""Run formatting with resumable, hash-bound stages and explicit acceptance states."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import time

from json_output import ensure_not_in_skill_root, ensure_protected_paths, write_json
from pipeline_state import PipelineState, digest, task_lock
from word_format.templates import load_template

SKILL_DIR = Path(__file__).resolve().parent.parent


def error_detail(exc):
    detail = getattr(exc, "stderr", None) or str(exc)
    return detail.decode("utf-8", errors="replace") if isinstance(detail, bytes) else detail


def run_pipeline(input_path, output_path, *, template=None, profile=None, resolutions=None,
                 match_structure=False, flexible_chapters=False, mapping=None, report_dir=None, overwrite=False,
                 render=False, resume=False, render_retries=0, render_timeout=300,
                 render_method="auto", visual_report=None, check_quality=False,
                 min_chinese_characters=None, max_chinese_characters=None,
                 min_abstract_chinese_characters=None, max_abstract_chinese_characters=None,
                 min_references=None, require_all_references_cited=False,
                 min_recent_references=None, recent_start_year=None):
    source, output = Path(input_path).resolve(), Path(output_path).resolve()
    reports = Path(report_dir).resolve() if report_dir else output.parent / f"{output.stem}-reports"
    ensure_not_in_skill_root(output)
    ensure_not_in_skill_root(reports)
    if resume and (reports / "pipeline-state.json").is_file():
        try:
            saved_id = json.loads((reports / "pipeline-state.json").read_text(encoding="utf-8")).get("identity", {})
        except (OSError, ValueError):
            saved_id = {}
        if isinstance(saved_id, dict):
            if not template and not profile:
                if saved_id.get("profile") not in {None, "None"}:
                    profile = saved_id["profile"]
                elif saved_id.get("template") not in {None, "None"}:
                    template = saved_id["template"]
            if mapping is None and saved_id.get("mapping") not in {None, "None"}:
                mapping = saved_id["mapping"]
            if resolutions is None and saved_id.get("resolutions") not in {None, "None"}:
                resolutions = saved_id["resolutions"]
            if not match_structure and saved_id.get("match_structure"):
                match_structure = True
            if not flexible_chapters and saved_id.get("flexible_chapters"):
                flexible_chapters = True
            if not overwrite and saved_id.get("overwrite"):
                overwrite = True
            if not render and (saved_id.get("render") or visual_report):
                render = bool(saved_id.get("render"))
            if render_method == "auto" and saved_id.get("render_method") in {"word", "libreoffice"}:
                render_method = saved_id["render_method"]
            saved_q = saved_id.get("quality_targets")
            if isinstance(saved_q, dict):
                check_quality = True
                if min_chinese_characters is None:
                    min_chinese_characters = saved_q.get("min_chinese_characters")
                if max_chinese_characters is None:
                    max_chinese_characters = saved_q.get("max_chinese_characters")
                if min_abstract_chinese_characters is None:
                    min_abstract_chinese_characters = saved_q.get("min_abstract_chinese_characters")
                if max_abstract_chinese_characters is None:
                    max_abstract_chinese_characters = saved_q.get("max_abstract_chinese_characters")
                if min_references is None:
                    min_references = saved_q.get("min_references")
                if not require_all_references_cited and saved_q.get("require_all_references_cited"):
                    require_all_references_cited = True
                if min_recent_references is None:
                    min_recent_references = saved_q.get("min_recent_references")
                if recent_start_year is None:
                    recent_start_year = saved_q.get("recent_start_year")
    quality_requested = bool(
        check_quality
        or require_all_references_cited
        or any(
            v is not None
            for v in (
                min_chinese_characters,
                max_chinese_characters,
                min_abstract_chinese_characters,
                max_abstract_chinese_characters,
                min_references,
                min_recent_references,
                recent_start_year,
            )
        )
    )
    if bool(template) == bool(profile):
        raise ValueError("Choose exactly one of template or profile")
    if (resolutions or match_structure) and not profile:
        raise ValueError("resolutions and match_structure require a profile")
    if not 0 <= render_retries <= 3 or render_timeout < 1:
        raise ValueError("渲染重试次数须为 0–3，超时须大于 0")
    if render_method not in {"auto", "word", "libreoffice"}:
        raise ValueError("Unknown renderer")
    if visual_report and not render:
        raise ValueError("视觉验收报告需要 --render 及有效渲染证据")
    protected = [source, output, profile, mapping, resolutions, visual_report]
    if profile:
        evidence = json.loads(Path(profile).read_text(encoding="utf-8-sig"))
        protected.append(evidence["source"]["path"])
    else:
        definition, authority = load_template(template)
        protected.extend([authority["path"], *[s["path"] for s in definition["sources"]]])
    if render:
        render_root = (reports / "render").resolve()
        if any(p and (Path(p).resolve() == render_root or render_root in Path(p).resolve().parents)
               for p in protected):
            raise ValueError("渲染目录不能包含输入文档、模板、映射、决策或视觉报告；请选择独立目录")
    # Validate every artifact, including state/lock/summary, before the first write.
    artifacts = [reports / name for name in ("plan.json", "apply.json", "verification.json",
                 "numbering.json", "structure.json", "pipeline-state.json", "summary.json", ".pipeline.lock",
                 "visual-verification.json", "quality.json")]
    if render:
        artifacts.append(reports / "render/render-manifest.json")
    for artifact in artifacts:
        ensure_protected_paths(artifact, *protected)
    # A page review is produced after rendering; permit attaching it on resume.
    files = {str(Path(p).resolve()): digest(p) for p in protected
             if p and Path(p).resolve() != output and p != visual_report}
    engine = {p.relative_to(SKILL_DIR).as_posix(): digest(p) for p in (SKILL_DIR / "scripts").rglob("*.py")}
    identity = {"files": files, "engine": hashlib.sha256(json.dumps(engine, sort_keys=True).encode()).hexdigest(),
                "output": str(output), "template": str(template), "profile": str(profile),
                "mapping": str(mapping), "resolutions": str(resolutions), "match_structure": match_structure,
                "overwrite": overwrite, "render": render, "render_method": render_method}
    if flexible_chapters:
        identity["flexible_chapters"] = True
    if quality_requested:
        identity["quality_targets"] = {
            "min_chinese_characters": min_chinese_characters,
            "max_chinese_characters": max_chinese_characters,
            "min_abstract_chinese_characters": min_abstract_chinese_characters,
            "max_abstract_chinese_characters": max_abstract_chinese_characters,
            "min_references": min_references,
            "require_all_references_cited": bool(require_all_references_cited),
            "min_recent_references": min_recent_references,
            "recent_start_year": recent_start_year,
        }
    reports.mkdir(parents=True, exist_ok=True)
    cli = str(SKILL_DIR / "scripts/format_cli.py")
    plan, applied, verification = (reports / n for n in ("plan.json", "apply.json", "verification.json"))
    state = None

    def summary(failed=False):
        stages = state.data["stages"]
        checks = {key: ("passed" if stages.get(stage, {}).get("status") == "succeeded" else
                        "failed" if stages.get(stage, {}).get("status") == "failed" else "pending")
                  for key, stage in (("format", "verify"), ("structure", "structure"),
                                     ("numbering", "numbering"), ("render", "render"))}
        if not (profile and match_structure):
            checks["structure"] = "skipped"
        elif checks["structure"] == "passed" and checks["format"] != "passed":
            # The early structure stage checks the draft; final acceptance also
            # requires the current output to have valid independent verification.
            checks["structure"] = "pending"
        if not render:
            checks["render"] = "skipped"
        checks["visual"] = "pending" if render else "skipped"
        if stages.get("visual", {}).get("status") == "failed":
            checks["visual"] = "failed"
        if stages.get("visual", {}).get("status") == "succeeded":
            checks["visual"] = json.loads((reports / "visual-verification.json").read_text(encoding="utf-8"))["visual"]["status"]
        if quality_requested:
            q_stage = stages.get("quality", {}).get("status")
            checks["content"] = "passed" if q_stage == "succeeded" else "failed" if q_stage == "failed" else "pending"
            content_note = "已在此流水线中核验字数、摘要长度、文献数量与正文引用映射；事实与文献真实性仍需人工或联网核实"
        else:
            checks["content"] = "not_performed"
            content_note = "字数、引用对应及事实与文献真实性均未在此流水线中核验"
        from word_format.remediation import build_pipeline_remediation
        value = {"schema_version": "1.0", "task_id": state.data["task_id"],
                 "status": "failed" if failed else "pending" if render and checks["visual"] != "passed" else "passed",
                 "execution_status": "failed" if failed else "completed", "checks": checks,
                 "output": str(output), "output_sha256": digest(output) if output.is_file() else None,
                 "failure": state.data.get("failure"),
                 "remediation": build_pipeline_remediation(checks, state.data.get("failure"), reports, output),
                 "content_note": content_note}
        write_json(reports / "summary.json", value, *protected)
        return value

    def stage(name, command, paths, *, retries=0):
        # Validate the saved manifest and pages before parsing the manifest.
        # Failed/incomplete renders will regenerate their inventory after running.
        expected = (list(state.data["stages"].get(name, {}).get("artifacts", {}))
                    if callable(paths) else paths)
        if state.reusable(name, expected):
            return
        if name == "apply" and output.exists() and resume:
            raise ValueError("输出已存在但缺少有效 apply 证据，拒绝自动覆盖；请选择新的输出及任务目录")
        for attempt in range(retries + 1):
            state.start(name)
            try:
                result = subprocess.run(command, check=True, capture_output=True, encoding="utf-8",
                                        errors="backslashreplace", timeout=render_timeout if name == "render" else None)
                state.succeeded(name, paths() if callable(paths) else paths)
                return result
            except (subprocess.SubprocessError, OSError, ValueError) as exc:
                detail = error_detail(exc)
                state.failed(name, exc)
                state.data["failure"]["detail"] = detail
                transient = isinstance(exc, (subprocess.TimeoutExpired, PermissionError)) or any(
                    text in detail.lower() for text in ("timeout", "timed out", "超时"))
                state.data["failure"]["retryable"] = bool(name == "render" and transient)
                state.save()
                if attempt >= retries or not transient:
                    raise
                time.sleep(min(attempt + 1, 3))

    def render_artifacts():
        manifest = reports / "render/render-manifest.json"
        paths = [manifest]
        if manifest.is_file():
            data = json.loads(manifest.read_text(encoding="utf-8"))
            if data.get("source_sha256") != digest(output) or Path(data.get("source", "")).resolve() != output:
                raise ValueError("渲染产物与输出摘要不一致")
            for page in data.get("pages", []):
                path = Path(page["path"]).resolve()
                if (reports / "render").resolve() not in path.parents:
                    raise ValueError("非法渲染产物路径")
                ensure_protected_paths(path, *protected)
                paths.append(path)
        return paths

    with task_lock(reports / ".pipeline.lock"):
        state = PipelineState(reports / "pipeline-state.json", identity, resume)
        try:
            if not visual_report and state.data["stages"].get("visual", {}).get("status") == "succeeded":
                evidence_paths = state.data["stages"]["visual"].get("artifacts", {})
                verification_path = str((reports / "visual-verification.json").resolve())
                reviews = [p for p in evidence_paths if p != verification_path]
                if len(reviews) != 1 or verification_path not in evidence_paths:
                    raise ValueError("保存的视觉阶段证据不完整，拒绝恢复")
                visual_report = reviews[0]
                # Validate saved evidence before exposing any completed visual state.
                state.reusable("visual", [reports / "visual-verification.json", Path(visual_report)])
            if profile and match_structure:
                struct_cmd = [sys.executable, cli, "check-structure", str(source), "--profile", str(profile),
                              "--output", str(reports / "structure.json")]
                if flexible_chapters:
                    struct_cmd.append("--flexible-chapters")
                stage("structure", struct_cmd, [reports / "structure.json"])
            authority_args = ["--profile", profile] if profile else ["--template", template]
            args = [sys.executable, cli, "plan", str(source), str(output), *authority_args, "--output", str(plan)]
            for flag, value in (("--mapping", mapping), ("--resolutions", resolutions)):
                if value:
                    args.extend([flag, str(value)])
            if match_structure:
                args.append("--match-structure")
            if flexible_chapters:
                args.append("--flexible-chapters")
            stage("plan", list(map(str, args)), [plan])
            stage("apply", [sys.executable, cli, "apply", "--plan", str(plan), "--report", str(applied),
                            *(["--overwrite"] if overwrite else [])], [applied, output])
            stage("verify", [sys.executable, cli, "verify", "--plan", str(plan), "--report", str(applied),
                             "--output", str(verification)], [verification])
            stage("numbering", [sys.executable, str(SKILL_DIR / "scripts/check_numbering.py"), str(output),
                                "--output", str(reports / "numbering.json")], [reports / "numbering.json"])
            if quality_requested:
                q_args = [sys.executable, str(SKILL_DIR / "scripts/thesis_quality.py"), str(output),
                          "--output", str(reports / "quality.json")]
                for flag, val in (
                    ("--min-chinese-characters", min_chinese_characters),
                    ("--max-chinese-characters", max_chinese_characters),
                    ("--min-abstract-chinese-characters", min_abstract_chinese_characters),
                    ("--max-abstract-chinese-characters", max_abstract_chinese_characters),
                    ("--min-references", min_references),
                    ("--min-recent-references", min_recent_references),
                    ("--recent-start-year", recent_start_year),
                ):
                    if val is not None:
                        q_args.extend([flag, str(val)])
                if require_all_references_cited:
                    q_args.append("--require-all-references-cited")
                stage("quality", q_args, [reports / "quality.json"])
            if render:
                stage("render", [sys.executable, str(SKILL_DIR / "scripts/render_docx.py"), str(output),
                                 "--output-dir", str(reports / "render"), "--method", render_method],
                      render_artifacts, retries=render_retries)
            if visual_report:
                stage("visual", [sys.executable, cli, "verify", "--plan", str(plan), "--report", str(applied),
                                 "--visual-report", str(visual_report), "--render-manifest",
                                 str(reports / "render/render-manifest.json"), "--output",
                                 str(reports / "visual-verification.json")], [reports / "visual-verification.json", Path(visual_report)])
            result = json.loads(verification.read_text(encoding="utf-8"))
            overview = summary()
        except (ValueError, OSError, subprocess.SubprocessError):
            summary(failed=True)
            raise
    content_display = overview["checks"]["content"] if quality_requested else "未执行"
    print(f"输出：{output}\n格式核对：{overview['checks']['format']}\n编号核对：{overview['checks']['numbering']}"
          f"\n页面视觉：{overview['checks']['visual']}\n内容核验：{content_display}\n汇总：{reports / 'summary.json'}")
    return result


def main():
    from cli_output import configure_cli_output
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("output")
    authority = parser.add_mutually_exclusive_group(required=False)
    authority.add_argument("--template")
    authority.add_argument("--profile")
    for name in ("resolutions", "mapping", "report-dir", "visual-report"):
        parser.add_argument("--" + name)
    for name in ("match-structure", "flexible-chapters", "overwrite", "render", "resume", "check-quality", "require-all-references-cited"):
        parser.add_argument("--" + name, action="store_true")
    for name in ("min-chinese-characters", "max-chinese-characters", "min-abstract-chinese-characters",
                 "max-abstract-chinese-characters", "min-references", "min-recent-references", "recent-start-year"):
        parser.add_argument("--" + name, type=int)
    parser.add_argument("--render-retries", type=int, default=0)
    parser.add_argument("--render-timeout", type=int, default=300)
    parser.add_argument("--render-method", choices=["auto", "word", "libreoffice"], default="auto")
    args = vars(parser.parse_args())
    try:
        run_pipeline(args.pop("input"), args.pop("output"), **args)
        return 0
    except (ValueError, OSError, subprocess.SubprocessError) as exc:
        print(json.dumps({"status": "failed", "code": "PIPELINE_INCOMPLETE", "error": str(exc),
                          "detail": error_detail(exc),
                          "next_actions": ["查看任务 summary.json；修复原因后用相同参数追加 --resume，证据过期时另建任务"]},
                         ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
