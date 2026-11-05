"""UTF-8 output contract for CLI reports, including redirected Windows pipes."""

import sys
from pathlib import Path


def configure_cli_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def summarize_result(command, result, artifact):
    """Return small diagnostics for an already persisted, complete CLI result."""
    verification = result.get("verification", result)
    format_result = verification.get("format", {})
    checks = result.get("checks", format_result.get("checks", []))
    failed = [(index, check) for index, check in enumerate(checks) if check["status"] == "failed"]
    issues = result.get("issues", [])
    counts = {
        key: len(result[key])
        for key in ("objects", "operations", "rules", "templates", "issues", "errors")
        if isinstance(result.get(key), list)
    }
    if "checks" in result or "checks" in format_result:
        counts.update(checks=len(checks), failed_checks=len(failed))
    if result.get("kind") == "word-format-plan":
        counts["targets"] = sum(len(item["targets"]) for item in result["operations"])
    states = {}
    for key in ("format", "structure", "preservation", "visual", "compliance"):
        state = verification.get(key) or result.get(key)
        if isinstance(state, dict) and "status" in state:
            states[key] = state["status"]
            if state.get("errors"):
                counts[f"{key}_errors"] = len(state["errors"])
    summary = {
        "command": command,
        "status": result.get("status", "completed"),
        "artifact": str(Path(artifact).resolve()),
        "counts": counts,
        "states": states,
    }
    for key in ("execution_status", "plan_sha256", "output_sha256", "output_path", "output"):
        if isinstance(result.get(key), str):
            summary[key] = result[key]
    if isinstance(result.get("source"), dict) and "sha256" in result["source"]:
        summary["source_sha256"] = result["source"]["sha256"]
    if issues:
        summary["issues"] = [
            {
                "index": index,
                **{key: str(issue[key])[:100] for key in ("id", "kind", "code") if key in issue},
            }
            for index, issue in enumerate(issues[:5])
        ]
        summary["omitted_issues"] = max(0, len(issues) - 5)
    if failed:
        summary["failed_checks"] = [
            {
                "index": index,
                "operation_id": check["operation_id"][:100],
                "target_id": (check.get("output_target") or check["target"])["id"][:100],
            }
            for index, check in failed[:5]
        ]
        summary["omitted_failed_checks"] = max(0, len(failed) - 5)
    return summary
