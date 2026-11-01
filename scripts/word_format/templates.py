"""Versioned templates: evidence compilation, target binding and coverage.

Templates contain appearance. This module never loads a fallback preset.
All writes still pass through build_plan and the capability registry.
"""

from __future__ import annotations

import copy
import json
import re
from collections import defaultdict
from pathlib import Path

from docx import Document

from .contracts import FormatError, json_hash, number, require_keys, sha256, source_path, text_value
from .inspect import DocumentIndex
from .profiles import validate_profile, validate_profile_source
from .registry import REGISTRY
from .selectors import resolve_targets

TEMPLATE_VERSION = "1.0"
BUILTINS = Path(__file__).resolve().parents[2] / "templates" / "builtin"


def validate_template(template):
    require_keys(
        template,
        {
            "schema_version",
            "kind",
            "id",
            "name",
            "version",
            "origin",
            "sources",
            "rules",
            "issues",
            "decisions",
            "limitations",
            "format",
            "evidence",
            "fallback_template",
            "metadata",
        },
        {"schema_version", "kind", "id", "version", "origin", "sources", "rules", "issues"},
        "template",
    )
    if template["kind"] != "word-format-template" or template["schema_version"] != TEMPLATE_VERSION:
        raise FormatError("Unsupported format template version")
    text_value(template["id"], "template ID")
    text_value(template["version"], "template version")
    if template["origin"] not in {"builtin", "user"}:
        raise FormatError("Template origin must be builtin or user")
    if template.get("fallback_template") is not None:
        raise FormatError(
            "Implicit template fallback is forbidden; compile explicit overrides instead"
        )
    for key in ("rules", "issues", "sources", "decisions", "limitations"):
        if not isinstance(template.get(key, []), list):
            raise FormatError(f"Template {key} must be an array")
    for source in template["sources"]:
        validate_profile_source(source)
    if "format" in template:
        if template["origin"] != "builtin":
            raise FormatError(
                "User templates must use evidence-bearing rules, without preset format fallback"
            )
        from .spec import validate_spec

        validate_spec(template["format"])
    seen = set()
    for rule in template["rules"]:
        require_keys(
            rule,
            {"id", "state", "action", "target", "params", "evidence", "required", "reason"},
            {"id", "state", "action", "target", "params", "evidence"},
            "template rule",
        )
        identifier = text_value(rule["id"], "rule ID")
        if identifier in seen:
            raise FormatError("Duplicate template rule ID")
        seen.add(identifier)
        if rule["state"] not in {"set", "preserve"}:
            raise FormatError(
                "Executable rule state must be set or preserve; put unresolved facts in issues"
            )
        text_value(rule["action"], "rule action")
        cap = REGISTRY.get(rule["action"])
        if (
            cap is None
            or not isinstance(rule["target"], dict)
            or rule["target"].get("type") not in cap.targets
        ):
            raise FormatError("Template rule needs a registered action and target", "unsupported")
        cap.validator(rule["params"])
        if not isinstance(rule["evidence"], list) or not rule["evidence"]:
            raise FormatError("Every template rule requires evidence")
        for evidence in rule["evidence"]:
            if not isinstance(evidence, dict):
                raise FormatError("Rule evidence must be an object")
            text_value(evidence.get("anchor"), "evidence anchor")
            if "confidence" in evidence:
                number(evidence["confidence"], 0, 1, "evidence confidence")
        if not isinstance(rule.get("required", True), bool):
            raise FormatError("Rule required must be boolean")
        if rule["state"] == "preserve":
            text_value(rule.get("reason"), "preservation reason")
    for issue in template["issues"]:
        if not isinstance(issue, dict):
            raise FormatError("Template issues must be objects")
        identifier = text_value(issue.get("id"), "issue ID")
        if identifier in seen:
            raise FormatError("Duplicate template item ID")
        seen.add(identifier)
        text_value(issue.get("reason"), "issue reason")
    if not isinstance(template.get("evidence", []), list):
        raise FormatError("Template evidence must be an array")
    for evidence in template.get("evidence", []):
        if not isinstance(evidence, dict) or not isinstance(evidence.get("snapshot", {}), dict):
            raise FormatError("Template evidence needs an object snapshot")
    return template


def load_template(identifier):
    candidate = Path(identifier)
    if candidate.is_file():
        path = candidate.resolve()
    elif re.fullmatch(r"[a-z][a-z0-9-]*", str(identifier)):
        path = BUILTINS / str(identifier) / "template.json"
    else:
        raise FormatError("Choose an existing template JSON or a registered builtin ID")
    if not path.is_file():
        raise FormatError(f"Template does not exist: {identifier}")
    template = json.loads(path.read_text(encoding="utf-8-sig"))
    validate_template(template)
    return template, {"path": str(path), "sha256": sha256(path)}


def list_templates():
    return [
        {"id": t["id"], "name": t.get("name", t["id"]), "version": t["version"]}
        for path in sorted(BUILTINS.glob("*/template.json"))
        for t in [json.loads(path.read_text(encoding="utf-8"))]
    ]


def resolve_issues(template, resolutions=None):
    result = copy.deepcopy(template)
    resolutions = resolutions or {}
    if not isinstance(resolutions, dict) or set(resolutions) - {i["id"] for i in result["issues"]}:
        raise FormatError("Resolution IDs must belong to this template")
    pending = []
    for issue in result["issues"]:
        decision = resolutions.get(issue["id"])
        if decision is None:
            pending.append(issue)
            continue
        require_keys(
            decision,
            {"decision", "reason", "candidate", "rules"},
            {"decision", "reason"},
            "template decision",
        )
        text_value(decision["reason"], "decision reason")
        mode = decision["decision"]
        if mode == "choose":
            candidates = issue.get("candidates", [])
            n = decision.get("candidate")
            if isinstance(n, bool) or not isinstance(n, int) or not 0 <= n < len(candidates):
                raise FormatError("Decision candidate must identify an existing rule")
            additions = [candidates[n]]
        elif mode == "replace":
            additions = decision.get("rules")
            if not isinstance(additions, list) or not additions:
                raise FormatError("Replacement decision requires rules")
        elif mode in {"ignore", "not_applicable"}:
            additions = []
            if mode == "ignore":
                result.setdefault("limitations", []).append(
                    {"id": issue["id"], "reason": decision["reason"]}
                )
        else:
            raise FormatError("Decision must be choose, replace, ignore or not_applicable")
        for i, rule in enumerate(additions, 1):
            result["rules"].append(
                {
                    **rule,
                    "id": f"{issue['id']}.resolved{i}",
                    "state": "set",
                    "evidence": rule.get("evidence")
                    or [{"anchor": issue["id"], "decision": decision}],
                }
            )
        result.setdefault("decisions", []).append({"issue": issue, "resolution": decision})
    result["issues"] = pending
    return result


def compile_template(profiles, *, identifier="user-template", resolutions=None):
    if not isinstance(profiles, list) or not profiles:
        raise FormatError("Compile at least one evidence profile")
    result = {
        "schema_version": TEMPLATE_VERSION,
        "kind": "word-format-template",
        "id": identifier,
        "version": "1",
        "origin": "user",
        "sources": [],
        "rules": [],
        "issues": [],
        "decisions": [],
        "limitations": [],
        "evidence": [],
    }
    for n, profile in enumerate(profiles, 1):
        validate_profile(profile)
        result["sources"].append(copy.deepcopy(profile["source"]))
        result["evidence"].append(
            {
                "source": n,
                "profile_sha256": json_hash(profile),
                "coverage": profile.get("coverage", {}),
                "snapshot": profile.get("snapshot", {}),
            }
        )
        for rule in profile["rules"]:
            result["rules"].append(
                {
                    **copy.deepcopy(rule),
                    "id": f"s{n}.{rule['id']}",
                    "state": "set",
                    "evidence": [{**e, "source": n} for e in rule["evidence"]],
                }
            )
        for issue in profile["issues"]:
            result["issues"].append(
                {**copy.deepcopy(issue), "id": f"s{n}.{issue['id']}", "source": n}
            )
    # Merge identical evidence; conflicting values remain decisions, never last-wins.
    grouped = defaultdict(list)
    for rule in result["rules"]:
        grouped[(json_hash(rule["target"]), rule["action"])].append(rule)
    result["rules"] = []
    for group in grouped.values():
        if len({json_hash(r["params"]) for r in group}) > 1:
            result["issues"].append(
                {
                    "id": f"conflict{len(result['issues']) + 1}",
                    "reason": "Sources disagree for the same target/property",
                    "candidates": group,
                }
            )
        else:
            rule = group[0]
            rule["evidence"] = [e for r in group for e in r["evidence"]]
            result["rules"].append(rule)
    result = resolve_issues(result, resolutions)
    validate_template(result)
    return result


def _heading_structure(index, operations, preserved=()):
    from .properties import read_font

    touched = set()
    explicit_style = set()
    explicit_outline = set()
    explicit_color = set()
    for op in operations:
        cap = REGISTRY[op["action"]]
        if cap.property == "heading_numbering":
            touched.update(
                oid for oid in index.top_ids if index.records[oid].get("role") == "heading"
            )
        if op["target"]["type"] in {"paragraph", "text_range"}:
            refs = resolve_targets(index, op["target"], cap.targets)
            touched.update(r["id"] for r in refs)
            if op["action"] == "paragraph.style.assign":
                explicit_style.update(r["id"] for r in refs)
            if op["action"] == "paragraph.outline_level.set":
                explicit_outline.update(r["id"] for r in refs)
            if op["action"] == "font.color.set":
                explicit_color.update(r["id"] for r in refs)
        elif op["target"]["type"] == "style" and op["action"] == "font.color.set":
            style_name = op["target"].get("name")
            explicit_color.update(
                oid
                for oid in index.top_ids
                if index.records[oid].get("role") == "heading"
                and (
                    index.records[oid].get("style_name") == style_name
                    or (
                        oid not in explicit_style
                        and index.records[oid].get("module") == "chapters"
                        and style_name
                        == f"Heading {index.records[oid].get('heading_level') or 1}"
                    )
                )
            )
    result = []
    front_matter_modules = {"cover", "declaration", "abstract", "toc"}
    for oid in index.top_ids:
        record = index.records[oid]
        if (
            oid in preserved
            or oid not in touched
            or record.get("role") != "heading"
            or record.get("module") == "toc"
        ):
            continue
        target = {"type": "paragraph", "id": oid}
        level = record.get("heading_level") or 1
        assigned_style = None
        if oid not in explicit_style:
            assigned_style = (
                f"Heading {level}"
                if record.get("module") == "chapters"
                else f"WFM {record.get('module')} Heading"
            )
            result.append(
                {
                    "id": f"semantic.{oid}.style",
                    "action": "paragraph.style.assign",
                    "target": target,
                    "params": {"name": assigned_style},
                }
            )
        # Preserve effective heading color so assigning Office's blue Heading N style
        # or enabling automatic numbering never turns black headings/numbers blue.
        if oid not in explicit_color:
            run_colors = index.font_values(oid, "color")
            if len(run_colors) == 1:
                vis_color = run_colors[0]
                if isinstance(vis_color, str) and (
                    vis_color == "auto" or re.fullmatch(r"[0-9a-fA-F]{6}", vis_color)
                ):
                    style_drift = (
                        assigned_style is not None
                        and "style:" + assigned_style in index.records
                        and index.font_values("style:" + assigned_style, "color") != vis_color
                    )
                    mark_drift = (
                        read_font(index.document, index.elements[oid], None, "color") != vis_color
                    )
                    if style_drift or mark_drift:
                        result.append(
                            {
                                "id": f"semantic.{oid}.color",
                                "action": "font.color.set",
                                "target": target,
                                "params": {"value": vis_color},
                            }
                        )
        # Only create missing outline semantics. Explicit sample rules are already resolved.
        # Front-matter headings (cover, declaration, abstract, toc) must default to
        # outline level 9 (body text) so Word TOC fields never index them.
        if oid not in explicit_outline:
            result.append(
                {
                    "id": f"semantic.{oid}.outline",
                    "action": "paragraph.outline_level.set",
                    "target": target,
                    "params": {
                        "value": 9 if record.get("module") in front_matter_modules else level - 1
                    },
                }
            )
        if record.get("module") != "chapters":
            result.append(
                {
                    "id": f"semantic.{oid}.unnumbered",
                    "action": "paragraph.numbering_suppressed.set",
                    "target": target,
                    "params": {"value": True},
                }
            )
    return result


def bind_template(index, template, mapping=None):
    validate_template(template)
    mapping = mapping or {}
    require_keys(mapping, {"rules", "sections", "preserve"}, label="template mapping")
    overrides, sections = mapping.get("rules", {}), mapping.get("sections", {})
    if not isinstance(overrides, dict) or not isinstance(sections, dict):
        raise FormatError("Template rule/section mappings must be objects")
    if set(overrides) - {r["id"] for r in template["rules"]}:
        raise FormatError("Unknown template rule mapping")
    for target in overrides.values():
        if not isinstance(target, dict):
            raise FormatError("Rule target mappings must be objects")
    for key, value in sections.items():
        if not isinstance(key, str) or not re.fullmatch(r"(?:s[1-9]\d*:)?[1-9]\d*", key):
            raise FormatError("Section mapping keys must be a section number or sSOURCE:SECTION")
        number(value, 1, len(index.document.sections), "mapped section", True)
    preserve = mapping.get("preserve", {})
    if not isinstance(preserve, dict):
        raise FormatError("mapping.preserve must map paragraph IDs to reasons")
    for oid, reason in preserve.items():
        if index.records.get(oid, {}).get("kind") != "paragraph":
            raise FormatError("Unknown preserved target paragraph ID")
        text_value(reason, "preserved paragraph reason")
    operations, ledger = [], []
    pending = copy.deepcopy(template["issues"])
    if "format" in template or any(s["kind"] == "template" for s in template["sources"]):
        for record in index.records.values():
            if (
                record["kind"] == "paragraph"
                and record.get("unsupported_container")
                and record.get("text", "").strip()
                and record["id"] not in preserve
            ):
                ledger.append(
                    {
                        "rule_id": f"protected.{record['id']}",
                        "state": "preserved",
                        "target_ids": [record["id"]],
                        "reason": "Complex container preserved; automatic formatting is not supported here",
                    }
                )
    if "format" in template:
        from .spec import spec_request

        expanded = spec_request(index, template["format"], isolated=True)
        for n, op in enumerate(expanded["operations"], 1):
            if preserve and op["target"].get("type") == "paragraph" and "ids" in op["target"]:
                kept = [oid for oid in op["target"]["ids"] if oid not in preserve]
                if not kept:
                    continue
                op = {**op, "target": {**op["target"], "ids": kept}}
            elif preserve and op["target"].get("id") in preserve:
                continue
            operation = {**op, "id": f"builtin.{n}"}
            operations.append(operation)
            ledger.append(
                {"rule_id": operation["id"], "state": "planned", "evidence": "builtin format"}
            )
    for rule in template["rules"]:
        target = copy.deepcopy(overrides.get(rule["id"], rule["target"]))
        if rule["state"] == "preserve" or set(target) == {"preserve"}:
            reason = target.get("preserve", rule.get("reason"))
            text_value(reason, "preserved rule reason")
            ledger.append({"rule_id": rule["id"], "state": "preserved", "reason": reason})
            continue
        if target.get("type") == "section" and "number" in target:
            source_number = rule["evidence"][0].get("source", 1)
            key = f"s{source_number}:{target['number']}"
            snapshots = template.get("evidence", [])
            multi_section = any(
                len(e.get("snapshot", {}).get("sections", [])) > 1 for e in snapshots
            )
            if key in sections:
                target["number"] = sections[key]
            elif str(target["number"]) in sections:
                target["number"] = sections[str(target["number"])]
            elif len(index.document.sections) > 1 or multi_section:
                pending.append(
                    {"id": rule["id"], "reason": "Map sample sections explicitly", "target": target}
                )
                continue
        cap = REGISTRY[rule["action"]]
        try:
            refs = resolve_targets(index, target, cap.targets)
        except FormatError as exc:
            if rule.get("required", True):
                pending.append({"id": rule["id"], "reason": str(exc), "target": target})
            else:
                ledger.append(
                    {"rule_id": rule["id"], "state": "not_applicable", "reason": str(exc)}
                )
            continue
        if preserve and target.get("type") in {"paragraph", "text_range"}:
            kept_refs = [r for r in refs if r["id"] not in preserve]
            if len(kept_refs) != len(refs):
                if not kept_refs:
                    ledger.append(
                        {
                            "rule_id": rule["id"],
                            "state": "not_applicable",
                            "reason": "All matching paragraphs are explicitly preserved",
                        }
                    )
                    continue
                if target.get("type") == "paragraph":
                    target = {"type": "paragraph", "ids": [r["id"] for r in kept_refs]}
                    refs = kept_refs
        if (
            any(e.get("confidence", 1) < 1 for e in rule["evidence"])
            and rule["id"] not in overrides
        ):
            pending.append(
                {"id": rule["id"], "reason": "Low-confidence evidence requires explicit mapping"}
            )
            continue
        operations.append(
            {
                "id": rule["id"],
                "action": rule["action"],
                "target": target,
                "params": cap.validator(rule["params"]),
            }
        )
        ledger.append(
            {
                "rule_id": rule["id"],
                "state": "planned",
                "target_ids": [r["id"] for r in refs],
                "evidence": rule["evidence"],
            }
        )
    for oid, reason in preserve.items():
        ledger.append(
            {
                "rule_id": f"preserve.{oid}",
                "state": "preserved",
                "target_ids": [oid],
                "reason": reason,
            }
        )
    if any(s["kind"] == "template" for s in template["sources"]):
        covered = {oid for entry in ledger for oid in entry.get("target_ids", [])}
        toc_enabled = any(
            op["action"] == "toc.configure" and op["params"].get("enabled") for op in operations
        )
        uncovered = [
            r["id"]
            for r in index.records.values()
            if r["kind"] == "paragraph"
            and r["part"] == "word/document.xml"
            and r.get("text", "").strip()
            and r["id"] not in covered
            and not (toc_enabled and r.get("semantic_role") == "toc_heading")
        ]
        if uncovered:
            pending.append(
                {
                    "id": "target-coverage",
                    "reason": "Target paragraphs have no sample rule; map explicitly or record mapping.preserve",
                    "target_ids": uncovered,
                }
            )
    semantics = _heading_structure(index, operations, preserve)
    operations = semantics + operations
    operations = [
        {**op, "params": REGISTRY[op["action"]].validator(op["params"])} for op in operations
    ]
    toc_ops = [op for op in operations if op["action"] == "toc.configure"]
    if toc_ops:
        operations = [op for op in operations if op["action"] != "toc.configure"] + toc_ops
    ledger = [
        {"rule_id": op["id"], "state": "planned", "evidence": "Word heading semantics"}
        for op in semantics
    ] + ledger
    return operations, ledger, pending


def template_plan(input_path, destination, template, *, mapping=None, template_source=None):
    from .api import build_plan

    document = Document(source_path(input_path))
    operations, coverage, pending = bind_template(DocumentIndex(document), template, mapping)
    if pending:
        return {
            "kind": "word-format-template-review",
            "status": "needs_clarification",
            "issues": pending,
            "coverage": coverage,
            "template_id": template["id"],
        }
    context = {
        "definition": template,
        "sha256": json_hash(template),
        "mapping": mapping or {},
        "coverage": coverage,
        "source": template_source,
    }
    return build_plan(
        input_path,
        destination,
        {"operations": operations, "template_context": context},
        _document=document,
    )


def validate_context(index, context, operations, destination):
    require_keys(
        context,
        {"definition", "sha256", "mapping", "coverage", "source"},
        {"definition", "sha256", "mapping", "coverage", "source"},
        "template context",
    )
    if json_hash(context["definition"]) != context["sha256"]:
        raise FormatError("Template definition changed; compile a new plan")
    source = context["source"]
    protected = list(context["definition"]["sources"])
    if source:
        if sha256(source["path"]) != source["sha256"]:
            raise FormatError("Template file changed; compile a new plan")
        protected.append(source)
        if (
            json_hash(json.loads(Path(source["path"]).read_text(encoding="utf-8-sig")))
            != context["sha256"]
        ):
            raise FormatError("Plan definition does not match the selected template file")
    for item in protected:
        path = Path(item["path"]).resolve()
        if destination == path or (destination.exists() and destination.samefile(path)):
            raise FormatError("Output must not overwrite a template or evidence source")
    expected, coverage, pending = bind_template(index, context["definition"], context["mapping"])
    if pending or expected != operations or coverage != context["coverage"]:
        raise FormatError("Template coverage is incomplete or changed; rebuild the plan")


def compliance(context, *, format_status, visual_status):
    partial = bool(context["definition"].get("limitations")) or any(
        entry["state"] == "preserved" for entry in context["coverage"]
    )
    status = (
        "failed"
        if format_status == "failed" or visual_status == "failed"
        else ("partial" if partial else "passed" if visual_status == "passed" else "pending")
    )
    return {
        "status": status,
        "template_id": context["definition"]["id"],
        "template_sha256": context["sha256"],
        "format": format_status,
        "visual": visual_status,
        "coverage": context["coverage"],
        "limitations": context["definition"].get("limitations", []),
    }
