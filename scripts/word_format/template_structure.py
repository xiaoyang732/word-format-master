"""Read and compare a sample's ordered thesis parts without writing content."""

from __future__ import annotations

import copy
import re
from collections import Counter
from difflib import SequenceMatcher

from document_structure import classify_module, heading_level, normalized_style_name

from .contracts import FormatError
from .namespaces import qn
from .properties import outline_level, paragraph_style_name, read_paragraph


def _variant(kind, text):
    compact = re.sub(r"\s+", "", text).lower().lstrip("【[（(")
    if kind in {"abstract", "keywords"}:
        return "en" if compact.startswith(("abstract", "keyword", "indexterm", "英文", "外文")) else "zh"
    if kind == "declaration":
        if "声明与授权" in compact or "授权声明" in compact:
            return "combined"
        if "授权" in compact:
            return "authorization"
        if "诚信" in compact:
            return "integrity"
        if "原创" in compact or "独创" in compact:
            return "originality"
        # Unrecognized declaration forms remain distinct evidence.
        return compact
    return ""


def _items(document):
    """Include table/image front matter, but do not treat table cells as headings."""
    page_group = 0
    for position, block in enumerate(document.element.body):
        if block.tag == qn("w:sectPr"):
            continue
        paragraphs = [block] if block.tag == qn("w:p") else list(block.iter(qn("w:p")))
        for offset, paragraph in enumerate(paragraphs):
            # Text boxes and revision containers require page review; their text
            # must not create a second phantom module inside the owning paragraph.
            if any(a.tag in {qn("w:txbxContent"), qn("w:del")} for a in paragraph.iterancestors()):
                continue
            in_table = any(a.tag == qn("w:tbl") for a in paragraph.iterancestors())
            if not in_table and read_paragraph(document, paragraph, "page_break_before"):
                page_group += 1
            content_group = None
            for node in paragraph.iter():
                if any(a.tag in {qn("w:txbxContent"), qn("w:del")} for a in node.iterancestors()):
                    continue
                if not in_table and (
                    (node.tag == qn("w:br") and node.get(qn("w:type")) == "page")
                    or node.tag == qn("w:lastRenderedPageBreak")
                ):
                    page_group += 1
                if content_group is None and (
                    (node.tag == qn("w:t") and (node.text or "").strip())
                    or node.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")}
                ):
                    content_group = page_group
            text = "".join(
                node.text or ""
                for node in paragraph.iter(qn("w:t"))
                if not any(
                    a.tag in {qn("w:txbxContent"), qn("w:del")} for a in node.iterancestors()
                )
            ).strip()
            nontext = any(
                node.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")}
                for node in paragraph.iter()
            )
            if text or nontext:
                yield {
                    "text": text,
                    "anchor": f"word/document.xml#block:{position + 1}/p:{offset + 1}",
                    "style_name": paragraph_style_name(document, paragraph),
                    "outline_level": outline_level(document, paragraph),
                    "boundary": not in_table,
                    "nontext": nontext,
                    "page_group": content_group if content_group is not None else page_group,
                }
            sect = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
            if sect is not None and not in_table:
                kind = sect.find(qn("w:type"))
                if kind is None or kind.get(qn("w:val")) in {"nextPage", "oddPage", "evenPage"}:
                    page_group += 1


def extract_template_structure(document):
    """Keep occurrences, abstract languages and chapter hierarchy in sample order.

    This is a detected contract for AI to review against the complete sample,
    including rendered pages; it is not a claim to understand arbitrary artwork.
    """
    items = list(_items(document))
    modules = []
    for item in items:
        text, style = item["text"], item["style_name"]
        kind = classify_module(text, style, item["outline_level"]) if item["boundary"] else None
        if not modules and kind is None:
            # A title/table/image before a semantic section is front matter.
            # Plain unmarked text alone is not evidence for a cover.
            kind = "front_matter"
        if kind:
            modules.append(
                {
                    "kind": kind,
                    "variant": _variant(kind, text),
                    "heading_text": text,
                    "anchor": item["anchor"],
                    "outline_levels": [],
                    "page_group": item["page_group"],
                }
            )
        if modules and modules[-1]["kind"] == "chapters" and item["boundary"]:
            level = heading_level(text, style, item["outline_level"])
            if level:
                modules[-1]["outline_levels"].append(level)
    if modules and modules[0]["kind"] == "front_matter":
        first = items[0]
        separate_page = len(modules) > 1 and modules[1]["page_group"] > first["page_group"]
        is_cover = (
            separate_page
            or first["nontext"]
            or not first["boundary"]
            or bool(re.search(r"封面|论文题目|毕业论文|学位论文", first["text"]))
        )
        is_title = normalized_style_name(first["style_name"]) in {"title", "subtitle"}
        modules[0]["kind"] = "cover" if is_cover else "title_block" if is_title else "body"
        modules[0]["separate_page"] = separate_page
    counts = Counter()
    for module in modules:
        module.pop("page_group")
        key = (module["kind"], module["variant"])
        counts[key] += 1
        module["occurrence"] = counts[key]
    return {"schema_version": "1.0", "match": "exact", "ordered_modules": modules}


def _key(module, *, flexible_chapters=False):
    outline_levels = tuple(module.get("outline_levels", []))
    if module["kind"] == "chapters" and flexible_chapters:
        outline_levels = tuple(sorted(set(outline_levels)))
    return (
        module["kind"],
        module.get("variant", ""),
        outline_levels if module["kind"] == "chapters" else (),
        module.get("separate_page", False) if module["kind"] != "cover" else True,
    )


def _collapse_chapter_runs(modules, *, flexible_chapters=False):
    if not flexible_chapters:
        return list(modules)
    collapsed = []
    for module in modules:
        item = copy.deepcopy(module)
        if (
            item.get("kind") == "chapters"
            and collapsed
            and collapsed[-1].get("kind") == "chapters"
        ):
            combined = list(collapsed[-1].get("outline_levels", [])) + list(
                item.get("outline_levels", [])
            )
            collapsed[-1]["outline_levels"] = sorted(set(combined))
        else:
            if item.get("kind") == "chapters":
                item["outline_levels"] = sorted(set(item.get("outline_levels", [])))
            collapsed.append(item)
    return collapsed


def validate_template_structure(contract):
    if not isinstance(contract, dict) or contract.get("match") not in {"exact", "flexible_chapters"}:
        raise FormatError("Template structure requires match: exact or flexible_chapters")
    if "flexible_chapters" in contract and not isinstance(contract["flexible_chapters"], bool):
        raise FormatError("Template structure flexible_chapters must be boolean")
    if contract.get("schema_version") != "1.0":
        raise FormatError("Unsupported template structure version")
    modules = contract.get("ordered_modules")
    if not isinstance(modules, list):
        raise FormatError("Template ordered_modules must be an array")
    for module in modules:
        if not isinstance(module, dict) or not isinstance(module.get("kind"), str):
            raise FormatError("Template modules require a kind")
        if not isinstance(module.get("variant", ""), str):
            raise FormatError("Template module variant must be text")
        if "separate_page" in module and not isinstance(module["separate_page"], bool):
            raise FormatError("Template front matter separate_page must be boolean")
        levels = module.get("outline_levels", [])
        if not isinstance(levels, list) or any(
            isinstance(level, bool) or not isinstance(level, int) or not 1 <= level <= 9
            for level in levels
        ):
            raise FormatError("Template chapter outline_levels must contain levels 1-9")


def compare_template_structure(document, contract):
    validate_template_structure(contract)
    flexible = bool(
        contract.get("flexible_chapters") or contract.get("match") == "flexible_chapters"
    )
    expected = _collapse_chapter_runs(
        contract["ordered_modules"], flexible_chapters=flexible
    )
    actual_raw = extract_template_structure(document)["ordered_modules"]
    actual = _collapse_chapter_runs(actual_raw, flexible_chapters=flexible)
    expected_keys = [_key(m, flexible_chapters=flexible) for m in expected]
    actual_keys = [_key(m, flexible_chapters=flexible) for m in actual]
    differences = []
    matcher = SequenceMatcher(a=expected_keys, b=actual_keys, autojunk=False)
    for tag, a, b, c, d in matcher.get_opcodes():
        if tag != "equal":
            differences.append(
                {
                    "change": tag,
                    "expected_position": a + 1,
                    "actual_position": c + 1,
                    "missing": copy.deepcopy(expected[a:b]),
                    "unexpected": copy.deepcopy(actual[c:d]),
                }
            )
    passed = expected_keys == actual_keys
    result = {
        "status": "passed" if passed else "failed",
        "expected": copy.deepcopy(contract["ordered_modules"]),
        "actual": actual_raw,
        "differences": differences,
        "errors": []
        if passed
        else ["Thesis parts, languages, counts, order or chapter hierarchy differ from the sample"],
    }
    if not passed:
        from .remediation import diagnose_structure_remediation

        result["remediation"] = diagnose_structure_remediation(result)
    return result
