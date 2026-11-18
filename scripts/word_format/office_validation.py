"""Compare Word-computed results with the document's declared native semantics."""

from __future__ import annotations

import re
from collections import Counter
from typing import Any

from docx.enum.style import WD_STYLE_TYPE

from .constants import (
    FONT_SIZE_TOLERANCE_PT,
    INDENT_TOLERANCE,
    LINE_SPACING_TOLERANCE_PT,
    WORD_POINTS_PER_LINE,
    mm_to_pt,
    parse_toc_style_level,
)
from .contracts import FormatError
from .namespaces import qn

FIELD_ERROR = re.compile(r"^\s*(?:Error[!:]|错误[！!：:]|錯誤[！!：:])", re.I)
CHECKED_ROLES = frozenset(
    {"body", "module_body", "heading", "reference", "list", "keywords", "caption"}
)
WORD_ALIGNMENTS: dict[str, int] = {"left": 0, "center": 1, "right": 2, "justify": 3}
WD_UNDEFINED_NUMERIC = 9999999
TEXT_PREVIEW_MAX_CHARS = 40
_CROSS_REF_CODE_RE = re.compile(r"^\s*(?:REF|PAGEREF)\b", re.I)
_NUMERIC_CITATION_SWITCH_RE = re.compile(r"(?:^|\s)\\n(?:\s|$)")
_CHARFORMAT_SWITCH_RE = re.compile(r"\\\*\s*CHARFORMAT\b", re.I)


def _text_preview(text: str, limit: int = TEXT_PREVIEW_MAX_CHARS) -> str:
    clean = " ".join(str(text or "").split())
    if len(clean) <= limit:
        return clean
    return clean[: max(1, limit - 3)] + "..."


def _unprotected_cross_ref_codes(node: Any) -> list[str]:
    """Return body cross-reference REF/PAGEREF fields missing the \\* CHARFORMAT switch."""
    if node is None:
        return []
    from .fields import field_codes

    codes: list[str] = []
    for code in field_codes(node):
        if (
            _CROSS_REF_CODE_RE.match(code)
            and not _NUMERIC_CITATION_SWITCH_RE.search(code)
            and not _CHARFORMAT_SWITCH_RE.search(code)
        ):
            codes.append(code)
    return codes


def field_errors(evidence: dict[str, Any]) -> list[dict[str, Any]]:
    """Scan Field.Result only; prose containing the word 'error' is ordinary text."""
    errors: list[dict[str, Any]] = []
    for field in evidence.get("fields", []):
        if not re.match(r"^\s*(?:REF|PAGEREF|STYLEREF|TOC|SEQ|PAGE)\b", field["code"], re.I):
            continue
        if field.get("target_exists") is False or FIELD_ERROR.match(field["result"]):
            errors.append(field)
        elif (
            re.match(r"^\s*REF\s+WFMRef_\w+\b", field["code"], re.I)
            and re.search(r"\\n(?:\s|$)", field["code"])
            and field["result"].strip() != (field.get("target_label") or "").strip()
        ):
            errors.append(field)
    return errors


def _number(value: int, fmt: str | None) -> str | None:
    if fmt in {"decimal", "decimalZero"}:
        return str(value).zfill(2) if fmt == "decimalZero" else str(value)
    if fmt in {"upperRoman", "lowerRoman"} and 0 < value < 4000:
        result = ""
        for amount, symbol in (
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ):
            count, value = divmod(value, amount)
            result += symbol * count
        return result.lower() if fmt == "lowerRoman" else result
    if fmt in {"upperLetter", "lowerLetter"} and value > 0:
        result = ""
        while value:
            value, digit = divmod(value - 1, 26)
            result = chr(65 + digit) + result
        return result.lower() if fmt == "lowerLetter" else result
    if fmt in {"chineseCounting", "chineseCountingThousand"} and 0 < value < 100:
        digits = "零一二三四五六七八九"
        tens, units = divmod(value, 10)
        return ((digits[tens] if tens > 1 else "") + "十" if tens else "") + (
            digits[units] if units else ""
        )
    if fmt == "none":
        return ""
    return None


def expected_labels(document: Any) -> dict[Any, str | None]:
    from .numbering import effective_numbering, level_spec

    counters: dict[tuple[str, int], int] = {}
    labels: dict[Any, str | None] = {}
    for p in document.element.body.iter(qn("w:p")):
        native = effective_numbering(document, p)
        if not native or native[2] is None:
            continue
        nid, ilvl, abstract, num = native
        levels = {int(n.get(qn("w:ilvl"))): level_spec(n) for n in abstract.findall(qn("w:lvl"))}
        overrides: dict[int, int] = {}
        for override in num.findall(qn("w:lvlOverride")):
            level = int(override.get(qn("w:ilvl")))
            start = override.find(qn("w:startOverride"))
            if start is not None:
                overrides[level] = int(start.get(qn("w:val")))
            replacement = override.find(qn("w:lvl"))
            if replacement is not None:
                levels[level] = level_spec(replacement)
        config = levels.get(ilvl)
        if config is None:
            continue
        counters[nid, ilvl] = (
            counters.get((nid, ilvl), overrides.get(ilvl, config["start"]) - 1) + 1
        )
        for deeper, settings in levels.items():
            restart = settings.get("restart", deeper)
            if deeper > ilvl and restart and ilvl < restart:
                counters.pop((nid, deeper), None)
        label = config.get("text", "")
        supported = True
        for match in list(re.finditer(r"%([1-9])", label))[::-1]:
            level = int(match[1]) - 1
            ancestor = levels.get(level, {})
            value = counters.get((nid, level), overrides.get(level, ancestor.get("start", 1)))
            fmt = "decimal" if config.get("legal") else ancestor.get("format")
            formatted = _number(value, fmt)
            if formatted is None:
                supported = False
                break
            label = label[: match.start()] + formatted + label[match.end() :]
        labels[p] = label if supported else None
    return labels


def inherited_field_codes(document: Any) -> Counter[str]:
    """Identify established missing targets/error caches without guessing replacements."""
    bookmarks = {
        node.get(qn("w:name"))
        for part in document.part.package.parts
        if hasattr(part, "element")
        for node in part.element.iter(qn("w:bookmarkStart"))
    }
    codes: Counter[str] = Counter()
    for part in document.part.package.parts:
        if not hasattr(part, "element"):
            continue
        for p in part.element.iter(qn("w:p")):
            stack: list[list[Any]] = []
            fields: list[tuple[str, str]] = []
            for node in p.iter():
                if node.tag == qn("w:fldSimple"):
                    fields.append(
                        (
                            node.get(qn("w:instr"), ""),
                            "".join(item.text or "" for item in node.iter(qn("w:t"))),
                        )
                    )
                elif node.tag == qn("w:fldChar"):
                    kind = node.get(qn("w:fldCharType"))
                    if kind == "begin":
                        stack.append(["", "", False])
                    elif stack and kind == "separate":
                        stack[-1][2] = True
                    elif stack and kind == "end":
                        fields.append((stack[-1][0], stack[-1][1]))
                        stack.pop()
                elif stack and node.tag == qn("w:instrText"):
                    stack[-1][0] += node.text or ""
                elif stack and stack[-1][2] and node.tag == qn("w:t"):
                    stack[-1][1] += node.text or ""
            for code, result in fields:
                target = re.match(r'^\s*(?:REF|PAGEREF)\s+(?:"([^"]+)"|(\S+))', code, re.I)
                if (target and (target[1] or target[2]) not in bookmarks) or FIELD_ERROR.match(
                    result
                ):
                    codes[" ".join(code.split())] += 1
    return codes


def _uses_theme_font(document: Any, paragraph: Any, run: Any, channel: str) -> bool:
    from .properties import rpr_sources

    for properties in rpr_sources(document, paragraph, run):
        fonts = properties.find(qn("w:rFonts")) if properties is not None else None
        if fonts is not None:
            if fonts.get(qn(f"w:{channel}Theme")):
                return True
            if fonts.get(qn(f"w:{channel}")):
                return False
    return False


def _find_matching_paragraph(
    actual: list[dict[str, Any]],
    cursor: int,
    text: str,
    *,
    allow_distant: bool = True,
    window: int = 8,
) -> int | None:
    window_end = min(len(actual), cursor + window)
    for idx in range(cursor, window_end):
        if actual[idx]["text"].strip("\r\n\x07 ") == text:
            return idx
    if not allow_distant:
        return None
    remaining = [
        idx
        for idx in range(window_end, len(actual))
        if actual[idx]["text"].strip("\r\n\x07 ") == text
    ]
    return remaining[0] if len(remaining) == 1 else None


def _compare_font_value(
    paragraph_id: str,
    name: str,
    expected: Any,
    observed: Any,
    *,
    uses_theme: bool,
    required_fonts: set[str],
    theme_fonts: set[tuple[str, str, str]],
    errors: list[dict[str, Any]],
    text_preview: str | None = None,
    node: Any = None,
) -> None:
    if name in {"font_east_asia", "font_latin"}:
        if isinstance(observed, str) and observed.startswith("+") and uses_theme:
            theme_fonts.add((name, str(expected), observed))
            return
        required_fonts.add(str(expected))
        mismatch = observed != expected
    else:
        mismatch = not (
            isinstance(observed, (int, float))
            and abs(float(observed) - float(expected)) <= FONT_SIZE_TOLERANCE_PT
        )
    if mismatch:
        error: dict[str, Any] = {
            "paragraph": paragraph_id,
            "property": name,
            "expected": expected,
            "actual": observed,
        }
        if text_preview is not None:
            error["text_preview"] = text_preview
        unprotected_refs = _unprotected_cross_ref_codes(node)
        is_mixed_undefined = observed == WD_UNDEFINED_NUMERIC or (
            name in {"font_east_asia", "font_latin"} and observed == ""
        )
        if unprotected_refs:
            error["fields"] = unprotected_refs
            error["hint"] = (
                "Cross-reference REF/PAGEREF field lacks \\* CHARFORMAT; "
                "Word may copy the target bookmark font/size into this paragraph."
            )
        elif is_mixed_undefined:
            error["hint"] = (
                "Word reported mixed/undefined run formatting across runs in this paragraph."
            )
        errors.append(error)


def _check_paragraph_typography(
    document: Any,
    oid: str,
    node: Any,
    row: dict[str, Any],
    *,
    is_toc_heading: bool,
    required_fonts: set[str],
    theme_fonts: set[tuple[str, str, str]],
    errors: list[dict[str, Any]],
) -> None:
    from .properties import read_font, read_paragraph, text_runs

    preview = _text_preview(row.get("text", ""))
    vis_runs = [
        r
        for r in text_runs(node)
        if any(t.text and t.text.strip() for t in r.findall(qn("w:t")))
    ]
    runs = vis_runs or list(text_runs(node))
    for name in ("font_east_asia", "font_latin", "font_size_pt"):
        if name not in row:
            continue
        values = [read_font(document, node, run, name) for run in runs]
        values = [value for value in values if value is not None]
        if (
            values
            and all(value == values[0] for value in values)
            and not isinstance(values[0], dict)
        ):
            channel = "eastAsia" if name == "font_east_asia" else "ascii"
            uses_theme = name in {"font_east_asia", "font_latin"} and all(
                _uses_theme_font(document, node, run, channel) for run in runs
            )
            _compare_font_value(
                oid,
                name,
                values[0],
                row.get(name),
                uses_theme=uses_theme,
                required_fonts=required_fonts,
                theme_fonts=theme_fonts,
                errors=errors,
                text_preview=preview,
                node=node,
            )

    alignment = WORD_ALIGNMENTS.get(read_paragraph(document, node, "alignment"))
    if alignment is not None and row.get("alignment") is not None and row["alignment"] != alignment:
        errors.append(
            {
                "paragraph": oid,
                "property": "alignment",
                "expected": alignment,
                "actual": row["alignment"],
                "text_preview": preview,
            }
        )
    if is_toc_heading:
        return

    indent = read_paragraph(document, node, "first_line_indent")
    if indent["unit"] == "chars":
        observed_indent = row.get("first_line_indent_chars")
        expected_indent = indent["value"]
    else:
        observed_indent = row.get("first_line_indent_pt")
        expected_indent = mm_to_pt(indent["value"])
    if observed_indent is not None and abs(observed_indent - expected_indent) > INDENT_TOLERANCE:
        errors.append(
            {
                "paragraph": oid,
                "property": "first_line_indent",
                "expected": expected_indent,
                "actual": observed_indent,
                "text_preview": preview,
            }
        )

    spacing = read_paragraph(document, node, "line_spacing")
    observed_rule = row.get("line_spacing_rule")
    if observed_rule is None:
        return
    rule = spacing["kind"]
    value = spacing["value"]
    allowed = (
        {1: {0, 5}, 1.5: {1, 5}, 2: {2, 5}}.get(value, {5})
        if rule == "multiple"
        else {4}
        if rule == "exact"
        else {3}
    )
    if observed_rule not in allowed:
        errors.append(
            {
                "paragraph": oid,
                "property": "line_spacing_rule",
                "expected": sorted(allowed),
                "actual": observed_rule,
                "text_preview": preview,
            }
        )
    elif rule == "multiple" and row.get("line_spacing_pt") is not None:
        expected_pt = float(value) * WORD_POINTS_PER_LINE
        if abs(float(row["line_spacing_pt"]) - expected_pt) > LINE_SPACING_TOLERANCE_PT:
            errors.append(
                {
                    "paragraph": oid,
                    "property": "line_spacing_pt",
                    "expected": expected_pt,
                    "actual": row["line_spacing_pt"],
                    "text_preview": preview,
                }
            )


def _check_toc_entry_typography(
    document: Any,
    actual: list[dict[str, Any]],
    required_fonts: set[str],
    theme_fonts: set[tuple[str, str, str]],
    errors: list[dict[str, Any]],
) -> None:
    from .inspect import _style_paragraph
    from .properties import read_font

    toc_style_lookup: dict[int, str] = {}
    for st in document.styles:
        if st.type == WD_STYLE_TYPE.PARAGRAPH:
            lvl = parse_toc_style_level(st.name)
            if lvl is not None:
                toc_style_lookup[lvl] = st.name
    for row in actual:
        lvl = parse_toc_style_level(str(row.get("style", "")))
        if lvl is None or not row.get("text", "").strip("\r\n\x07 "):
            continue
        st_name = toc_style_lookup.get(lvl)
        if not st_name:
            continue
        probe_p = _style_paragraph(st_name, document)
        for name in ("font_east_asia", "font_latin", "font_size_pt"):
            expected = read_font(document, probe_p, None, name)
            if expected is None or isinstance(expected, dict):
                continue
            channel = "eastAsia" if name == "font_east_asia" else "ascii"
            uses_theme = name in {"font_east_asia", "font_latin"} and _uses_theme_font(
                document, probe_p, None, channel
            )
            _compare_font_value(
                f"toc:{lvl}:{row['text'].strip()[:24]}",
                name,
                expected,
                row.get(name),
                uses_theme=uses_theme,
                required_fonts=required_fonts,
                theme_fonts=theme_fonts,
                errors=errors,
            )


def _check_body_page_start(
    document: Any,
    matches: dict[Any, dict[str, Any]],
    evidence: dict[str, Any],
    errors: list[dict[str, Any]],
    warnings: list[dict[str, Any]],
) -> None:
    from .body_sections import body_start

    try:
        first = body_start(document)
        first_row = matches.get(first)
        body_children = list(document.element.body)
        section_index = sum(
            1
            for node in body_children[: body_children.index(first)]
            if node.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None
        )
        page_type = document.sections[section_index]._sectPr.find(qn("w:pgNumType"))
        start = page_type.get(qn("w:start")) if page_type is not None else None
        if start is not None and first_row:
            computed_sections = evidence.get("sections", [])
            offset = 0
            if section_index < len(computed_sections) and "physical_page" in first_row:
                offset = (
                    first_row["physical_page"] - computed_sections[section_index]["physical_page"]
                )
            expected_start = int(start) + offset
            if first_row["page"] != expected_start:
                errors.append(
                    {
                        "property": "body_page_start",
                        "expected": expected_start,
                        "actual": first_row["page"],
                    }
                )
    except FormatError:
        pass  # No chapters in this document, e.g. a standalone cover.
    except (ValueError, IndexError):
        warnings.append({"reason": "Body page start unavailable"})


def _check_fields_and_toc_pages(
    evidence: dict[str, Any],
    baseline_document: Any,
    errors: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    inherited_codes = (
        inherited_field_codes(baseline_document) if baseline_document is not None else Counter()
    )
    failed_fields = field_errors(evidence)
    inherited: list[dict[str, Any]] = []
    for field in failed_fields:
        code = " ".join(field["code"].split())
        if inherited_codes[code] > 0:
            inherited.append(field)
            inherited_codes[code] -= 1
        else:
            errors.append({"field": field})
    for entry in evidence.get("toc_entries", []):
        label = re.search(r"\t\s*([0-9]+|[ivxlcdm]+)\s*$", entry["text"], re.I)
        if label:
            expected = entry["page"]
            actual_label = label[1]
            if actual_label not in {
                str(expected),
                _number(expected, "lowerRoman"),
                _number(expected, "upperRoman"),
            }:
                errors.append({"property": "toc_target_page", "entry": entry, "expected": expected})
    return failed_fields, inherited


def _check_available_fonts(
    evidence: dict[str, Any],
    actual: list[dict[str, Any]],
    required_fonts: set[str],
    warnings: list[dict[str, Any]],
) -> list[str]:
    if "available_fonts" not in evidence:
        return []
    installed = {font.casefold() for font in evidence["available_fonts"]}
    observed_fonts = {
        row.get(name, "") for row in actual for name in ("font_east_asia", "font_latin")
    }
    unavailable = sorted(
        font
        for font in observed_fonts
        if font and not font.startswith("+") and font.casefold() not in installed
    )
    required_unavailable = sorted(
        font
        for font in required_fonts
        if not font.startswith("+") and font.casefold() not in installed
    )
    if required_unavailable:
        warnings.append(
            {
                "reason": "Requested fonts unavailable; substitution unresolved",
                "fonts": required_unavailable,
            }
        )
    return unavailable


def validate_office(
    document: Any,
    evidence: dict[str, Any],
    *,
    baseline_document: Any = None,
) -> dict[str, Any]:
    """Run after Word updates fields, before binding render and visual evidence."""
    from .inspect import DocumentIndex

    index = DocumentIndex(document)
    labels = expected_labels(document)
    actual: list[dict[str, Any]] = evidence.get("paragraphs", [])
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    checked = 0
    required_fonts: set[str] = set()
    theme_fonts: set[tuple[str, str, str]] = set()
    cursor = 0
    matches: dict[Any, dict[str, Any]] = {}
    remaining_doc_texts = Counter(
        index.records[oid].get("text", "").strip()
        for oid in index.top_ids
        if not (
            index.records[oid].get("semantic_role") != "toc_heading"
            and (
                index.records[oid].get("semantic_role") == "toc_entry"
                or index.records[oid].get("role") == "toc"
            )
        )
        and index.records[oid].get("text", "").strip()
    )

    for oid in index.top_ids:
        record = index.records[oid]
        is_toc_heading = record.get("semantic_role") == "toc_heading"
        if not is_toc_heading and (
            record.get("semantic_role") == "toc_entry" or record.get("role") == "toc"
        ):
            continue
        node = index.elements[oid]
        text = record.get("text", "").strip()
        if not text:
            continue
        has_later_duplicate = remaining_doc_texts[text] > 1
        remaining_doc_texts[text] -= 1
        found = _find_matching_paragraph(
            actual, cursor, text, allow_distant=not has_later_duplicate
        )
        if found is None:
            if node in labels:
                errors.append({"paragraph": oid, "text": text, "reason": "Word paragraph missing"})
            continue
        row = actual[found]
        matches[node] = row
        cursor = found + 1
        if node in labels:
            expected = labels[node]
            if expected is None:
                warnings.append({"paragraph": oid, "reason": "Number format not modeled"})
            else:
                checked += 1
                if str(row.get("list_string", "")).strip() != expected.strip():
                    errors.append(
                        {
                            "paragraph": oid,
                            "text": text,
                            "expected_label": expected,
                            "actual_label": row.get("list_string", ""),
                        }
                    )
        should_check_format = is_toc_heading or (
            record.get("module") != "cover"
            and record.get("object_kind") in {None, "text"}
            and record.get("role") in CHECKED_ROLES
        )
        if should_check_format:
            _check_paragraph_typography(
                document,
                oid,
                node,
                row,
                is_toc_heading=is_toc_heading,
                required_fonts=required_fonts,
                theme_fonts=theme_fonts,
                errors=errors,
            )

    _check_toc_entry_typography(document, actual, required_fonts, theme_fonts, errors)
    _check_body_page_start(document, matches, evidence, errors, warnings)
    failed_fields, inherited = _check_fields_and_toc_pages(evidence, baseline_document, errors)
    unavailable = _check_available_fonts(evidence, actual, required_fonts, warnings)

    if not evidence.get("pagination_stable"):
        errors.append({"reason": "Word pagination did not stabilize"})
    if not failed_fields and any(
        item.get("pass") == evidence.get("update_passes")
        for item in evidence.get("field_update_errors", [])
    ):
        errors.append(
            {
                "reason": "Word returned a field update error in its final pass",
                "updates": evidence["field_update_errors"],
            }
        )
    return {
        "status": "failed" if errors else "partial" if warnings else "passed",
        "numbered_paragraphs_checked": checked,
        "errors": errors,
        "warnings": warnings,
        "inherited_field_errors": inherited,
        "observed_unavailable_fonts": unavailable,
        "theme_font_aliases": [
            {"property": name, "package_font": package_font, "word_alias": alias}
            for name, package_font, alias in sorted(theme_fonts)
        ],
        "delivery_ready": not errors and not warnings and not inherited,
        "application": evidence["application"],
        "coverage": evidence.get("coverage", {}),
    }
