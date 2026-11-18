"""Compare Word-computed results with the document's declared native semantics."""

from __future__ import annotations

import re
from collections import Counter

from .contracts import FormatError
from .namespaces import qn

FIELD_ERROR = re.compile(r"^\s*(?:Error[!:]|错误[！!：:]|錯誤[！!：:])", re.I)


def field_errors(evidence):
    """Scan Field.Result only; prose containing the word 'error' is ordinary text."""
    errors = []
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


def _number(value, fmt):
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


def expected_labels(document):
    from .numbering import effective_numbering, level_spec

    counters = {}
    labels = {}
    for p in document.element.body.iter(qn("w:p")):
        native = effective_numbering(document, p)
        if not native or native[2] is None:
            continue
        nid, ilvl, abstract, num = native
        levels = {int(n.get(qn("w:ilvl"))): level_spec(n) for n in abstract.findall(qn("w:lvl"))}
        overrides = {}
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


def inherited_field_codes(document):
    """Identify established missing targets/error caches without guessing replacements."""
    bookmarks = {
        node.get(qn("w:name"))
        for part in document.part.package.parts
        if hasattr(part, "element")
        for node in part.element.iter(qn("w:bookmarkStart"))
    }
    codes = Counter()
    for part in document.part.package.parts:
        if not hasattr(part, "element"):
            continue
        for p in part.element.iter(qn("w:p")):
            stack = []
            fields = []
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
                        fields.append(tuple(stack.pop()[:2]))
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


def _uses_theme_font(document, paragraph, run, channel):
    from .properties import rpr_sources

    for properties in rpr_sources(document, paragraph, run):
        fonts = properties.find(qn("w:rFonts")) if properties is not None else None
        if fonts is not None:
            if fonts.get(qn(f"w:{channel}Theme")):
                return True
            if fonts.get(qn(f"w:{channel}")):
                return False
    return False


def validate_office(document, evidence, *, baseline_document=None):
    """Run after Word updates fields, before binding render and visual evidence."""
    from .body_sections import body_start
    from .inspect import DocumentIndex, _style_paragraph
    from .properties import read_font, read_paragraph, text_runs

    index = DocumentIndex(document)
    labels = expected_labels(document)
    actual = evidence.get("paragraphs", [])
    errors = []
    warnings = []
    checked = 0
    required_fonts = set()
    unavailable = []
    theme_fonts = set()
    cursor = 0
    matches = {}
    checked_roles = {"body", "module_body", "heading", "reference", "list", "keywords", "caption"}
    for oid in index.top_ids:
        record = index.records[oid]
        is_toc_heading = record.get("semantic_role") == "toc_heading"
        if not is_toc_heading and (
            record.get("semantic_role") == "toc_entry"
            or record.get("role") == "toc"
        ):
            continue
        node = index.elements[oid]
        text = record.get("text", "").strip()
        if not text:
            continue
        found = next(
            (i for i in range(cursor, len(actual)) if actual[i]["text"].strip("\r\n\x07 ") == text),
            None,
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
            and record.get("role") in checked_roles
        )
        if should_check_format:
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
                    expected = values[0]
                    observed = row.get(name)
                    if name in {"font_east_asia", "font_latin"}:
                        channel = "eastAsia" if name == "font_east_asia" else "ascii"
                        if (
                            isinstance(observed, str)
                            and observed.startswith("+")
                            and all(_uses_theme_font(document, node, run, channel) for run in runs)
                        ):
                            theme_fonts.add((name, expected, observed))
                            continue
                        required_fonts.add(expected)
                        mismatch = observed != expected
                    else:
                        mismatch = not (
                            isinstance(observed, (int, float))
                            and abs(float(observed) - float(expected)) <= 0.05
                        )
                    if mismatch:
                        errors.append(
                            {
                                "paragraph": oid,
                                "property": name,
                                "expected": expected,
                                "actual": observed,
                            }
                        )
            alignment = {"left": 0, "center": 1, "right": 2, "justify": 3}.get(
                read_paragraph(document, node, "alignment")
            )
            if alignment is not None and row.get("alignment") is not None and row["alignment"] != alignment:
                errors.append(
                    {
                        "paragraph": oid,
                        "property": "alignment",
                        "expected": alignment,
                        "actual": row["alignment"],
                    }
                )
            if not is_toc_heading:
                indent = read_paragraph(document, node, "first_line_indent")
                if indent["unit"] == "chars":
                    observed = row.get("first_line_indent_chars")
                    expected = indent["value"]
                else:
                    observed = row.get("first_line_indent_pt")
                    expected = indent["value"] * 72 / 25.4
                if observed is not None and abs(observed - expected) > 0.25:
                    errors.append(
                        {
                            "paragraph": oid,
                            "property": "first_line_indent",
                            "expected": expected,
                            "actual": observed,
                        }
                    )
                spacing = read_paragraph(document, node, "line_spacing")
                observed_rule = row.get("line_spacing_rule")
                if observed_rule is not None:
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
                            }
                        )
                    elif rule == "multiple" and row.get("line_spacing_pt") is not None:
                        expected_pt = float(value) * 12.0
                        if abs(float(row["line_spacing_pt"]) - expected_pt) > 0.5:
                            errors.append(
                                {
                                    "paragraph": oid,
                                    "property": "line_spacing_pt",
                                    "expected": expected_pt,
                                    "actual": row["line_spacing_pt"],
                                }
                            )
    toc_style_Lookup = {}
    for st in document.styles:
        if st.type == 1:
            m = re.fullmatch(r"(?:toc|目录)\s*([1-9])", (st.name or "").strip(), re.I)
            if m:
                toc_style_Lookup[int(m.group(1))] = st.name
    for row in actual:
        m = re.fullmatch(r"(?:toc|目录)\s*([1-9])", str(row.get("style", "")).strip(), re.I)
        if not m or not row.get("text", "").strip("\r\n\x07 "):
            continue
        lvl = int(m.group(1))
        st_name = toc_style_Lookup.get(lvl)
        if not st_name:
            continue
        probe_p = _style_paragraph(st_name, document)
        for name in ("font_east_asia", "font_latin", "font_size_pt"):
            expected = read_font(document, probe_p, None, name)
            if expected is None or isinstance(expected, dict):
                continue
            observed = row.get(name)
            if name in {"font_east_asia", "font_latin"}:
                channel = "eastAsia" if name == "font_east_asia" else "ascii"
                if (
                    isinstance(observed, str)
                    and observed.startswith("+")
                    and _uses_theme_font(document, probe_p, None, channel)
                ):
                    theme_fonts.add((name, expected, observed))
                    continue
                required_fonts.add(expected)
                mismatch = observed != expected
            else:
                mismatch = not (
                    isinstance(observed, (int, float))
                    and abs(float(observed) - float(expected)) <= 0.05
                )
            if mismatch:
                errors.append(
                    {
                        "paragraph": f"toc:{lvl}:{row['text'].strip()[:24]}",
                        "property": name,
                        "expected": expected,
                        "actual": observed,
                    }
                )
    try:
        first = body_start(document)
        first_row = matches.get(first)
        section_index = sum(
            1
            for node in list(document.element.body)[: list(document.element.body).index(first)]
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
    inherited_codes = (
        inherited_field_codes(baseline_document) if baseline_document is not None else Counter()
    )
    failed_fields = field_errors(evidence)
    inherited = []
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
    if "available_fonts" in evidence:
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
