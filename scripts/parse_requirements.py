#!/usr/bin/env python3
"""Parse common Chinese and English Word formatting requirements into JSON."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


CN_SIZE_PT = {
    "初号": 42,
    "小初号": 36,
    "小初": 36,
    "一号": 26,
    "小一号": 24,
    "小一": 24,
    "二号": 22,
    "小二号": 18,
    "小二": 18,
    "三号": 16,
    "小三号": 15,
    "小三": 15,
    "四号": 14,
    "小四号": 12,
    "小四": 12,
    "五号": 10.5,
    "小五号": 9,
    "小五": 9,
    "六号": 7.5,
    "小六号": 6.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

FONT_NAMES = [
    "Times New Roman",
    "Aptos",
    "Calibri",
    "Arial",
    "Georgia",
    "Cambria",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "微软雅黑",
    "方正小标宋",
]


def measure_to_mm(value: float, unit: str) -> float:
    unit = unit.lower()
    if unit in {"cm", "厘米"}:
        return round(value * 10, 3)
    if unit in {"mm", "毫米"}:
        return round(value, 3)
    if unit in {"in", "inch", "inches", "英寸"}:
        return round(value * 25.4, 3)
    if unit in {"pt", "磅"}:
        return round(value * 25.4 / 72, 3)
    raise ValueError(f"Unsupported unit: {unit}")


def _set_nested(target: dict[str, Any], path: str, value: Any) -> None:
    cursor = target
    parts = path.split(".")
    for part in parts[:-1]:
        cursor = cursor.setdefault(part, {})
    cursor[parts[-1]] = value


def _number(value: str) -> float:
    return float(value.replace(",", "."))


def _size_from_text(text: str) -> float | None:
    for label, points in sorted(CN_SIZE_PT.items(), key=lambda x: len(x[0]), reverse=True):
        if label in text:
            return points
    match = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:pt|磅|points?)", text, re.I)
    return _number(match.group(1)) if match else None


def parse_requirements(text: str) -> dict[str, Any]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    segments = [
        part.strip()
        for part in re.split(r"[\n。；;]+", normalized)
        if part.strip()
    ]
    spec: dict[str, Any] = {
        "schema_version": "1.0",
        "id": "detected-requirements",
        "name": "Detected requirements",
        "mode": "parameterized",
        "template_required": False,
        "page": {},
        "body": {},
        "headings": [],
        "captions": {},
        "headers_footers": {},
        "page_numbers": {},
        "table_of_contents": {},
        "references": {},
        "notes": [],
    }
    evidence: list[dict[str, Any]] = []
    consumed: set[str] = set()

    def record(field: str, value: Any, match_text: str, confidence: float = 0.95) -> None:
        _set_nested(spec, field, value)
        cleaned = " ".join(match_text.split())
        evidence.append(
            {
                "field": field,
                "value": value,
                "source": "pasted requirements",
                "confidence": confidence,
                "evidence": cleaned,
            }
        )
        consumed.add(cleaned)

    if re.search(r"(?<![A-Za-z0-9])A4(?![A-Za-z0-9])|A4\s*纸", normalized, re.I):
        record("page.size", "A4", "A4", 0.99)
        record("page.width_mm", 210, "A4", 0.99)
        record("page.height_mm", 297, "A4", 0.99)
    elif re.search(r"\b(?:US\s*)?Letter\b|信纸", normalized, re.I):
        record("page.size", "Letter", "Letter", 0.99)
        record("page.width_mm", 215.9, "Letter", 0.99)
        record("page.height_mm", 279.4, "Letter", 0.99)

    if re.search(r"横向|landscape", normalized, re.I):
        record("page.orientation", "landscape", "landscape / 横向", 0.98)
    elif re.search(r"纵向|portrait", normalized, re.I):
        record("page.orientation", "portrait", "portrait / 纵向", 0.98)

    unit_pattern = r"(\d+(?:[.,]\d+)?)\s*(cm|mm|inches?|inch|in|pt|厘米|毫米|英寸|磅)"
    all_margin = re.search(
        rf"(?:四周|全部|所有)?\s*(?:页边距|margins?)\s*(?:均为|均|为|:|：)?\s*{unit_pattern}",
        normalized,
        re.I,
    )
    if all_margin:
        value = measure_to_mm(_number(all_margin.group(1)), all_margin.group(2))
        for side in ("top", "right", "bottom", "left"):
            record(f"page.margin_{side}_mm", value, all_margin.group(0), 0.96)

    sides = {
        "top": r"上(?:页?边距)?|top(?:\s+margin)?",
        "right": r"右(?:页?边距)?|right(?:\s+margin)?",
        "bottom": r"下(?:页?边距)?|bottom(?:\s+margin)?",
        "left": r"左(?:页?边距)?|left(?:\s+margin)?",
    }
    for side, label in sides.items():
        match = re.search(rf"(?:{label})\s*(?:为|:|：)?\s*{unit_pattern}", normalized, re.I)
        if match:
            value = measure_to_mm(_number(match.group(1)), match.group(2))
            record(f"page.margin_{side}_mm", value, match.group(0), 0.98)

    body_lines = [
        segment for segment in segments
        if re.search(r"正文|body(?:\s+text)?|normal", segment, re.I)
    ]
    generic_body_lines = [
        segment for segment in segments
        if not re.search(
            r"[一二三四五]级标题|heading\s*[1-5]|图题|图注|图标|figure\s+caption|"
            r"表题|表注|表标|table\s+caption|参考文献|bibliograph|references|页眉|页脚|页码|"
            r"目录|table\s+of\s+contents|\bTOC\b|header|footer|page\s+numbers?",
            segment,
            re.I,
        )
    ]
    body_context = "\n".join(body_lines or generic_body_lines)
    for font in FONT_NAMES:
        if re.search(re.escape(font), body_context, re.I):
            target = "body.font_east_asia" if re.search(r"[\u4e00-\u9fff]", font) else "body.font_latin"
            record(target, font, font, 0.94 if body_lines else 0.72)
    if "body.font_east_asia" in spec.get("body", {}) and "body.font_latin" not in spec.get("body", {}):
        if re.search(r"Times\s*New\s*Roman|英文|西文", body_context, re.I):
            record("body.font_latin", "Times New Roman", "Times New Roman", 0.92)

    body_size = _size_from_text(body_context)
    if body_size is not None:
        match = re.search(r"(?:正文|body|normal)[^\n。；;]{0,60}", body_context, re.I)
        record("body.font_size_pt", body_size, match.group(0) if match else str(body_size), 0.95 if body_lines else 0.70)

    multiple = re.search(r"(\d+(?:[.,]\d+)?)\s*倍(?:行距)?", body_context)
    if multiple:
        record(
            "body.line_spacing",
            {"kind": "multiple", "value": _number(multiple.group(1))},
            multiple.group(0),
            0.98,
        )
    elif re.search(r"双倍行距|double[- ]?spaced?|double\s+spacing", body_context, re.I):
        record("body.line_spacing", {"kind": "multiple", "value": 2.0}, "双倍行距 / double spacing", 0.99)
    elif re.search(r"单倍行距|single[- ]?spaced?|single\s+spacing", body_context, re.I):
        record("body.line_spacing", {"kind": "multiple", "value": 1.0}, "单倍行距 / single spacing", 0.99)
    else:
        exact = re.search(r"(?:固定值|exact(?:ly)?)\s*[，,:\s]?\s*(\d+(?:[.,]\d+)?)\s*(?:pt|磅)", body_context, re.I)
        if exact:
            record("body.line_spacing", {"kind": "exact", "value_pt": _number(exact.group(1))}, exact.group(0), 0.98)

    indent_chars = re.search(
        r"(?:首行缩进|first[- ]?line\s+indent)\s*(\d+(?:[.,]\d+)?)\s*(?:个)?\s*(?:字符|chars?)",
        body_context,
        re.I,
    )
    indent_measure = re.search(
        r"(?:首行缩进|first[- ]?line\s+indent)\s*(\d+(?:[.,]\d+)?)\s*(cm|mm|inches?|inch|in|pt|厘米|毫米|英寸|磅)",
        body_context,
        re.I,
    )
    if indent_chars:
        amount = _number(indent_chars.group(1))
        record("body.first_line_indent_chars", amount, indent_chars.group(0), 0.95)
    elif indent_measure:
        amount = _number(indent_measure.group(1))
        value = measure_to_mm(amount, indent_measure.group(2))
        record("body.first_line_indent_mm", value, indent_measure.group(0), 0.98)

    alignments = [
        ("justify", r"两端对齐|justif(?:y|ied)"),
        ("center", r"居中|cent(?:er|red)"),
        ("left", r"左对齐|left[- ]?align"),
        ("right", r"右对齐|right[- ]?align"),
    ]
    for value, pattern in alignments:
        match = re.search(rf"(?:正文|body)[^\n。；;]{{0,50}}(?:{pattern})", normalized, re.I)
        if match:
            record("body.alignment", value, match.group(0), 0.91)
            break

    heading_results: dict[int, dict[str, Any]] = {}
    heading_patterns = [
        (3, r"(?:二级节标题|小节标题|三级标题|heading\s*3)"),
        (2, r"(?:一级节标题|二级标题|(?<!二级)节标题|heading\s*2)"),
        (1, r"(?:章标题|章名|一级标题|第一章|heading\s*1)"),
    ]
    for line in segments:
        if re.search(r"目录|table\s+of\s+contents|\bTOC\b", line, re.I):
            continue
        level = None
        for cand_level, pat in heading_patterns:
            if re.search(pat, line, re.I):
                level = cand_level
                break
        if level is None:
            level_match = re.search(r"([一二三四五1-5])级标题|heading\s*([1-5])", line, re.I)
            if level_match:
                raw_l = level_match.group(1) or level_match.group(2)
                level = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5}.get(raw_l, int(raw_l) if raw_l.isdigit() else 1)
        if level is None:
            continue
        heading = heading_results.get(level, {"level": level})
        size = _size_from_text(line)
        if size is not None:
            heading["font_size_pt"] = size
            evidence.append({"field": f"headings[{level}].font_size_pt", "value": size, "source": "pasted requirements", "confidence": 0.95, "evidence": line.strip()})
        for font in FONT_NAMES:
            if re.search(re.escape(font), line, re.I):
                heading["font_east_asia" if re.search(r"[\u4e00-\u9fff]", font) else "font_latin"] = font
        if re.search(r"加粗|粗体|bold", line, re.I):
            heading["bold"] = True
        if re.search(r"居中|center", line, re.I):
            heading["alignment"] = "center"
        elif re.search(r"左对齐|居左|left", line, re.I):
            heading["alignment"] = "left"
        heading_multiple = re.search(r"(\d+(?:[.,]\d+)?)\s*倍(?:行距)?", line)
        if heading_multiple:
            heading["line_spacing"] = {"kind": "multiple", "value": _number(heading_multiple.group(1))}
        heading_results[level] = heading
        consumed.add(" ".join(line.split()))
    spec["headings"] = [heading_results[level] for level in sorted(heading_results)]

    # Recognize supported citation names in user requirements only. This parser
    # does not register, download, or imply an embedded publisher template.
    citation_patterns = [
        ("GB/T 7714-2025", r"GB\s*/?\s*T\s*7714[-—]?2025"),
        ("GB/T 7714-2015", r"GB\s*/?\s*T\s*7714[-—]?2015"),
        ("GB/T 7714-2005", r"GB\s*/?\s*T\s*7714[-—]?2005"),
        ("APA 7", r"APA\s*(?:7|第七版|7th)"),
        ("MLA 9", r"MLA\s*(?:9|第九版|9th)"),
    ]
    for system, pattern in citation_patterns:
        match = re.search(pattern, normalized, re.I)
        if match:
            record("references.citation_system", system, match.group(0), 0.99)
            break

    caption_roles = {
        "figure": r"图序|图名|图题|图注|图标|插图|figure\s+caption",
        "table": r"表序|表名|表题|表注|表标|表格|table\s+caption",
    }
    for role, pattern in caption_roles.items():
        for segment in segments:
            if not re.search(pattern, segment, re.I):
                continue
            label_default = "图" if role == "figure" and re.search(r"[\u4e00-\u9fff]", segment) else "Figure"
            if role == "table":
                label_default = "表" if re.search(r"[\u4e00-\u9fff]", segment) else "Table"
            record(f"captions.{role}.label", label_default, segment, 0.88)
            size = _size_from_text(segment)
            if size is not None:
                record(f"captions.{role}.font_size_pt", size, segment, 0.95)
            for font in FONT_NAMES:
                if re.search(re.escape(font), segment, re.I):
                    target = "font_east_asia" if re.search(r"[\u4e00-\u9fff]", font) else "font_latin"
                    record(f"captions.{role}.{target}", font, segment, 0.92)
                    break
            if re.search(r"居中|center", segment, re.I):
                record(f"captions.{role}.alignment", "center", segment, 0.96)
            elif re.search(r"左对齐|居左|left", segment, re.I):
                record(f"captions.{role}.alignment", "left", segment, 0.96)
            elif re.search(r"右对齐|居右|right", segment, re.I):
                record(f"captions.{role}.alignment", "right", segment, 0.96)
            above = re.search(r"上方|图上|表上|above", segment, re.I)
            below = re.search(r"下方|图下|表下|below", segment, re.I)
            if above or below:
                record(f"captions.{role}.position", "above" if above else "below", segment, 0.97)
            else:
                record(f"captions.{role}.position", "above" if role == "table" else "below", segment, 0.85)
            if re.search(r"自动编号|连续编号|分章编号|SEQ", segment, re.I):
                record(f"captions.{role}.numbering_mode", "seq", segment, 0.98)

    for kind, pattern in (("header", r"页眉|header"), ("footer", r"页脚|footer")):
        for segment in segments:
            if not re.search(pattern, segment, re.I):
                continue
            distance = re.search(rf"(?:{pattern})(?:距|距离)?\s*(?:页面)?(?:顶端|底端|边界)?\s*(?:为|:|：)?\s*{unit_pattern}", segment, re.I)
            if distance:
                value = measure_to_mm(_number(distance.group(1)), distance.group(2))
                record(f"page.{kind}_distance_mm", value, distance.group(0), 0.96)
                continue
            text_match = re.search(rf"(?:{pattern})\s*(?:文字|内容)?\s*[:：]\s*(.+)$", segment, re.I)
            if text_match and text_match.group(1).strip():
                record(f"headers_footers.{kind}.enabled", True, segment, 0.96)
                record(f"headers_footers.{kind}.text", text_match.group(1).strip(), segment, 0.91)
            for alignment, alignment_pattern in alignments:
                if re.search(alignment_pattern, segment, re.I):
                    record(f"headers_footers.{kind}.alignment", alignment, segment, 0.94)
                    break
    if re.search(r"首页不同|different\s+first\s+page", normalized, re.I):
        record("headers_footers.different_first_page", True, "首页不同 / different first page", 0.98)

    page_number_segment = next(
        (segment for segment in segments if re.search(r"页码|page\s+numbers?", segment, re.I)),
        None,
    )
    if page_number_segment:
        first_page_exclusion = bool(re.search(r"首页(?:不显示|无)页码|no\s+number\s+on\s+first", normalized, re.I))
        disabled = bool(re.search(r"^(?:全文|整篇)?\s*(?:无页码|不显示页码)|^no\s+page\s+numbers?", page_number_segment, re.I))
        enabled = not disabled
        record("page_numbers.enabled", enabled, page_number_segment, 0.98)
        if enabled:
            location = "header" if re.search(r"页眉|顶部|header|top", page_number_segment, re.I) else "footer"
            record("page_numbers.location", location, page_number_segment, 0.90)
            for alignment, alignment_pattern in alignments:
                if re.search(alignment_pattern, page_number_segment, re.I):
                    record("page_numbers.alignment", alignment, page_number_segment, 0.96)
                    break
            if re.search(r"共\s*\d*\s*页|of\s+(?:total|\d+)|总页数", page_number_segment, re.I):
                record("page_numbers.format", "page-x-of-y", page_number_segment, 0.94)
            elif re.search(r"第\s*\d*\s*页|Page\s*\d+", page_number_segment, re.I):
                record("page_numbers.format", "page-number", page_number_segment, 0.92)
            else:
                record("page_numbers.format", "number", page_number_segment, 0.90)
            start_match = re.search(r"(?:起始|从|start(?:ing)?(?:\s+at)?)\s*(\d+)", page_number_segment, re.I)
            if start_match:
                record("page_numbers.start", int(start_match.group(1)), page_number_segment, 0.95)
            record(
                "page_numbers.show_on_first_page",
                not first_page_exclusion,
                page_number_segment,
                0.93,
            )

    toc_segments = [
        segment for segment in segments
        if re.search(r"目录|table\s+of\s+contents|\bTOC\b", segment, re.I)
    ]
    chinese_levels = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    for toc_segment in toc_segments:
        disabled = bool(re.search(r"(?:不插入|不生成|无|不要)目录|no\s+(?:table\s+of\s+contents|toc)", toc_segment, re.I))
        record("table_of_contents.enabled", not disabled, toc_segment, 0.97)
        if disabled:
            continue
        level_match = re.search(r"(?:到|至|含|包含到)?\s*([1-9一二三四五六七八九])\s*级(?:标题)?|heading\s*([1-9])", toc_segment, re.I)
        if level_match:
            token = level_match.group(1) or level_match.group(2)
            level = chinese_levels[token] if token in chinese_levels else int(token)
            record("table_of_contents.max_heading_level", level, toc_segment, 0.94)
        if re.search(r"目录后(?:分页|另起一页)|page\s+break\s+after", toc_segment, re.I):
            record("table_of_contents.page_break_after", True, toc_segment, 0.92)
        elif re.search(r"目录后不分页|no\s+page\s+break", toc_segment, re.I):
            record("table_of_contents.page_break_after", False, toc_segment, 0.92)

    reference_segments = [segment for segment in segments if re.search(r"参考文献|bibliograph|references", segment, re.I)]
    for segment in reference_segments:
        size = _size_from_text(segment)
        if size is not None:
            record("references.font_size_pt", size, segment, 0.93)
        for font in FONT_NAMES:
            if re.search(re.escape(font), segment, re.I):
                target = "font_east_asia" if re.search(r"[\u4e00-\u9fff]", font) else "font_latin"
                record(f"references.{target}", font, segment, 0.92)
        hanging = re.search(rf"(?:悬挂缩进|hanging\s+indent)\s*(?:为|:|：)?\s*{unit_pattern}", segment, re.I)
        if hanging:
            value = measure_to_mm(_number(hanging.group(1)), hanging.group(2))
            record("references.hanging_indent_mm", value, hanging.group(0), 0.98)
        ref_multiple = re.search(r"(\d+(?:[.,]\d+)?)\s*倍(?:行距)?", segment)
        if ref_multiple:
            record("references.line_spacing", {"kind": "multiple", "value": _number(ref_multiple.group(1))}, segment, 0.97)
        elif re.search(r"单倍行距|single\s+spacing", segment, re.I):
            record("references.line_spacing", {"kind": "multiple", "value": 1.0}, segment, 0.98)
        for alignment, alignment_pattern in alignments:
            if re.search(alignment_pattern, segment, re.I):
                record("references.alignment", alignment, segment, 0.94)
                break
        if re.search(r"自动编号|编号项|numbered\s+list", segment, re.I):
            record("references.numbering_mode", "word-numbering", segment, 0.97)

    list_segments = [segment for segment in segments if re.search(r"列表|编号|标号|小节以下层次|正文层次", segment, re.I)]
    for segment in list_segments:
        if re.search(r"（[1-9一二]）|\([1-9]\)|括号为序", segment):
            style = "decimal-fullwidth-parentheses" if "（" in segment or "括号" in segment else "decimal-parentheses"
            record("lists.numbered.style", style, segment, 0.92)
            record("lists.numbered.start", 1, segment, 0.92)
            break
        elif re.search(r"[1-9]\.|阿拉伯数字", segment):
            record("lists.numbered.style", "decimal-period", segment, 0.92)
            record("lists.numbered.start", 1, segment, 0.92)
            break
        elif re.search(r"[一二三四五]、", segment):
            record("lists.numbered.style", "chinese-period", segment, 0.92)
            record("lists.numbered.start", 1, segment, 0.92)

    clauses = [" ".join(part.split()) for part in segments]
    matched_evidence = {item["evidence"] for item in evidence}
    unresolved = [
        clause for clause in clauses
        if len(clause) >= 4
        and not any(clause in item or item in clause for item in matched_evidence)
    ][:50]

    return {
        "schema_version": "1.0",
        "route": "interpret-requirements",
        "spec_patch": spec,
        "evidence": evidence,
        "unresolved": unresolved,
        "warnings": [
            "Only explicit supported fields were populated.",
            "Review unresolved clauses and confirm all inferred values before applying them.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", help="UTF-8 text file containing formatting requirements")
    parser.add_argument("--text", help="Requirements supplied directly on the command line")
    parser.add_argument("--output", help="Write JSON to this path instead of stdout")
    args = parser.parse_args()

    if bool(args.input) == bool(args.text):
        parser.error("provide exactly one of input or --text")
    text = args.text if args.text is not None else Path(args.input).read_text(encoding="utf-8-sig")
    result = parse_requirements(text)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(rendered + "\n", encoding="utf-8")
    else:
        sys.stdout.write(rendered + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
