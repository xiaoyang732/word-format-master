#!/usr/bin/env python3
"""Inspect DOCX/DOTX package structure and infer supported format tokens."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


from config import MAX_PARTS, MAX_UNCOMPRESSED
from document_structure import classify_role, detect_module_spans, validate_module_order

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = f"{{{NS['w']}}}"
STRICT_NAMESPACE_MAP = {
    "http://purl.oclc.org/ooxml/wordprocessingml/main": NS["w"],
    "http://purl.oclc.org/ooxml/officeDocument/relationships": NS["r"],
    "http://purl.oclc.org/ooxml/officeDocument/math": NS["m"],
}

CONTENT_PART_PREFIXES = ("word/header", "word/footer")
CONTENT_PART_NAMES = {
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
}
TRACKED_CHANGE_TAGS = {
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "pPrChange",
    "rPrChange",
    "sectPrChange",
    "tblPrChange",
    "tblGridChange",
    "trPrChange",
    "tcPrChange",
}


def w_attr(node: ET.Element | None, name: str) -> str | None:
    return None if node is None else node.get(W + name)


def bool_value(node: ET.Element | None) -> bool | None:
    if node is None:
        return None
    value = w_attr(node, "val")
    return value not in {"0", "false", "off"}


def twips_to_mm(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return round(length_to_pt(value) * 25.4 / 72, 3)
    except (TypeError, ValueError):
        return None


def length_to_pt(value: str | None) -> float | None:
    """Convert OOXML twips or Strict OOXML universal measures to points."""
    if value is None:
        return None
    cleaned = str(value).strip().lower()
    units = {
        "mm": 72 / 25.4,
        "cm": 72 / 2.54,
        "in": 72,
        "pt": 1,
        "pc": 12,
        "pi": 12,
    }
    for suffix, factor in units.items():
        if cleaned.endswith(suffix):
            return float(cleaned[: -len(suffix)]) * factor
    return float(cleaned) / 20


def half_points(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        cleaned = str(value).strip().lower()
        if cleaned.endswith("pt"):
            return round(float(cleaned[:-2]), 2)
        return round(float(cleaned) / 2, 2)
    except (TypeError, ValueError):
        return None


def read_xml(package: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        root = ET.fromstring(package.read(name))
        normalize_strict_namespaces(root)
        return root
    except KeyError:
        return None
    except ET.ParseError as exc:
        raise ValueError(f"Malformed XML in {name}: {exc}") from exc


def normalize_strict_namespaces(root: ET.Element) -> None:
    """Normalize ISO Strict namespaces so one parser handles both OOXML profiles."""
    for element in root.iter():
        if element.tag.startswith("{"):
            namespace, local = element.tag[1:].split("}", 1)
            if namespace in STRICT_NAMESPACE_MAP:
                element.tag = f"{{{STRICT_NAMESPACE_MAP[namespace]}}}{local}"
        if element.attrib:
            normalized: dict[str, str] = {}
            changed = False
            for key, value in element.attrib.items():
                new_key = key
                if key.startswith("{"):
                    namespace, local = key[1:].split("}", 1)
                    if namespace in STRICT_NAMESPACE_MAP:
                        new_key = f"{{{STRICT_NAMESPACE_MAP[namespace]}}}{local}"
                        changed = True
                normalized[new_key] = value
            if changed:
                element.attrib.clear()
                element.attrib.update(normalized)


def detect_conformance(package: zipfile.ZipFile) -> str:
    try:
        raw = package.read("word/document.xml")
    except KeyError:
        return "unknown"
    if b"http://purl.oclc.org/ooxml/wordprocessingml/main" in raw:
        return "strict"
    if b"http://schemas.openxmlformats.org/wordprocessingml/2006/main" in raw:
        return "transitional"
    return "unknown"


def is_content_part(name: str) -> bool:
    return name in CONTENT_PART_NAMES or (
        name.endswith(".xml") and name.startswith(CONTENT_PART_PREFIXES)
    )


def has_element(roots: list[ET.Element], namespace: str, *local_names: str) -> bool:
    qnames = {f"{{{NS[namespace]}}}{local_name}" for local_name in local_names}
    return any(element.tag in qnames for root in roots for element in root.iter())


def package_inventory(path: Path) -> tuple[zipfile.ZipFile, list[dict[str, Any]]]:
    if path.suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("Only DOCX and DOTX files are supported")
    if not zipfile.is_zipfile(path):
        raise ValueError("Input is not a valid OOXML ZIP package")
    package = zipfile.ZipFile(path)
    infos = package.infolist()
    if len(infos) > MAX_PARTS:
        package.close()
        raise ValueError(f"Package has too many parts: {len(infos)}")
    total = sum(info.file_size for info in infos)
    if total > MAX_UNCOMPRESSED:
        package.close()
        raise ValueError(f"Package expands beyond {MAX_UNCOMPRESSED} bytes")
    inventory = [
        {
            "path": info.filename,
            "compressed_bytes": info.compress_size,
            "uncompressed_bytes": info.file_size,
        }
        for info in infos
    ]
    return package, inventory


def parse_run_properties(parent: ET.Element | None) -> dict[str, Any]:
    if parent is None:
        return {}
    fonts = parent.find("w:rFonts", NS)
    result: dict[str, Any] = {}
    if fonts is not None:
        font_values = {
            "ascii": w_attr(fonts, "ascii"),
            "hAnsi": w_attr(fonts, "hAnsi"),
            "eastAsia": w_attr(fonts, "eastAsia"),
            "cs": w_attr(fonts, "cs"),
            "asciiTheme": w_attr(fonts, "asciiTheme"),
            "eastAsiaTheme": w_attr(fonts, "eastAsiaTheme"),
        }
        result["fonts"] = {key: value for key, value in font_values.items() if value}
    size = half_points(w_attr(parent.find("w:sz", NS), "val"))
    if size is not None:
        result["font_size_pt"] = size
    for key, tag in (("bold", "b"), ("italic", "i"), ("small_caps", "smallCaps")):
        value = bool_value(parent.find(f"w:{tag}", NS))
        if value is not None:
            result[key] = value
    color = w_attr(parent.find("w:color", NS), "val")
    if color:
        result["color"] = color
    return result


def merge_properties(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = merge_properties(result[key], value)
        else:
            result[key] = value
    return result


def parse_theme_fonts(package: zipfile.ZipFile) -> dict[str, str]:
    theme = read_xml(package, "word/theme/theme1.xml")
    if theme is None:
        return {}
    settings = read_xml(package, "word/settings.xml")
    east_asia_language = ""
    if settings is not None:
        east_asia_language = str(w_attr(settings.find("w:themeFontLang", NS), "eastAsia") or "").lower()
    script = "Hans"
    if east_asia_language.startswith("zh-tw") or east_asia_language.startswith("zh-hk"):
        script = "Hant"
    elif east_asia_language.startswith("ja"):
        script = "Jpan"
    elif east_asia_language.startswith("ko"):
        script = "Hang"

    result: dict[str, str] = {}
    for family, prefix in (("majorFont", "major"), ("minorFont", "minor")):
        node = theme.find(f".//a:{family}", NS)
        if node is None:
            continue
        latin = node.find("a:latin", NS)
        east_asia = node.find("a:ea", NS)
        latin_name = latin.get("typeface") if latin is not None else None
        east_asia_name = east_asia.get("typeface") if east_asia is not None else None
        if not east_asia_name:
            for font in node.findall("a:font", NS):
                if font.get("script") == script and font.get("typeface"):
                    east_asia_name = font.get("typeface")
                    break
        if latin_name:
            result[f"{prefix}HAnsi"] = latin_name
            result[f"{prefix}Ascii"] = latin_name
        if east_asia_name:
            result[f"{prefix}EastAsia"] = east_asia_name
    return result


def resolve_theme_font_names(properties: dict[str, Any], theme_fonts: dict[str, str]) -> dict[str, Any]:
    result = dict(properties)
    fonts = dict(result.get("fonts") or {})
    if not fonts.get("ascii"):
        fonts["ascii"] = theme_fonts.get(str(fonts.get("asciiTheme") or ""))
    if not fonts.get("hAnsi"):
        fonts["hAnsi"] = theme_fonts.get(str(fonts.get("asciiTheme") or ""))
    if not fonts.get("eastAsia"):
        fonts["eastAsia"] = theme_fonts.get(str(fonts.get("eastAsiaTheme") or ""))
    result["fonts"] = {key: value for key, value in fonts.items() if value}
    return result


def parse_paragraph_properties(parent: ET.Element | None) -> dict[str, Any]:
    if parent is None:
        return {}
    result: dict[str, Any] = {}
    alignment = w_attr(parent.find("w:jc", NS), "val")
    alignment = {"start": "left", "end": "right", "both": "justify"}.get(alignment, alignment)
    if alignment:
        result["alignment"] = alignment
    tabs = parent.find("w:tabs", NS)
    if tabs is not None:
        result["tabs"] = [
            {
                key: value for key, value in {
                    "alignment": w_attr(tab, "val"),
                    "position": w_attr(tab, "pos"),
                }.items() if value is not None
            }
            for tab in tabs.findall("w:tab", NS)
        ]
    spacing = parent.find("w:spacing", NS)
    if spacing is not None:
        line_value = w_attr(spacing, "line")
        values = {
            "before_pt": length_to_pt(w_attr(spacing, "before")),
            "after_pt": length_to_pt(w_attr(spacing, "after")),
            "line": line_value,
            "line_pt": length_to_pt(line_value) if line_value and not line_value.lstrip("-.").isdigit() else None,
            "line_rule": w_attr(spacing, "lineRule"),
        }
        result["spacing"] = {key: value for key, value in values.items() if value is not None}
    ind = parent.find("w:ind", NS)
    if ind is not None:
        first_line_chars = w_attr(ind, "firstLineChars")
        left_chars = w_attr(ind, "leftChars")
        hanging_chars = w_attr(ind, "hangingChars")
        values = {
            "left_mm": twips_to_mm(w_attr(ind, "left") or w_attr(ind, "start")),
            "left_chars": round(float(left_chars) / 100, 2) if left_chars and left_chars.replace("-", "").isdigit() else None,
            "right_mm": twips_to_mm(w_attr(ind, "right") or w_attr(ind, "end")),
            "first_line_mm": twips_to_mm(w_attr(ind, "firstLine")),
            "first_line_chars": round(float(first_line_chars) / 100, 2) if first_line_chars and first_line_chars.replace("-", "").isdigit() else None,
            "hanging_mm": twips_to_mm(w_attr(ind, "hanging")),
            "hanging_chars": round(float(hanging_chars) / 100, 2) if hanging_chars and hanging_chars.replace("-", "").isdigit() else None,
        }
        result["indent"] = {key: value for key, value in values.items() if value is not None}
    numbering = parent.find("w:numPr", NS)
    if numbering is not None:
        num_id = w_attr(numbering.find("w:numId", NS), "val")
        level = w_attr(numbering.find("w:ilvl", NS), "val")
        result["numbering"] = {
            key: value for key, value in {
                "num_id": int(num_id) if num_id and num_id.isdigit() else num_id,
                "level": int(level) if level and level.isdigit() else level,
            }.items() if value is not None
        }
    for key, tag in (
        ("keep_with_next", "keepNext"),
        ("keep_together", "keepLines"),
        ("page_break_before", "pageBreakBefore"),
        ("widow_control", "widowControl"),
    ):
        value = bool_value(parent.find(f"w:{tag}", NS))
        if value is not None:
            result[key] = value
    outline = w_attr(parent.find("w:outlineLvl", NS), "val")
    if outline is not None:
        result["outline_level"] = int(outline)
    return result


def parse_styles(
    styles_root: ET.Element | None,
    theme_fonts: dict[str, str] | None = None,
) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    if styles_root is None:
        return [], {}
    default_rpr = parse_run_properties(
        styles_root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
    )
    default_ppr = parse_paragraph_properties(
        styles_root.find("w:docDefaults/w:pPrDefault/w:pPr", NS)
    )
    raw: dict[str, dict[str, Any]] = {}
    ordered: list[str] = []
    for style in styles_root.findall("w:style", NS):
        style_id = w_attr(style, "styleId")
        if not style_id:
            continue
        ordered.append(style_id)
        raw[style_id] = {
            "style_id": style_id,
            "name": w_attr(style.find("w:name", NS), "val") or style_id,
            "type": w_attr(style, "type"),
            "default": w_attr(style, "default") == "1",
            "based_on": w_attr(style.find("w:basedOn", NS), "val"),
            "next": w_attr(style.find("w:next", NS), "val"),
            "quick_format": style.find("w:qFormat", NS) is not None,
            "rPr": parse_run_properties(style.find("w:rPr", NS)),
            "pPr": parse_paragraph_properties(style.find("w:pPr", NS)),
        }

    resolved: dict[str, dict[str, Any]] = {}

    def resolve(style_id: str, stack: set[str] | None = None) -> dict[str, Any]:
        if style_id in resolved:
            return resolved[style_id]
        style = raw[style_id]
        stack = set() if stack is None else set(stack)
        if style_id in stack:
            return style
        stack.add(style_id)
        base_id = style.get("based_on")
        effective_rpr: dict[str, Any] = dict(default_rpr)
        effective_ppr: dict[str, Any] = dict(default_ppr)
        if base_id in raw:
            base = resolve(base_id, stack)
            effective_rpr = merge_properties(effective_rpr, base.get("effective_rPr", {}))
            effective_ppr = merge_properties(effective_ppr, base.get("effective_pPr", {}))
        effective_rpr = merge_properties(effective_rpr, style.get("rPr", {}))
        effective_ppr = merge_properties(effective_ppr, style.get("pPr", {}))
        effective_rpr = resolve_theme_font_names(effective_rpr, theme_fonts or {})
        item = dict(style)
        item["effective_rPr"] = effective_rpr
        item["effective_pPr"] = effective_ppr
        resolved[style_id] = item
        return item

    return [resolve(style_id) for style_id in ordered], resolved


def parse_sections(document_root: ET.Element | None) -> list[dict[str, Any]]:
    if document_root is None:
        return []
    sections: list[dict[str, Any]] = []
    for index, sect in enumerate(document_root.findall(".//w:sectPr", NS), start=1):
        size = sect.find("w:pgSz", NS)
        margins = sect.find("w:pgMar", NS)
        columns = sect.find("w:cols", NS)
        column_items = []
        if columns is not None:
            for column in columns.findall("w:col", NS):
                column_items.append({
                    "width_mm": twips_to_mm(w_attr(column, "w")),
                    "space_mm": twips_to_mm(w_attr(column, "space")),
                })
        references = []
        for kind in ("header", "footer"):
            for reference in sect.findall(f"w:{kind}Reference", NS):
                references.append({
                    "kind": kind,
                    "type": w_attr(reference, "type") or "default",
                    "relationship_id": reference.get(f"{{{NS['r']}}}id"),
                })
        page_numbering = sect.find("w:pgNumType", NS)
        section = {
            "index": index,
            "orientation": (w_attr(size, "orient") or "portrait") if size is not None else None,
            "width_mm": twips_to_mm(w_attr(size, "w")),
            "height_mm": twips_to_mm(w_attr(size, "h")),
            "margins_mm": {
                "top": twips_to_mm(w_attr(margins, "top")),
                "right": twips_to_mm(w_attr(margins, "right")),
                "bottom": twips_to_mm(w_attr(margins, "bottom")),
                "left": twips_to_mm(w_attr(margins, "left")),
                "gutter": twips_to_mm(w_attr(margins, "gutter")),
                "header": twips_to_mm(w_attr(margins, "header")),
                "footer": twips_to_mm(w_attr(margins, "footer")),
            },
            "columns": int(w_attr(columns, "num") or 1) if columns is not None else 1,
            "column_spacing_mm": twips_to_mm(w_attr(columns, "space")) if columns is not None else None,
            "column_details": column_items,
            "title_page": sect.find("w:titlePg", NS) is not None,
            "section_type": w_attr(sect.find("w:type", NS), "val") or "nextPage",
            "header_footer_references": references,
            "page_numbering": {
                key: value for key, value in {
                    "start": w_attr(page_numbering, "start"),
                    "format": w_attr(page_numbering, "fmt"),
                }.items() if value is not None
            } if page_numbering is not None else {},
        }
        sections.append(section)
    return sections


def parse_numbering(numbering_root: ET.Element | None) -> list[dict[str, Any]]:
    if numbering_root is None:
        return []
    abstracts: dict[str, ET.Element] = {}
    for abstract in numbering_root.findall("w:abstractNum", NS):
        abstract_id = w_attr(abstract, "abstractNumId")
        if abstract_id is not None:
            abstracts[abstract_id] = abstract
    result = []
    for num in numbering_root.findall("w:num", NS):
        num_id = w_attr(num, "numId")
        abstract_id = w_attr(num.find("w:abstractNumId", NS), "val")
        levels = []
        abstract = abstracts.get(abstract_id or "")
        if abstract is not None:
            for level in abstract.findall("w:lvl", NS):
                levels.append({
                    "level": w_attr(level, "ilvl"),
                    "start": w_attr(level.find("w:start", NS), "val"),
                    "format": w_attr(level.find("w:numFmt", NS), "val"),
                    "text": w_attr(level.find("w:lvlText", NS), "val"),
                    "paragraph_style": w_attr(level.find("w:pStyle", NS), "val"),
                })
        result.append({"num_id": num_id, "abstract_num_id": abstract_id, "levels": levels})
    return result


def field_inventory(roots: list[ET.Element]) -> dict[str, Any]:
    codes: list[str] = []
    for root in roots:
        for node in root.findall(".//w:instrText", NS):
            if node.text and node.text.strip():
                codes.append(" ".join(node.text.split()))
        for node in root.findall(".//w:fldSimple", NS):
            instruction = w_attr(node, "instr")
            if instruction:
                codes.append(" ".join(instruction.split()))
    counts = Counter(code.split()[0].upper() for code in codes if code.split())
    return {"count": len(codes), "types": dict(counts), "codes": codes}


def summarize_header_footer_parts(package: zipfile.ZipFile, names: set[str]) -> list[dict[str, Any]]:
    result = []
    for name in sorted(names):
        if not name.endswith(".xml") or not name.startswith(CONTENT_PART_PREFIXES):
            continue
        root = read_xml(package, name)
        if root is None:
            continue
        text = "".join(node.text or "" for node in root.findall(".//w:t", NS)).strip()
        fields = field_inventory([root])
        result.append({
            "part": name,
            "kind": "header" if name.startswith("word/header") else "footer",
            "text": text,
            "field_types": fields["types"],
        })
    return result


def _paragraph_text_outside_fields(paragraph: ET.Element) -> str:
    """Return literal paragraph text without cached field results."""
    paragraph_chunks: list[str] = []
    field_state = [False]

    def visit(node: ET.Element) -> None:
        for child in list(node):
            if child.tag == W + "fldSimple":
                continue
            if child.tag == W + "fldChar":
                field_type = w_attr(child, "fldCharType")
                if field_type == "begin":
                    field_state[0] = True
                elif field_type == "end":
                    field_state[0] = False
                continue
            if child.tag == W + "t" and not field_state[0] and child.text:
                paragraph_chunks.append(child.text)
            visit(child)

    visit(paragraph)
    return "".join(paragraph_chunks).strip()


def _text_outside_fields(root: ET.Element) -> str:
    """Return literal header/footer text without cached field results."""
    chunks: list[str] = []
    for paragraph in root.findall(".//w:p", NS):
        text = _paragraph_text_outside_fields(paragraph)
        if text:
            chunks.append(text)
    return "\n".join(chunks)


def extract_document_text(path: str | Path, limit: int = 15000) -> str:
    """Extract literal text for conversational semantic analysis."""
    package, inventory = package_inventory(Path(path).resolve())
    try:
        names = {item["path"] for item in inventory}
        chunks: list[str] = []
        for name in sorted(names):
            if name == "word/document.xml" or is_content_part(name):
                root = read_xml(package, name)
                if root is not None:
                    text = _text_outside_fields(root)
                    if text:
                        chunks.append(text)
        return "\n\n".join(chunks)[:limit]
    finally:
        package.close()


def infer_visual_paragraph_alignment(
    paragraph: ET.Element,
    effective_ppr: dict[str, Any],
) -> str:
    positional_tabs = paragraph.findall(".//w:ptab", NS)
    if positional_tabs:
        alignment = w_attr(positional_tabs[0], "alignment")
        if alignment in {"left", "center", "right"}:
            return alignment

    tabs_before_text = 0
    for node in paragraph.iter():
        if node.tag == W + "tab":
            tabs_before_text += 1
        elif node.tag == W + "t" and str(node.text or "").strip():
            break
    tab_stops = [
        tab for tab in effective_ppr.get("tabs", [])
        if tab.get("alignment") in {"left", "center", "right"}
    ]
    if tabs_before_text and tab_stops:
        selected = tab_stops[min(tabs_before_text - 1, len(tab_stops) - 1)]
        return str(selected["alignment"])

    alignment = str(effective_ppr.get("alignment") or "left")
    return "justify" if alignment == "both" else alignment


def infer_header_footer_tokens(
    package: zipfile.ZipFile,
    names: set[str],
    styles: dict[str, dict[str, Any]],
    theme_fonts: dict[str, str],
    header_footer_parts: list[dict[str, Any]],
) -> dict[str, Any]:
    has_any_content = any(
        part.get("text", "").strip() or part.get("field_types")
        for part in header_footer_parts
    )
    tokens: dict[str, Any] = {
        "preserve_existing": True if has_any_content else False,
        "header": {"enabled": False, "text": ""},
        "footer": {"enabled": False, "text": ""},
    }
    for kind in ("header", "footer"):
        candidates = [part for part in header_footer_parts if part.get("kind") == kind]
        literal_chunks: list[str] = []
        typography: dict[str, Any] = {}
        for part in candidates:
            root = read_xml(package, str(part["part"]))
            if root is None:
                continue
            literal = _text_outside_fields(root)
            if literal:
                literal_chunks.append(literal)
            paragraph = next(
                (
                    item for item in root.findall(".//w:p", NS)
                    if _paragraph_text_outside_fields(item)
                ),
                None,
            )
            if paragraph is not None and not typography:
                ppr = paragraph.find("w:pPr", NS)
                pstyle_id = w_attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
                fallback_style_id = "Header" if kind == "header" else "Footer"
                effective = copy.deepcopy(
                    styles.get(pstyle_id) or styles.get(fallback_style_id) or {}
                )
                effective["effective_rPr"] = merge_properties(
                    effective.get("effective_rPr", {}),
                    parse_run_properties(paragraph.find("w:r/w:rPr", NS)),
                )
                effective["effective_pPr"] = merge_properties(
                    effective.get("effective_pPr", {}),
                    parse_paragraph_properties(ppr),
                )
                effective["effective_pPr"]["alignment"] = infer_visual_paragraph_alignment(
                    paragraph,
                    effective["effective_pPr"],
                )
                effective["effective_rPr"] = resolve_theme_font_names(
                    effective["effective_rPr"], theme_fonts,
                )
                typography = style_tokens(effective)
        literal_text = "\n".join(dict.fromkeys(literal_chunks)).strip()
        tokens[kind] = {
            "enabled": bool(literal_text),
            "text": literal_text,
            **typography,
        }
    return tokens


def semantic_style_summary(styles: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    roles = {"figure_captions": [], "table_captions": [], "table_content": [], "references": []}
    for style in styles:
        marker = f"{style.get('style_id', '')} {style.get('name', '')}".lower().replace("_", " ")
        role = None
        if "figure" in marker and "caption" in marker:
            role = "figure_captions"
        elif "table" in marker and ("caption" in marker or "head" in marker):
            role = "table_captions"
        elif "table" in marker and ("copy" in marker or "footnote" in marker or "col" in marker):
            role = "table_content"
        elif "reference" in marker or "bibliograph" in marker:
            role = "references"
        if role:
            roles[role].append({
                "style_id": style.get("style_id"),
                "name": style.get("name"),
                "usage_count": style.get("usage_count", 0),
                "rPr": style.get("effective_rPr", {}),
                "pPr": style.get("effective_pPr", {}),
            })
    return roles


def infer_document_structure(
    document_root: ET.Element | None,
    styles_by_id: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Extract ordered document modules without inventing absent sections.

    The result is a structural snapshot: paragraph indexes and marker evidence are
    retained for auditing, while representative style tokens are used by the
    deterministic module formatter.
    """
    if document_root is None:
        return {"schema_version": "1.0", "strict_order": True, "ordered_modules": [], "order_errors": []}
    body = document_root.find("w:body", NS)
    if body is None:
        return {"schema_version": "1.0", "strict_order": True, "ordered_modules": [], "order_errors": []}
    paragraphs = [
        node for node in list(body)
        if node.tag == W + "p" and "".join(item.text or "" for item in node.findall(".//w:t", NS)).strip()
    ]
    if not paragraphs:
        return {"schema_version": "1.0", "strict_order": True, "ordered_modules": [], "order_errors": []}
    paragraph_items = []
    for index, paragraph in enumerate(paragraphs):
        ppr = paragraph.find("w:pPr", NS)
        style_id = w_attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
        style = styles_by_id.get(style_id or "")
        paragraph_items.append({
            "index": index,
            "text": "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip(),
            "style_name": style.get("name", "") if style else "",
            "outline_level": style.get("effective_pPr", {}).get("outline_level") if style else None,
        })
    shared_spans = detect_module_spans(paragraph_items)
    if not shared_spans:
        return {"schema_version": "1.0", "strict_order": True, "ordered_modules": [], "order_errors": []}
    starts = [(int(item["paragraph_start"]), str(item["kind"])) for item in shared_spans]
    modules: list[dict[str, Any]] = []
    for position, (start, kind) in enumerate(starts):
        end = (starts[position + 1][0] - 1) if position + 1 < len(starts) else len(paragraphs) - 1
        selected = paragraphs[start : end + 1]
        style_roles: dict[str, dict[str, Any]] = {}
        role_counts: Counter[str] = Counter()
        for offset, paragraph in enumerate(selected):
            index = start + offset
            text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()
            ppr = paragraph.find("w:pPr", NS)
            style_id = w_attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
            style = styles_by_id.get(style_id or "")
            role = classify_role(
                text,
                style.get("name", "") if style else "",
                first=index == start,
            )
            role_counts[role] += 1
            if role not in style_roles and style:
                style_roles[role] = {
                    "style_id": style.get("style_id"),
                    "style_name": style.get("name"),
                    "tokens": style_tokens(style),
                }
        modules.append({
            "id": kind,
            "kind": kind,
            "repeatable": kind == "chapters",
            "paragraph_start": start,
            "paragraph_end": end,
            "heading_text": "".join(node.text or "" for node in selected[0].findall(".//w:t", NS)).strip() if selected else "",
            "paragraph_count": len(selected),
            "role_counts": dict(role_counts),
            "style_roles": style_roles,
        })
    return {
        "schema_version": "1.0",
        "strict_order": True,
        "ordered_modules": modules,
        "order_errors": validate_module_order(modules),
    }


def style_tokens(style: dict[str, Any] | None) -> dict[str, Any]:
    if not style:
        return {}
    rpr = style.get("effective_rPr", {})
    ppr = style.get("effective_pPr", {})
    fonts = rpr.get("fonts", {})
    spacing = ppr.get("spacing", {})
    indent = ppr.get("indent", {})
    result = {
        "style_name": style.get("name"),
        "source_style_id": style.get("style_id"),
        "font_latin": fonts.get("ascii") or fonts.get("hAnsi"),
        "font_east_asia": fonts.get("eastAsia"),
        "font_size_pt": rpr.get("font_size_pt"),
        "bold": rpr.get("bold", False),
        "italic": rpr.get("italic", False),
        "alignment": {"both": "justify", "start": "left", "end": "right"}.get(ppr.get("alignment", "left"), ppr.get("alignment", "left")),
        "space_before_pt": spacing.get("before_pt", 0.0),
        "space_after_pt": spacing.get("after_pt", 0.0),
        "left_indent_mm": indent.get("left_mm", 0.0),
        "left_indent_chars": indent.get("left_chars", 0.0),
        "right_indent_mm": indent.get("right_mm", 0.0),
        "first_line_indent_mm": indent.get("first_line_mm", 0.0),
        "first_line_indent_chars": indent.get("first_line_chars", 0.0),
        "hanging_indent_mm": indent.get("hanging_mm", 0.0),
        "hanging_indent_chars": indent.get("hanging_chars", 0.0),
        "keep_with_next": ppr.get("keep_with_next", False),
        "keep_together": ppr.get("keep_together", False),
        "widow_control": ppr.get("widow_control", True),
        "numbering": ppr.get("numbering"),
    }
    if spacing.get("line"):
        if spacing.get("line_pt") is not None:
            result["line_spacing"] = {"kind": "exact", "value_pt": round(spacing["line_pt"], 2)}
        elif spacing.get("line_rule") in {None, "auto"}:
            result["line_spacing"] = {"kind": "multiple", "value": round(float(spacing["line"]) / 240, 3)}
        else:
            result["line_spacing"] = {"kind": "exact", "value_pt": round(float(spacing["line"]) / 20, 2)}
    else:
        result["line_spacing"] = {"kind": "single"}
    return {key: value for key, value in result.items() if value is not None}


def infer_spec(
    styles: dict[str, dict[str, Any]],
    sections: list[dict[str, Any]],
    header_footer_parts: list[dict[str, Any]],
    fields: dict[str, Any],
    header_footer_tokens: dict[str, Any],
    document_structure: dict[str, Any] | None = None,
    numbering: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    spec: dict[str, Any] = {
        "schema_version": "1.0",
        "id": "distilled-template",
        "name": "Distilled DOCX template",
        "mode": "parameterized",
        "template_required": False,
        "page": {},
        "body": {},
        "headings": [],
        "captions": {},
        "headers_footers": header_footer_tokens,
        "page_numbers": {"enabled": False},
        "table_of_contents": {"enabled": False},
        "references": {},
        "document_structure": document_structure or {"schema_version": "1.0", "strict_order": True, "ordered_modules": []},
        "notes": ["Inferred from OOXML; review before applying."],
    }
    toc_code = next((code for code in fields.get("codes", []) if re.search(r"\bTOC\b", code, re.I)), None)
    if toc_code:
        spec["table_of_contents"]["enabled"] = True
        level_match = re.search(r'\\o\s+"1-(\d)"', toc_code, re.I)
        if level_match:
            spec["table_of_contents"]["max_heading_level"] = int(level_match.group(1))
    if sections:
        section = sections[0]
        margins = section.get("margins_mm", {})
        width = section.get("width_mm")
        height = section.get("height_mm")
        page_size = "Custom"
        if width and height:
            dims = sorted((round(width, 1), round(height, 1)))
            if abs(dims[0] - 210) < 1 and abs(dims[1] - 297) < 1:
                page_size = "A4"
            elif abs(dims[0] - 215.9) < 1 and abs(dims[1] - 279.4) < 1:
                page_size = "Letter"
        spec["page"] = {
            "size": page_size,
            "width_mm": width,
            "height_mm": height,
            "orientation": section.get("orientation"),
            "margin_top_mm": margins.get("top"),
            "margin_right_mm": margins.get("right"),
            "margin_bottom_mm": margins.get("bottom"),
            "margin_left_mm": margins.get("left"),
            "gutter_mm": margins.get("gutter"),
            "header_distance_mm": margins.get("header"),
            "footer_distance_mm": margins.get("footer"),
        }
    body_style = next(
        (
            style for style in styles.values()
            if style["style_id"].lower().replace(" ", "") == "bodytext"
            or style["name"].lower().replace(" ", "") == "bodytext"
        ),
        None,
    )
    normal = body_style or next(
        (
            style for style in styles.values()
            if style["style_id"].lower() == "normal"
            or style["name"].lower() in {"normal", "正文"}
        ),
        None,
    )
    if normal:
        spec["body"] = style_tokens(normal)
    for style in styles.values():
        name = style["name"].lower()
        level = None
        for candidate in range(1, 6):
            if name in {f"heading {candidate}", f"标题 {candidate}", f"标题{candidate}"}:
                level = candidate
                break
        if level is None:
            outline = style.get("effective_pPr", {}).get("outline_level")
            if outline is not None and 0 <= outline <= 4:
                level = outline + 1
        if level is None:
            continue
        heading = {"level": level, **style_tokens(style)}
        spec["headings"].append({key: value for key, value in heading.items() if value is not None})
    spec["headings"] = sorted(spec["headings"], key=lambda item: item["level"])
    def find_style(*markers: str) -> dict[str, Any] | None:
        for style in styles.values():
            if not style.get("usage_count", 0):
                continue
            haystack = f"{style.get('style_id', '')} {style.get('name', '')}".lower().replace("_", " ")
            if all(marker in haystack for marker in markers):
                return style
        return None

    figure_style = find_style("figure", "caption") or find_style("caption")
    table_style = next(
        (
            style for style in styles.values()
            if style.get("usage_count", 0)
            and (
                style.get("style_id", "").lower().replace(" ", "") == "tablehead"
                or style.get("name", "").lower().replace(" ", "") == "tablehead"
            )
        ),
        None,
    ) or find_style("table", "caption")
    reference_style = find_style("reference") or find_style("bibliograph")
    if figure_style:
        spec["captions"]["figure"] = style_tokens(figure_style)
    if table_style:
        spec["captions"]["table"] = style_tokens(table_style)
    if reference_style:
        spec["references"] = style_tokens(reference_style)
    if sections:
        spec["headers_footers"]["different_first_page"] = bool(sections[0].get("title_page"))
    page_parts = [
        part for part in header_footer_parts
        if part.get("field_types", {}).get("PAGE")
    ]
    if page_parts:
        location = str(page_parts[0].get("kind") or "footer")
        has_total = bool(fields.get("types", {}).get("NUMPAGES"))
        start = 1
        if sections:
            raw_start = sections[0].get("page_numbering", {}).get("start")
            if raw_start is not None:
                try:
                    start = int(raw_start)
                except (TypeError, ValueError):
                    start = 1
        spec["page_numbers"] = {
            "enabled": True,
            "location": location,
            "format": "page-x-of-y" if has_total else "number",
            "start": start,
            "show_on_first_page": not bool(sections and sections[0].get("title_page")),
        }
    if header_footer_parts:
        spec["headers_footers"]["detected_parts"] = header_footer_parts
    # Numbering is a template feature, not a universal fallback. Preserve the
    # actual definitions and leave absent list sections empty for a distilled
    # snapshot; the selected preset may provide its own explicit numbering.
    if numbering:
        spec["lists"] = {"source_definitions": numbering}
    return spec


def fuse_dual_track_spec(
    base_spec: dict[str, Any],
    document_root: ET.Element | None,
    styles_by_id: dict[str, dict[str, Any]],
    semantic_spec: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Cross-validate Track 1 (OOXML structures & empirical paragraph metrics) with Track 2 (AI/Semantic requirements)."""
    spec = copy.deepcopy(base_spec)
    corroborations = []

    if document_root is None:
        return spec, {"status": "single_track", "corroborations": []}

    # Track 1 empirical extraction
    paragraphs = document_root.findall(".//w:p", NS)
    body_indents = []
    body_spacings = []
    empirical_captions: dict[str, Any] = {}
    empirical_references: dict[str, Any] = {}
    empirical_lists: dict[str, Any] = {}

    for p in paragraphs:
        t = "".join(el.text or "" for el in p.findall(".//w:t", NS)).strip()
        ppr = p.find("w:pPr", NS)
        style_id = w_attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None

        # Empirical lists
        if re.match(r"^\s*（\d+）", t):
            empirical_lists["numbered"] = {"style": "decimal-fullwidth-parentheses", "start": 1}
        elif re.match(r"^\s*\d+\.", t) and "numbered" not in empirical_lists:
            empirical_lists["numbered"] = {"style": "decimal-period", "start": 1}

        # Empirical captions.  Keep only observations from the supplied
        # document; never fill missing caption rules with a house default.
        if re.match(r"^\s*(?:图|Figure|Fig\.)\s*[\d\.\-]+", t, re.I):
            label_match = re.match(r"^\s*(图|Figure|Fig\.)", t, re.I)
            style = styles_by_id.get(style_id or "")
            empirical_captions["figure"] = {
                **style_tokens(style),
                "label": label_match.group(1) if label_match else "Figure",
            }
        elif re.match(r"^\s*(?:表|Table)\s*[\d\.\-]+", t, re.I):
            label_match = re.match(r"^\s*(表|Table)", t, re.I)
            style = styles_by_id.get(style_id or "")
            empirical_captions["table"] = {
                **style_tokens(style),
                "label": label_match.group(1) if label_match else "Table",
            }

        # Empirical references
        if re.match(r"^\s*\[\d+\]", t):
            empirical_references = {
                **style_tokens(styles_by_id.get(style_id or "")),
                "citation_mode": "numeric",
            }

        # Body paragraph properties
        is_heading = style_id and any(h in style_id.lower() for h in ["heading", "title", "toc"])
        if not is_heading and len(t) > 20:
            if ppr is not None:
                ind = ppr.find("w:ind", NS)
                if ind is not None:
                    fl = w_attr(ind, "firstLine")
                    flc = w_attr(ind, "firstLineChars")
                    if fl:
                        body_indents.append(twips_to_mm(fl))
                    elif flc:
                        body_indents.append(round(int(flc) / 100 * 12 * 25.4 / 72, 3))
                spacing = ppr.find("w:spacing", NS)
                if spacing is not None:
                    line = w_attr(spacing, "line")
                    rule = w_attr(spacing, "lineRule") or "auto"
                    if line:
                        if rule == "auto":
                            body_spacings.append({"kind": "multiple", "value": round(float(line) / 240, 2)})
                        elif rule == "exact":
                            body_spacings.append({"kind": "exact", "value_pt": round(float(line) / 20, 1)})

    # Track 2: semantic rules supplied by the active conversational AI, or the
    # deterministic parser when no conversational result is available yet.
    req_spec: dict[str, Any] = {}
    ai_engine_info = "本地排版认知与规则分析引擎"
    ai_summary = ""
    if semantic_spec:
        req_spec = copy.deepcopy(semantic_spec)
        ai_engine_info = "当前对话 AI"
        ai_summary = str(req_spec.get("ai_analysis_summary") or "")
    else:
        # Document prose is evidence of content, not a formatting requirement.
        # Requirement parsing belongs to the explicit requirements route and
        # must never reinterpret arbitrary template text as a rule.
        req_spec = {}

    # Semantic requirements can describe settings that are not represented by
    # the document's currently applied styles (for example margins or a TOC).
    for section in ("page", "headers_footers", "page_numbers", "table_of_contents"):
        if isinstance(req_spec.get(section), dict):
            spec[section] = merge_properties(spec.get(section, {}), req_spec[section])
    if isinstance(req_spec.get("body"), dict):
        spec["body"] = merge_properties(spec.get("body", {}), req_spec["body"])

    # Dual-track cross-validation & fusion
    # 1. Body line spacing
    if req_spec.get("body", {}).get("line_spacing"):
        spec["body"]["line_spacing"] = req_spec["body"]["line_spacing"]
        corroborations.append({"field": "body.line_spacing", "value": spec["body"]["line_spacing"], "track": f"{ai_engine_info} 条款判定"})
    elif body_spacings:
        most_common_spacing_str = Counter(json.dumps(s, sort_keys=True) for s in body_spacings).most_common(1)[0][0]
        spec["body"]["line_spacing"] = json.loads(most_common_spacing_str)
        corroborations.append({"field": "body.line_spacing", "value": spec["body"]["line_spacing"], "track": "Track 1 段落实测统计分布"})

    # 2. Body first line indent
    if req_spec.get("body", {}).get("first_line_indent_mm"):
        spec["body"]["first_line_indent_mm"] = req_spec["body"]["first_line_indent_mm"]
        corroborations.append({"field": "body.first_line_indent_mm", "value": spec["body"]["first_line_indent_mm"], "track": f"{ai_engine_info} 条款判定"})
    elif body_indents:
        two_chars = [i for i in body_indents if i is not None and 7.0 <= i <= 9.5]
        if two_chars:
            indent = round(Counter(round(i, 3) for i in two_chars).most_common(1)[0][0], 3)
            spec["body"]["first_line_indent_mm"] = indent
            corroborations.append({"field": "body.first_line_indent_mm", "value": indent, "track": "Track 1 段落实测统计分布"})

    # 3. Body fonts & alignment
    if req_spec.get("body", {}).get("font_east_asia"):
        spec["body"]["font_east_asia"] = req_spec["body"]["font_east_asia"]
    if req_spec.get("body", {}).get("font_latin"):
        spec["body"]["font_latin"] = req_spec["body"]["font_latin"]
    if req_spec.get("body", {}).get("font_size_pt"):
        spec["body"]["font_size_pt"] = req_spec["body"]["font_size_pt"]
    if "alignment" not in spec.get("body", {}) or spec["body"]["alignment"] in [None, "left", "both"]:
        spec["body"]["alignment"] = req_spec.get("body", {}).get("alignment", "justify")

    # 4. Headings
    if req_spec.get("headings"):
        detected_by_level = {item.get("level"): item for item in spec.get("headings", [])}
        spec["headings"] = [
            merge_properties(detected_by_level.get(item.get("level"), {}), item)
            for item in req_spec["headings"]
            if isinstance(item, dict) and item.get("level")
        ]
        corroborations.append({"field": "headings", "levels": len(spec["headings"]), "track": f"{ai_engine_info} 标题层级识别"})

    # 5. Captions
    fused_captions = {**spec.get("captions", {}), **empirical_captions, **req_spec.get("captions", {})}
    spec["captions"] = fused_captions
    if fused_captions:
        corroborations.append({"field": "captions", "details": "仅保留模板或文档证据中实际出现的图表题注规则", "track": "双轨交叉印证"})

    # 6. References
    fused_refs = {**spec.get("references", {}), **empirical_references, **req_spec.get("references", {})}
    spec["references"] = fused_refs
    if fused_refs:
        corroborations.append({"field": "references", "standard": spec["references"].get("citation_system"), "track": "双轨交叉印证"})

    # 7. Lists
    fused_lists = {**spec.get("lists", {}), **empirical_lists, **req_spec.get("lists", {})}
    spec["lists"] = fused_lists
    if fused_lists:
        corroborations.append({"field": "lists", "headings": "保留模板或文档实际编号定义", "track": "双轨交叉印证"})

    cv_report = {
        "status": "dual_track_verified" if semantic_spec else "document_only",
        "ai_engine": ai_engine_info,
        "ai_summary": ai_summary,
        "tracks": {
            "track_1": "OOXML 结构与段落实测统计",
            "track_2": ai_engine_info if semantic_spec else "not_run (no explicit requirements supplied)",
        },
        "corroborations": corroborations,
    }
    return spec, cv_report


def analyze_docx(path: str | Path, semantic_spec: dict[str, Any] | None = None) -> dict[str, Any]:
    source = Path(path).resolve()
    package, inventory = package_inventory(source)
    try:
        conformance = detect_conformance(package)
        styles_root = read_xml(package, "word/styles.xml")
        document_root = read_xml(package, "word/document.xml")
        numbering_root = read_xml(package, "word/numbering.xml")
        theme_fonts = parse_theme_fonts(package)
        styles, styles_by_id = parse_styles(styles_root, theme_fonts)
        sections = parse_sections(document_root)
        numbering = parse_numbering(numbering_root)
        numbering_used = bool(document_root is not None and document_root.findall(".//w:numPr", NS))
        usage = Counter()
        paragraph_count = 0
        run_count = 0
        paragraph_direct = 0
        run_direct = 0
        table_count = 0
        if document_root is not None:
            paragraphs = document_root.findall(".//w:p", NS)
            paragraph_count = len(paragraphs)
            table_count = len(document_root.findall(".//w:tbl", NS))
            for paragraph in paragraphs:
                ppr = paragraph.find("w:pPr", NS)
                style_id = w_attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
                usage[style_id or "(none)"] += 1
                if ppr is not None and any(child.tag != W + "pStyle" for child in list(ppr)):
                    paragraph_direct += 1
                runs = paragraph.findall("w:r", NS)
                run_count += len(runs)
                run_direct += sum(1 for run in runs if run.find("w:rPr", NS) is not None)

        names = {item["path"] for item in inventory}
        content_roots: list[ET.Element] = []
        if document_root is not None:
            content_roots.append(document_root)
        for name in sorted(names):
            if name != "word/document.xml" and is_content_part(name):
                root = read_xml(package, name)
                if root is not None:
                    content_roots.append(root)
        fields = field_inventory(content_roots)
        header_footer_parts = summarize_header_footer_parts(package, names)
        document_structure = infer_document_structure(document_root, styles_by_id)
        features = {
            "numbering_definitions": len(numbering),
            "headers": len([name for name in names if name.startswith("word/header") and name.endswith(".xml")]),
            "footers": len([name for name in names if name.startswith("word/footer") and name.endswith(".xml")]),
            "comments": "word/comments.xml" in names,
            "footnotes": "word/footnotes.xml" in names,
            "endnotes": "word/endnotes.xml" in names,
            "tracked_changes": has_element(content_roots, "w", *sorted(TRACKED_CHANGE_TAGS)),
            "fields": bool(fields["count"]),
            "drawings": has_element(content_roots, "w", "drawing", "pict", "object"),
            "text_boxes": has_element(content_roots, "w", "txbxContent"),
            "content_controls": has_element(content_roots, "w", "sdt"),
            "equations": has_element(content_roots, "m", "oMath", "oMathPara"),
            "macros": any(name.endswith("vbaProject.bin") for name in names),
        }
        unsupported = [key for key, present in features.items() if present is True and key in {"tracked_changes", "drawings", "text_boxes", "content_controls", "equations", "macros"}]
        unsupported_field_types = sorted(set(fields["types"]) - {"PAGE", "NUMPAGES", "TOC"})
        if unsupported_field_types:
            unsupported.append(f"fields: {', '.join(unsupported_field_types)}")
        for style in styles:
            style["usage_count"] = usage.get(style["style_id"], 0)
        header_footer_tokens = infer_header_footer_tokens(
            package,
            names,
            styles_by_id,
            theme_fonts,
            header_footer_parts,
        )
        inferred_spec = infer_spec(
            styles_by_id,
            sections,
            header_footer_parts,
            fields,
            header_footer_tokens,
            document_structure,
            numbering if numbering_used else [],
        )
        inferred_spec, cross_validation = fuse_dual_track_spec(
            inferred_spec,
            document_root,
            styles_by_id,
            semantic_spec=semantic_spec,
        )
        return {
            "schema_version": "1.0",
            "source": {
                "path": str(source),
                "filename": source.name,
                "bytes": source.stat().st_size,
                "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            },
            "package": {
                "ooxml_conformance": conformance,
                "part_count": len(inventory),
                "uncompressed_bytes": sum(item["uncompressed_bytes"] for item in inventory),
                "inventory": inventory,
            },
            "document": {
                "paragraphs": paragraph_count,
                "runs": run_count,
                "tables": table_count,
                "sections": sections,
                "style_usage": dict(usage.most_common()),
                "paragraphs_with_direct_formatting": paragraph_direct,
                "runs_with_direct_formatting": run_direct,
                "semantic_styles": semantic_style_summary(styles),
                "numbering": numbering,
                "fields": fields,
                "headers_footers": header_footer_parts,
            },
            "features": features,
            "template_identity": None,
            "styles": styles,
            "inferred_spec": inferred_spec,
            "cross_validation": cross_validation,
            "unsupported": unsupported,
            "warnings": [
                "Inferred values come from OOXML style definitions and the first section; review every distinct section and page pattern.",
                "Rendered page inspection is required before using this analysis as a reusable template.",
            ],
        }
    finally:
        package.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="DOCX or DOTX file")
    parser.add_argument("--output", help="Write JSON to this path instead of stdout")
    args = parser.parse_args()
    try:
        result = analyze_docx(args.input)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(rendered + "\n", encoding="utf-8")
    else:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        sys.stdout.write(rendered + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
