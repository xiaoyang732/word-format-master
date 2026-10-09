#!/usr/bin/env python3
"""Apply supported format-spec tokens to a copy of a DOCX document."""

from __future__ import annotations

import argparse
import base64
import copy
import hashlib
import json
import re
import sys
import uuid
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm, Pt, RGBColor
from lxml import etree

from citation_workflow import apply_citations
from config import LIST_NUMBERING_STYLES, SUPPORTED_SPEC_SCHEMA_VERSIONS
from document_structure import validate_module_order


ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

PAGE_SIZES_MM = {
    "A4": (210.0, 297.0),
    "Letter": (215.9, 279.4),
}

from word_format.spec import legacy_paths

SUPPORTED_APPLICATION_PATHS = legacy_paths()

CHINESE_DIGITS = {
    1: "一", 2: "二", 3: "三", 4: "四", 5: "五",
    6: "六", 7: "七", 8: "八", 9: "九", 10: "十",
    11: "十一", 12: "十二", 13: "十三", 14: "十四", 15: "十五",
    16: "十六", 17: "十七", 18: "十八", 19: "十九", 20: "二十",
    21: "二十一", 22: "二十二", 23: "二十三", 24: "二十四", 25: "二十五",
    26: "二十六", 27: "二十七", 28: "二十八", 29: "二十九", 30: "三十",
}
CN_TO_DIGIT = {v: k for k, v in CHINESE_DIGITS.items()}


def _to_chinese_num(n: int) -> str:
    if n in CHINESE_DIGITS:
        return CHINESE_DIGITS[n]
    if n < 100:
        tens, ones = divmod(n, 10)
        t_str = "十" if tens == 1 else CHINESE_DIGITS.get(tens, str(tens)) + "十"
        o_str = CHINESE_DIGITS.get(ones, "") if ones else ""
        return t_str + o_str
    return str(n)


def _to_arabic_num(s: str) -> int | None:
    if s.isdigit():
        return int(s)
    if s in CN_TO_DIGIT:
        return CN_TO_DIGIT[s]
    return None



def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def load_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _resolved_path(value: str | Path, label: str) -> Path:
    try:
        return Path(value).resolve(strict=False)
    except (OSError, TypeError) as exc:
        raise ValueError(f"{label} path is invalid") from exc


def _without_citations(spec: dict[str, Any]) -> dict[str, Any]:
    """Keep the user-confirmed layout immutable while AI adds citation placements."""
    result = copy.deepcopy(spec)
    result.pop("citations", None)
    return result


def load_confirmed_handoff(
    path: str | Path,
    source: str | Path,
    destination: str | Path,
) -> dict[str, Any]:
    """Load a dashboard handoff and bind it to the exact document operation."""
    handoff = load_json(path)
    if not isinstance(handoff, dict):
        raise ValueError("Dashboard handoff must be a JSON object")
    if handoff.get("kind") != "word-format-ai-handoff" or handoff.get("status") != "ready-for-ai":
        raise ValueError("Formatting requires a confirmed Dashboard handoff")
    if handoff.get("execution_mode") != "parameterized":
        raise ValueError("Official-template handoffs must follow their publisher workflow")
    source_info = handoff.get("source")
    if not isinstance(source_info, dict) or not str(source_info.get("path") or ""):
        raise ValueError("Dashboard handoff is missing its source document")
    actual_source = _resolved_path(source, "Input")
    handoff_source = _resolved_path(str(source_info["path"]), "Handoff source")
    if actual_source != handoff_source:
        raise ValueError("Input DOCX does not match the confirmed Dashboard handoff")
    actual_destination = _resolved_path(destination, "Output")
    handoff_destination = _resolved_path(str(handoff.get("output_path") or ""), "Handoff output")
    if actual_destination != handoff_destination:
        raise ValueError("Output DOCX does not match the confirmed Dashboard handoff")
    if not actual_source.is_file():
        raise ValueError("Input DOCX no longer exists")
    actual_sha256 = hashlib.sha256(actual_source.read_bytes()).hexdigest()
    if actual_sha256.lower() != str(source_info.get("sha256") or "").lower():
        raise ValueError("Input DOCX changed after Dashboard confirmation; reopen the Dashboard")
    if not isinstance(handoff.get("spec"), dict):
        raise ValueError("Dashboard handoff is missing its confirmed specification")
    return handoff


def normalize_spec(data: dict[str, Any]) -> dict[str, Any]:
    if "spec_patch" in data:
        return data["spec_patch"]
    if "spec" in data:
        return data["spec"]
    return data


def get_or_create_paragraph_style(document: Document, name: str):
    try:
        return document.styles[name]
    except KeyError:
        return document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def resolve_paragraph_style(document: Document, tokens: dict[str, Any], fallback: str):
    source_style_id = str(tokens.get("source_style_id") or "")
    if source_style_id:
        for style in document.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH and style.style_id == source_style_id:
                return style
    for candidate in (tokens.get("style_name"), fallback):
        if not candidate:
            continue
        try:
            style = document.styles[str(candidate)]
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                return style
        except KeyError:
            pass
    return get_or_create_paragraph_style(document, fallback)


def set_style_font(style, tokens: dict[str, Any]) -> None:
    from word_format.properties import element, write_font
    from word_format.spec import FONT, token_operations
    for action, params in token_operations({k: v for k, v in tokens.items() if k in FONT}):
        key = next(k for k, v in FONT.items() if v == action)
        write_font(element(style.element, "rPr"), key, params["value"])


def set_paragraph_format(style, tokens: dict[str, Any]) -> None:
    from word_format.properties import element, write_paragraph
    from word_format.spec import PARA, INDENTS, token_operations
    selected = {k: v for k, v in tokens.items() if k in PARA or any(k.startswith(n + "_") for n in INDENTS)}
    if "font_size_pt" in tokens:
        selected["font_size_pt"] = tokens["font_size_pt"]
    for action, params in token_operations(selected):
        if not action.startswith("paragraph."):
            continue
        key = action[len("paragraph."):-len(".set")]
        key = {"spacing_before": "space_before_pt", "spacing_after": "space_after_pt"}.get(key, key)
        value = params if key == "line_spacing" or key.endswith("indent") else params["value"]
        node = style._p if hasattr(style, "_p") else style.element
        write_paragraph(element(node, "pPr"), key, value)


def set_run_font(run, tokens: dict[str, Any]) -> None:
    from word_format.properties import element, write_font
    from word_format.spec import FONT, token_operations
    for action, params in token_operations({k: v for k, v in tokens.items() if k in FONT}):
        key = next(k for k, v in FONT.items() if v == action)
        write_font(element(run._r, "rPr"), key, params["value"])


def clear_paragraph_content(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def get_managed_paragraph(container, style_name: str):
    for paragraph in container.paragraphs:
        if paragraph.style and paragraph.style.name == style_name:
            clear_paragraph_content(paragraph)
            return paragraph
    if len(container.paragraphs) == 1 and not container.paragraphs[0].text.strip():
        paragraph = container.paragraphs[0]
        clear_paragraph_content(paragraph)
        return paragraph
    return container.add_paragraph()


def append_field(paragraph, instruction: str, cached_text: str = "1") -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction_node = OxmlElement("w:instrText")
    instruction_node.set(qn("xml:space"), "preserve")
    instruction_node.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    nodes = [begin, instruction_node, separate]
    if cached_text:
        cached_lines = str(cached_text).splitlines() or [str(cached_text)]
        for index, line in enumerate(cached_lines):
            if index:
                nodes.append(OxmlElement("w:br"))
            text = OxmlElement("w:t")
            text.text = line
            nodes.append(text)
    nodes.append(end)
    for node in nodes:
        run._r.append(node)


def mark_fields_for_update(document: Document) -> None:
    settings = document.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


def apply_tables(document: Document, body: dict[str, Any], tables: dict[str, Any]) -> list[str]:
    if not tables or not document.tables:
        return []
    changes: list[str] = []
    three_line = bool(tables.get("three_line_table", False))
    repeat_header = bool(tables.get("repeat_header_row", False))
    alignment = tables.get("alignment")
    latin = tables.get("font_latin") or body.get("font_latin")
    east_asia = tables.get("font_east_asia") or body.get("font_east_asia")
    font_size_pt = tables.get("font_size_pt")

    for tbl_idx, table in enumerate(document.tables):
        tblPr = table._tbl.tblPr
        if alignment in ("center", "centre"):
            existing_jc = tblPr.find(qn("w:jc"))
            if existing_jc is not None:
                tblPr.remove(existing_jc)
            tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
        elif alignment in ("left", "right"):
            existing_jc = tblPr.find(qn("w:jc"))
            if existing_jc is not None:
                tblPr.remove(existing_jc)
            tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{alignment}"/>'))

        snapshot_key = "{urn:word-format-master:format}originalBorders"
        if three_line and snapshot_key not in tblPr.attrib:
            border_nodes = [tblPr] + [cell for row in table.rows for cell in row.cells]
            snapshot = []
            seen_cells = set()
            for node in border_nodes:
                owner = node if node is tblPr else node._tc.get_or_add_tcPr()
                if owner in seen_cells:
                    continue
                seen_cells.add(owner)
                border = owner.find(qn("w:tblBorders" if node is tblPr else "w:tcBorders"))
                snapshot.append(base64.b64encode(etree.tostring(border)).decode() if border is not None else None)
            tblPr.set(snapshot_key, json.dumps(snapshot))
        if "three_line_table" in tables and not three_line and snapshot_key in tblPr.attrib:
            snapshot = json.loads(tblPr.attrib.pop(snapshot_key))
            owners = [tblPr]
            for row in table.rows:
                for cell in row.cells:
                    owner = cell._tc.get_or_add_tcPr()
                    if owner not in owners:
                        owners.append(owner)
            if len(snapshot) != len(owners):
                raise ValueError("Table structure changed; saved borders cannot be restored")
            for index, (owner, saved) in enumerate(zip(owners, snapshot)):
                border = owner.find(qn("w:tblBorders" if index == 0 else "w:tcBorders"))
                if border is not None:
                    owner.remove(border)
                if saved:
                    owner.append(parse_xml(base64.b64decode(saved)))
        if three_line:
            existing_borders = tblPr.find(qn("w:tblBorders"))
            if existing_borders is not None:
                tblPr.remove(existing_borders)
            three_line_xml = parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                f'<w:left w:val="none"/>'
                f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
                f'<w:right w:val="none"/>'
                f'<w:insideH w:val="none"/>'
                f'<w:insideV w:val="none"/>'
                f'</w:tblBorders>'
            )
            tblPr.append(three_line_xml)

            if table.rows:
                header_row = table.rows[0]
                for cell in header_row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    existing_tc_borders = tcPr.find(qn("w:tcBorders"))
                    if existing_tc_borders is not None:
                        tcPr.remove(existing_tc_borders)
                    tc_border_xml = parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>'
                        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
                        f'</w:tcBorders>'
                    )
                    tcPr.append(tc_border_xml)

        if "repeat_header_row" in tables and not repeat_header and table.rows:
            trPr = table.rows[0]._tr.trPr
            header = trPr.find(qn("w:tblHeader")) if trPr is not None else None
            if header is not None:
                trPr.remove(header)
        if repeat_header and table.rows:
            header_tr = table.rows[0]._tr
            trPr = header_tr.get_or_add_trPr()
            if trPr.find(qn("w:tblHeader")) is None:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        if latin or east_asia or font_size_pt is not None:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if font_size_pt is not None:
                                set_run_font(run, {"font_size_pt":float(font_size_pt)})
                            rpr = run._r.get_or_add_rPr()
                            rfonts = rpr.get_or_add_rFonts()
                            if latin:
                                run.font.name = latin
                                rfonts.set(qn("w:ascii"), latin)
                                rfonts.set(qn("w:hAnsi"), latin)
                            if east_asia:
                                rfonts.set(qn("w:eastAsia"), east_asia)

    desc = []
    if three_line:
        desc.append("three-line borders")
    if repeat_header:
        desc.append("repeat header row")
    if alignment:
        desc.append(f"alignment={alignment}")
    if font_size_pt:
        desc.append(f"font_size={font_size_pt}pt")
    changes.append(f"Applied table formatting ({', '.join(desc)}) to {len(document.tables)} table(s)")
    return changes


def apply_caption_styles(document: Document, body: dict[str, Any], captions: dict[str, Any]) -> list[str]:
    changes: list[str] = []
    fallbacks = {"figure": "Figure Caption", "table": "Table Caption"}
    default_labels = {"figure": ["Figure", "Fig.", "图"], "table": ["Table", "表"]}
    for role, fallback in fallbacks.items():
        tokens = captions.get(role) or {}
        if not tokens:
            continue
        style = resolve_paragraph_style(document, tokens, fallback)
        if style.base_style is None:
            style.base_style = get_or_create_paragraph_style(document, "Normal")
        inherited_font = {
            key: body[key]
            for key in ("font_latin", "font_east_asia", "font_size_pt", "color")
            if body.get(key) is not None
        }
        effective = deep_merge(inherited_font, tokens)
        set_style_font(style, effective)
        set_paragraph_format(style, effective)
        labels = [str(tokens.get("label") or "")] + default_labels[role]
        labels = [label for label in dict.fromkeys(labels) if label]
        pattern = re.compile(rf"^\s*(?:{'|'.join(re.escape(label) for label in labels)})\s*[\[（(]?[A-Za-z0-9一二三四五六七八九十]+", re.I)
        styled = 0
        for paragraph in document.paragraphs:
            if paragraph.text and pattern.match(paragraph.text):
                paragraph.style = style
                styled += 1
        changes.append(f"Updated {role} caption style '{style.name}' ({styled} existing caption(s) matched)")
    return changes


def _is_empty_top_level_paragraph(element) -> bool:
    drawing_tags = {qn("w:drawing"), qn("w:pict")}
    has_drawing = any(child.tag in drawing_tags for child in element.iter())
    return element.tag == qn("w:p") and not has_drawing and not "".join(element.itertext()).strip()


def _is_figure_paragraph(element) -> bool:
    if element.tag != qn("w:p"):
        return False
    drawing_tags = {qn("w:drawing"), qn("w:pict")}
    return any(child.tag in drawing_tags for child in element.iter())


def _neighboring_content_element(document: Document, paragraph_element, direction: int):
    body = document.element.body
    children = list(body)
    index = children.index(paragraph_element) + direction
    while 0 <= index < len(children):
        candidate = children[index]
        if candidate.tag == qn("w:sectPr") or _is_empty_top_level_paragraph(candidate):
            index += direction
            continue
        return candidate
    return None


def _move_caption_next_to_target(document: Document, caption_element, target_element, position: str) -> bool:
    if target_element is None or target_element is caption_element:
        return False
    body = document.element.body
    children = list(body)
    caption_index = children.index(caption_element)
    target_index = children.index(target_element)
    desired_index = target_index if position == "above" else target_index + 1
    if caption_index == desired_index or (position == "above" and caption_index == target_index - 1) or (position == "below" and caption_index == target_index + 1):
        return False
    body.remove(caption_element)
    children = list(body)
    target_index = children.index(target_element)
    body.insert(target_index if position == "above" else target_index + 1, caption_element)
    return True


def apply_caption_positions(document: Document, captions: dict[str, Any]) -> list[str]:
    """Move only adjacent top-level caption paragraphs around their table/figure.

    The engine is deliberately conservative: it does not guess an anchor when a
    caption is separated from the object by prose, a page break, or a table cell.
    """
    changes: list[str] = []
    default_labels = {"figure": ["Figure", "Fig.", "图"], "table": ["Table", "表"]}
    paragraphs = list(document.paragraphs)
    for role in ("figure", "table"):
        tokens = captions.get(role) or {}
        position = tokens.get("position")
        if position not in {"above", "below"}:
            continue
        labels = [str(tokens.get("label") or "")] + default_labels[role]
        labels = [label for label in dict.fromkeys(labels) if label]
        pattern = re.compile(rf"^\s*(?:{'|'.join(re.escape(label) for label in labels)})\s*[\[（(]?[A-Za-z0-9一二三四五六七八九十]+", re.I)
        moved = 0
        for paragraph in paragraphs:
            if paragraph._p.getparent() is not document.element.body or not pattern.match(paragraph.text or ""):
                continue
            previous = _neighboring_content_element(document, paragraph._p, -1)
            following = _neighboring_content_element(document, paragraph._p, 1)
            target = previous if (previous is not None and (previous.tag == qn("w:tbl") if role == "table" else _is_figure_paragraph(previous))) else None
            if target is None and following is not None and (following.tag == qn("w:tbl") if role == "table" else _is_figure_paragraph(following)):
                target = following
            if _move_caption_next_to_target(document, paragraph._p, target, position):
                moved += 1
        if moved:
            changes.append(f"Moved {moved} {role} caption paragraph(s) to the {position} of adjacent object(s)")
        else:
            changes.append(f"Checked {role} caption placement; no safe adjacent object move was required")
    return changes


def iter_header_footer_containers(document: Document) -> Iterable:
    for section in document.sections:
        for name in (
            "header",
            "footer",
            "first_page_header",
            "first_page_footer",
            "even_page_header",
            "even_page_footer",
        ):
            yield getattr(section, name)


def clear_managed_header_footer_paragraphs(document: Document, style_names: set[str]) -> int:
    cleared = 0
    seen: set[int] = set()
    for container in iter_header_footer_containers(document):
        identity = id(container._element)
        if identity in seen:
            continue
        seen.add(identity)
        for paragraph in container.paragraphs:
            if (paragraph.style and paragraph.style.name in style_names) or (
                "WFM Header" in style_names and paragraph._p.get("{urn:word-format-master}managed-header") == "true"
            ):
                clear_paragraph_content(paragraph)
                paragraph.style = get_or_create_paragraph_style(document, "Normal")
                cleared += 1
    return cleared


def clear_all_header_footer_content(document: Document) -> int:
    cleared = 0
    seen: set[int] = set()
    for container in iter_header_footer_containers(document):
        identity = id(container._element)
        if identity in seen:
            continue
        seen.add(identity)
        for child in list(container._element):
            container._element.remove(child)
            cleared += 1
    return cleared


def _field_codes(paragraph) -> list[str]:
    return [
        " ".join(str(node.text or "").split())
        for node in paragraph._p.iter(qn("w:instrText"))
        if str(node.text or "").strip()
    ]


def _standalone_page_number_paragraph(container):
    for paragraph in container.paragraphs:
        codes = _field_codes(paragraph)
        if not any(re.search(r"\bPAGE\b", code, re.I) for code in codes):
            continue
        if any(not re.fullmatch(r"(?:PAGE|NUMPAGES)(?:\s+\\\S+)*", code, re.I) for code in codes):
            continue
        visible = paragraph.text.strip()
        if not visible or re.fullmatch(r"(?:Page\s*)?(?:第\s*)?\d+(?:\s*(?:of|/|页，共)\s*\d+\s*页?)?", visible, re.I):
            return paragraph
    return None


def apply_headers_footers(document: Document, tokens: dict[str, Any]) -> list[str]:
    if not tokens:
        return []
    changes: list[str] = []
    if tokens.get("preserve_existing") is False:
        cleared = clear_all_header_footer_content(document)
        changes.append(f"Cleared {cleared} existing header/footer content element(s)")
    if document.sections and tokens.get("different_first_page") is not None:
        for section in document.sections:
            section.different_first_page_header_footer = bool(tokens["different_first_page"])
        changes.append(f"Updated first-page header/footer behavior in {len(document.sections)} section(s)")
    if tokens.get("different_odd_even") is not None and hasattr(document.settings, "odd_and_even_pages_header_footer"):
        document.settings.odd_and_even_pages_header_footer = bool(tokens["different_odd_even"])
        changes.append("Updated odd/even header/footer behavior")
    for kind in ("header", "footer"):
        config = tokens.get(kind) or {}
        if config.get("enabled") is False or (config.get("enabled") and not str(config.get("text", "")).strip()):
            cleared = clear_managed_header_footer_paragraphs(document, {f"WFM {kind.title()}"})
            changes.append(f"Removed managed {kind} text from {cleared} part(s)")
            continue
        if not config.get("enabled"):
            continue
        if kind == "header":
            from word_format.inspect import DocumentIndex
            from word_format.headers import set_content, content_params
            from word_format.spec import token_operations
            from word_format.registry import REGISTRY
            index = DocumentIndex(document)
            for oid, record in index.records.items():
                if record["kind"] != "section": continue
                ref = index.ref(oid)
                set_content(index, ref, content_params({"mode":"text", "text":config["text"]}))
                for action, params in token_operations({k:v for k,v in config.items() if k not in {"enabled","text"}}):
                    action = "header." + (action if action.startswith("font.") else action.removeprefix("paragraph."))
                    cap = REGISTRY[action]
                    cap.handler(index, ref, cap.validator(params))
            changes.append("Updated header through registered section header handlers")
            continue
        style_name = f"WFM {kind.title()}"
        style = get_or_create_paragraph_style(document, style_name)
        set_style_font(style, config)
        set_paragraph_format(style, config)
        seen: set[int] = set()
        for section in document.sections:
            container = getattr(section, kind)
            identity = id(container._element)
            if identity in seen:
                continue
            seen.add(identity)
            paragraph = get_managed_paragraph(container, style_name)
            paragraph.style = style
            paragraph.alignment = ALIGNMENTS.get(config.get("alignment"), WD_ALIGN_PARAGRAPH.CENTER)
            run = paragraph.add_run(str(config["text"]))
            set_run_font(run, config)
        changes.append(f"Added managed {kind} text to {len(seen)} unique part(s)")
    return changes


def apply_page_numbers(document: Document, tokens: dict[str, Any]) -> list[str]:
    if not tokens:
        return []
    cleared = clear_managed_header_footer_paragraphs(document, {"WFM Page Number"})
    if not tokens.get("enabled"):
        return [f"Removed managed page-number fields from {cleared} part(s)"]
    location = tokens.get("location", "footer")
    style_name = "WFM Page Number"
    style = get_or_create_paragraph_style(document, style_name)
    set_style_font(style, tokens)
    set_paragraph_format(style, tokens)
    start = int(tokens.get("start", 1))
    if document.sections:
        sect_pr = document.sections[0]._sectPr
        page_type = sect_pr.find(qn("w:pgNumType"))
        if page_type is None:
            page_type = OxmlElement("w:pgNumType")
            sect_pr.append(page_type)
        page_type.set(qn("w:start"), str(start))
        if not tokens.get("show_on_first_page", True):
            document.sections[0].different_first_page_header_footer = True
    containers = []
    for section in document.sections:
        containers.append(getattr(section, location))
        if document.settings.odd_and_even_pages_header_footer:
            containers.append(getattr(section, f"even_page_{location}"))
    if tokens.get("show_on_first_page", True):
        for section in document.sections:
            if section.different_first_page_header_footer:
                containers.append(getattr(section, f"first_page_{location}"))
    seen: set[int] = set()
    reused = 0
    for container in containers:
        identity = id(container._element)
        if identity in seen:
            continue
        seen.add(identity)
        paragraph = _standalone_page_number_paragraph(container)
        if paragraph is not None:
            clear_paragraph_content(paragraph)
            reused += 1
        else:
            paragraph = get_managed_paragraph(container, style_name)
        paragraph.style = style
        paragraph.alignment = ALIGNMENTS.get(tokens.get("alignment"), WD_ALIGN_PARAGRAPH.CENTER)
        field_format = tokens.get("format", "number")
        if field_format in {"page-number", "page-x-of-y"}:
            paragraph.add_run(str(tokens.get("prefix", "Page ")))
        append_field(paragraph, "PAGE")
        if field_format == "page-x-of-y":
            paragraph.add_run(str(tokens.get("total_separator", " of ")))
            append_field(paragraph, "NUMPAGES")
        for run in paragraph.runs:
            set_run_font(run, tokens)
    mark_fields_for_update(document)
    return [
        f"Replaced {cleared} managed page-number paragraph(s)",
        f"Reused {reused} existing standalone page-number paragraph(s)",
        f"Inserted PAGE field(s) in {len(seen)} unique {location} part(s), starting at {start}",
    ]


TOC_MANAGED_STYLES = {"WFM TOC Heading", "WFM TOC Field", "WFM TOC Break"}


def _remove_managed_toc(document: Document) -> int:
    removed = 0
    for paragraph in list(document.paragraphs):
        if paragraph.style and paragraph.style.name in TOC_MANAGED_STYLES:
            paragraph._p.getparent().remove(paragraph._p)
            removed += 1
    return removed


def _insert_body_paragraph_before(document: Document, anchor, style):
    paragraph = document.add_paragraph(style=style)
    if anchor is not None:
        anchor.addprevious(paragraph._p)
    return paragraph


def _find_toc_insertion_anchor(document: Document):
    """Find (anchor_paragraph, placeholder_toc_paragraph) for managed TOC insertion.

    TOC must be placed AFTER all abstract sections (Chinese abstract + keywords,
    and optional English abstract + keywords) and BEFORE subsequent document modules
    (e.g., Chapter 1 / symbols / body).

    If the document contains an unmanaged placeholder TOC heading, that position is
    used and the redundant placeholder paragraph will be removed upon insertion.
    """
    re_abstract_cn = re.compile(r"^\s*(?:摘\s*要|中文摘要|内容摘要)(?:\s*[:：].*)?$", re.I)
    re_abstract_en = re.compile(r"^\s*(?:abstracts?|英文摘要)(?:\s*[:：].*)?$", re.I)
    re_keywords_cn = re.compile(r"^\s*(?:关\s*键\s*词)(?:\s*[:：].*)?$", re.I)
    re_keywords_en = re.compile(r"^\s*(?:key\s*words?)(?:\s*[:：].*)?$", re.I)
    re_toc = re.compile(r"^\s*(?:目\s*录|table\s+of\s+contents|contents)\s*$", re.I)
    re_next_module = re.compile(
        r"^\s*(?:第\s*[一二三四五六七八九十0-9]+\s*章|\d+(?:[、.\s]|\s*)[^\s\d.]|符号说明|符号表|术语表|symbols?|glossary|参考文献|参考资料|致\s*谢|附\s*录|references|bibliography|acknowledg(?:e)?ments?|appendix)",
        re.I,
    )

    paragraphs = [p for p in document.paragraphs if p.text.strip() or p._p.find(qn("w:drawing")) is not None]
    if not paragraphs:
        return None, None

    placeholder_toc_p = None
    for p in paragraphs:
        if p.style and p.style.name in TOC_MANAGED_STYLES:
            continue
        if re_toc.match(p.text.strip()):
            placeholder_toc_p = p
            break

    state = "SEEKING_ABSTRACT"
    anchor_p = None

    for p in paragraphs:
        if p.style and p.style.name in TOC_MANAGED_STYLES:
            continue
        t = p.text.strip()
        if not t:
            continue
        style_name = p.style.name.lower() if p.style else ""
        is_h1 = style_name.startswith("heading 1") or style_name.startswith("标题 1") or style_name == "heading"

        if state == "SEEKING_ABSTRACT":
            if re_abstract_cn.match(t):
                state = "IN_CHINESE_ABSTRACT"
            elif re_abstract_en.match(t):
                state = "IN_ENGLISH_ABSTRACT"
            elif re_toc.match(t):
                anchor_p = p
                break
            elif re_next_module.match(t) or is_h1:
                anchor_p = p
                break
        elif state == "IN_CHINESE_ABSTRACT":
            if re_keywords_cn.match(t) or re_keywords_en.match(t):
                state = "IN_CHINESE_KEYWORDS"
            elif re_abstract_en.match(t):
                state = "IN_ENGLISH_ABSTRACT"
            elif re_toc.match(t) or re_next_module.match(t) or is_h1:
                anchor_p = p
                break
        elif state == "IN_CHINESE_KEYWORDS":
            if re_abstract_en.match(t):
                state = "IN_ENGLISH_ABSTRACT"
            else:
                anchor_p = p
                break
        elif state == "IN_ENGLISH_ABSTRACT":
            if re_keywords_en.match(t) or re_keywords_cn.match(t):
                state = "IN_ENGLISH_KEYWORDS"
            elif re_toc.match(t) or re_next_module.match(t) or is_h1:
                anchor_p = p
                break
        elif state == "IN_ENGLISH_KEYWORDS":
            anchor_p = p
            break

    if placeholder_toc_p is not None and anchor_p is None:
        anchor_p = placeholder_toc_p
    elif anchor_p is None and state == "SEEKING_ABSTRACT":
        anchor_p = paragraphs[0] if paragraphs else None

    return anchor_p, placeholder_toc_p


def apply_table_of_contents(document: Document, tokens: dict[str, Any]) -> list[str]:
    if not tokens:
        return []
    removed = _remove_managed_toc(document)
    if not tokens.get("enabled"):
        return [f"Removed managed Word TOC block ({removed} paragraph(s))"]

    title = str(tokens.get("title") or "目录")
    max_level = int(tokens.get("max_heading_level", 3))
    heading_style = get_or_create_paragraph_style(document, "WFM TOC Heading")
    field_style = get_or_create_paragraph_style(document, "WFM TOC Field")
    break_style = get_or_create_paragraph_style(document, "WFM TOC Break")
    for style in (heading_style, field_style, break_style):
        if style.base_style is None:
            style.base_style = get_or_create_paragraph_style(document, "Normal")
    set_style_font(heading_style, {"bold": True, "font_size_pt": 14})
    set_paragraph_format(heading_style, {"alignment": "center", "space_before_pt": 0, "space_after_pt": 8, "keep_with_next": True})
    set_paragraph_format(field_style, {"space_before_pt": 0, "space_after_pt": 8})
    set_paragraph_format(break_style, {"space_before_pt": 0, "space_after_pt": 0})

    anchor_p, placeholder_toc_p = _find_toc_insertion_anchor(document)
    anchor_xml = anchor_p._p if anchor_p is not None else None

    # Check if there is preceding content before the TOC (e.g. abstract / keywords / cover)
    # and ensure there is a page break on the last page of the abstract so TOC and abstract are on separate pages.
    has_preceding_content = False
    if anchor_xml is not None:
        prev = anchor_xml.getprevious()
        while prev is not None:
            if prev.tag == qn("w:p"):
                has_preceding_content = True
                break
            prev = prev.getprevious()
    elif len(document.paragraphs) > 0:
        has_preceding_content = True

    if has_preceding_content:
        prev_p = anchor_xml.getprevious() if anchor_xml is not None else (document.paragraphs[-1]._p if document.paragraphs else None)
        prev_has_break = False
        while prev_p is not None:
            if prev_p.tag == qn("w:p"):
                for br in prev_p.iter(qn("w:br")):
                    if br.get(qn("w:type")) == "page":
                        prev_has_break = True
                        break
                pPr = prev_p.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None:
                    prev_has_break = True
                break
            prev_p = prev_p.getprevious()

        if not prev_has_break:
            leading_break = _insert_body_paragraph_before(document, anchor_xml, break_style)
            leading_break.add_run().add_break(WD_BREAK.PAGE)

    title_paragraph = _insert_body_paragraph_before(document, anchor_xml, heading_style)
    title_paragraph.add_run(title)
    field_paragraph = _insert_body_paragraph_before(document, anchor_xml, field_style)
    cached_entries = []
    for paragraph in document.paragraphs:
        style_name = paragraph.style.name if paragraph.style else ""
        if style_name in TOC_MANAGED_STYLES or style_name.lower().startswith("wfm toc"):
            continue
        text = paragraph.text.strip()
        if not text:
            continue
        if re.match(r"^\s*(?:目\s*录|table\s+of\s+contents|contents)\s*$", text, re.I):
            continue
        match = re.fullmatch(r"Heading ([1-9])", style_name, re.I)
        if match and int(match.group(1)) <= max_level:
            level = int(match.group(1))
            indent = "  " * (level - 1)
            cached_entries.append(f"{indent}{text}")
    append_field(
        field_paragraph,
        f'TOC \\o "1-{max_level}" \\h \\z \\u',
        cached_text="\n".join(cached_entries) or "目录将在 Word 中更新",
    )
    if tokens.get("page_break_after", True):
        break_paragraph = _insert_body_paragraph_before(document, anchor_xml, break_style)
        break_paragraph.add_run().add_break(WD_BREAK.PAGE)

    if placeholder_toc_p is not None and placeholder_toc_p._p.getparent() is not None:
        placeholder_toc_p._p.getparent().remove(placeholder_toc_p._p)

    mark_fields_for_update(document)
    break_note = " and page break" if tokens.get("page_break_after", True) else ""
    return [
        f"Inserted managed Word TOC field for Heading 1-{max_level}{break_note}",
        f"Added {len(cached_entries)} cached TOC heading label(s) for non-Word renderers",
        f"Removed {removed} previous managed TOC paragraph(s)",
    ]


def _set_paragraph_numbering(paragraph, num_id: int, ilvl: int) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    num_pr = pPr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        pPr.append(num_pr)
    else:
        num_pr.clear()

    ilvl_node = OxmlElement("w:ilvl")
    ilvl_node.set(qn("w:val"), str(ilvl))
    num_id_node = OxmlElement("w:numId")
    num_id_node.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl_node)
    num_pr.append(num_id_node)


def _remove_paragraph_numbering(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    num_pr = pPr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        pPr.append(num_pr)
    else:
        num_pr.clear()
    num_id_node = OxmlElement("w:numId")
    num_id_node.set(qn("w:val"), "0")
    num_pr.append(num_id_node)


def _bind_style_numbering(document: Document, style_name: str, num_id: int, ilvl: int) -> None:
    try:
        style = get_or_create_paragraph_style(document, style_name)
        pPr = style._element.get_or_add_pPr()
        num_pr = pPr.find(qn("w:numPr"))
        if num_pr is None:
            num_pr = OxmlElement("w:numPr")
            pPr.append(num_pr)
        else:
            num_pr.clear()
        ilvl_node = OxmlElement("w:ilvl")
        ilvl_node.set(qn("w:val"), str(ilvl))
        num_id_node = OxmlElement("w:numId")
        num_id_node.set(qn("w:val"), str(num_id))
        num_pr.append(ilvl_node)
        num_pr.append(num_id_node)
    except Exception:
        pass


def _ensure_heading_multilevel_numbering(
    document: Document,
    level1_format: str = "chinese",
    level2_format: str = "arabic",
    level3_format: str = "arabic",
    headings_spec: list[dict[str, Any]] | None = None,
) -> int:
    numbering_elm = _get_numbering_element(document, create=True)

    abstract_id = max(_numbering_max_id(numbering_elm, "abstractNum", "abstractNumId"), 0) + 1
    num_id = max(_numbering_max_id(numbering_elm, "num", "numId"), 0) + 1

    h1_style_id = document.styles["Heading 1"].style_id if "Heading 1" in document.styles else "Heading1"
    h2_style_id = document.styles["Heading 2"].style_id if "Heading 2" in document.styles else "Heading2"
    h3_style_id = document.styles["Heading 3"].style_id if "Heading 3" in document.styles else "Heading3"

    if level1_format == "chinese":
        l1_fmt = "chineseCountingThousand"
        l1_txt = "第%1章"
    elif level1_format == "arabic":
        l1_fmt = "decimal"
        l1_txt = "%1、"
    elif level1_format == "arabic-space":
        l1_fmt = "decimal"
        l1_txt = "%1"
    else:
        l1_fmt = "chineseCountingThousand"
        l1_txt = "第%1章"

    def _lvl_ind_xml(lvl_idx: int, default_size: float) -> str:
        h = next((item for item in (headings_spec or []) if int(item.get("level", 0)) == lvl_idx + 1), {})
        f_sz = float(h.get("font_size_pt") or default_size)
        fl_c = h.get("first_line_indent_chars")
        fl_m = h.get("first_line_indent_mm")
        l_c = h.get("left_indent_chars")
        l_m = h.get("left_indent_mm")

        attrs = []
        if fl_c is not None:
            fl_val = float(fl_c)
            attrs.append(f'w:firstLineChars="{int(round(fl_val * 100))}"')
            attrs.append(f'w:firstLine="{int(round(fl_val * f_sz * 20))}"')
        elif fl_m is not None and float(fl_m) > 0:
            chars = round(float(fl_m) / (f_sz * 25.4 / 72), 2)
            attrs.append(f'w:firstLineChars="{int(round(chars * 100))}"')
            attrs.append(f'w:firstLine="{int(round(float(fl_m) * 56.6929))}"')
        else:
            attrs.append('w:firstLine="0"')
            attrs.append('w:firstLineChars="0"')

        if l_c is not None:
            l_val = float(l_c)
            attrs.append(f'w:leftChars="{int(round(l_val * 100))}"')
            attrs.append(f'w:left="{int(round(l_val * f_sz * 20))}"')
        elif l_m is not None and float(l_m) > 0:
            chars = round(float(l_m) / (f_sz * 25.4 / 72), 2)
            attrs.append(f'w:leftChars="{int(round(chars * 100))}"')
            attrs.append(f'w:left="{int(round(float(l_m) * 56.6929))}"')
        else:
            attrs.append('w:left="0"')
            attrs.append('w:leftChars="0"')

        return f'<w:pPr><w:ind {" ".join(attrs)}/></w:pPr>'

    ind_l1 = _lvl_ind_xml(0, 16.0)
    ind_l2 = _lvl_ind_xml(1, 14.0)
    ind_l3 = _lvl_ind_xml(2, 12.0)

    abstract_xml = f'''
    <w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">
      <w:multiLevelType w:val="multilevel"/>
      <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="{l1_fmt}"/>
        <w:pStyle w:val="{h1_style_id}"/>
        <w:suff w:val="space"/>
        <w:lvlText w:val="{l1_txt}"/>
        <w:lvlJc w:val="left"/>
        {ind_l1}
        <w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>
      </w:lvl>
      <w:lvl w:ilvl="1">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:pStyle w:val="{h2_style_id}"/>
        <w:isLgl/>
        <w:suff w:val="space"/>
        <w:lvlText w:val="%1.%2"/>
        <w:lvlJc w:val="left"/>
        {ind_l2}
        <w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>
      </w:lvl>
      <w:lvl w:ilvl="2">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:pStyle w:val="{h3_style_id}"/>
        <w:isLgl/>
        <w:suff w:val="space"/>
        <w:lvlText w:val="%1.%2.%3"/>
        <w:lvlJc w:val="left"/>
        {ind_l3}
        <w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>
      </w:lvl>
    </w:abstractNum>
    '''
    numbering_elm.append(parse_xml(abstract_xml.strip()))
    num_xml = f'<w:num {nsdecls("w")} w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'
    numbering_elm.append(parse_xml(num_xml.strip()))

    for level, mode in ((1, level1_format), (2, level2_format), (3, level3_format)):
        if mode != "keep":
            _bind_style_numbering(document, f"Heading {level}", num_id, level - 1)
    return num_id


def normalize_heading_styles(
    document: Document,
    heading_numbering: dict[str, Any] | str | None = None,
    headings_spec: list[dict[str, Any]] | None = None,
) -> list[str]:
    changes: list[str] = []
    level1_style = get_or_create_paragraph_style(document, "Heading 1")
    level2_style = get_or_create_paragraph_style(document, "Heading 2")
    level3_style = get_or_create_paragraph_style(document, "Heading 3")
    title_style = get_or_create_paragraph_style(document, "Title")

    if isinstance(heading_numbering, str):
        level1_format = heading_numbering
        level2_format = "arabic"
        level3_format = "arabic"
    elif isinstance(heading_numbering, dict):
        level1_format = heading_numbering.get("level1", "keep")
        level2_format = heading_numbering.get("level2", "keep")
        level3_format = heading_numbering.get("level3", "keep")
    else:
        # A distilled template that has no heading numbering must preserve the
        # target's existing numbering instead of receiving a built-in default.
        level1_format = "keep"
        level2_format = "keep"
        level3_format = "keep"

    re_l1 = re.compile(r"^\s*(?:第[一二三四五六七八九十0-9]+章|\d+(?:[、\.\s]|\s*))\s*[^\s\d\.]")
    re_l2 = re.compile(r"^\s*\d+\.\d+(?:[、\.\s]|\s*)\s*[^\s\d\.]")
    re_l3 = re.compile(r"^\s*\d+\.\d+\.\d+(?:[、\.\s]|\s*)\s*[^\s\d\.]")
    re_special = re.compile(
        r"^\s*(?:摘\s*要|中文摘要|英文摘要|ABSTRACT|abstract|目\s*录|table\s+of\s+contents|contents|绪\s*论|引\s*言|前\s*言|结论|参考文献|参考资料|致\s*谢|acknowledg(?:e)?ments?|符号说明|符号表|术语表|symbols?|glossary|附\s*录|appendix)(?:\s*[\(（].*?[\)）])?\s*$",
        re.I,
    )

    re_arabic_chapter = re.compile(r"^\s*(\d+)(?:[、\.\s]|\s*)(.+)$")
    re_digit_chapter = re.compile(r"^\s*第\s*(\d+)\s*章(?:\s*|、)(.+)$")
    re_cn_chapter = re.compile(r"^\s*第\s*([一二三四五六七八九十0-9]+)\s*章(?:\s*|、)(.+)$")
    re_l2_pattern = re.compile(r"^\s*(\d+)\.(\d+)(?:[、\.\s]|\s*)(.+)$")
    re_l3_pattern = re.compile(r"^\s*(\d+)\.(\d+)\.(\d+)(?:[、\.\s]|\s*)(.+)$")

    num_id: int | None = None
    if level1_format != "keep" or level2_format != "keep" or level3_format != "keep":
        num_id = _ensure_heading_multilevel_numbering(
            document, level1_format, level2_format, level3_format, headings_spec
        )

    # Detect if the first heading before 摘要/ABSTRACT is the document/thesis title
    first_heading_idx = None
    abstract_idx = None
    paragraphs = list(document.paragraphs)
    for idx, p in enumerate(paragraphs):
        t = p.text.strip()
        if not t:
            continue
        c_style = p.style.name.lower() if p.style else ""
        is_h = c_style.startswith("heading") or c_style.startswith("标题") or c_style in {"title", "subtitle"}
        if is_h and first_heading_idx is None:
            first_heading_idx = idx
        if re_special.match(t) and ("摘要" in t or "ABSTRACT" in t.upper() or "摘" in t):
            abstract_idx = idx
            break

    if first_heading_idx is not None and abstract_idx is not None and first_heading_idx < abstract_idx:
        p_title = paragraphs[first_heading_idx]
        p_title_text = p_title.text.strip()
        if not re_l1.match(p_title_text) and not re_special.match(p_title_text):
            p_title.style = title_style
            _remove_paragraph_numbering(p_title)

    adjusted = 0
    converted_l1 = 0
    converted_l2 = 0
    converted_l3 = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        current_style = paragraph.style.name if paragraph.style else ""
        if current_style in TOC_MANAGED_STYLES or current_style.lower().startswith("wfm toc"):
            continue
        from document_structure import is_heading_candidate
        from word_format.properties import strip_text_prefix
        is_heading_or_candidate = is_heading_candidate(text, current_style)
        actual_level = next((level for level, pattern in ((3, re_l3), (2, re_l2), (1, re_l1)) if pattern.match(text)), None)
        if actual_level is None:
            style_match = re.fullmatch(r"Heading ([1-3])", current_style, re.I)
            actual_level = int(style_match.group(1)) if style_match else None
        if actual_level and {1: level1_format, 2: level2_format, 3: level3_format}[actual_level] == "keep":
            continue
        if current_style.lower() == "title":
            _remove_paragraph_numbering(paragraph)
            continue
        if re_special.match(text) and is_heading_or_candidate:
            if paragraph.style != level1_style:
                paragraph.style = level1_style
                adjusted += 1
            _remove_paragraph_numbering(paragraph)
        elif re_l3.match(text) and is_heading_or_candidate:
            if paragraph.style != level3_style:
                paragraph.style = level3_style
                adjusted += 1
            if level3_format == "arabic" and num_id is not None:
                m_l3 = re_l3_pattern.match(text)
                if m_l3:
                    clean_text = m_l3.group(4).strip().lstrip("、.．").strip()
                    strip_text_prefix(paragraph._p, paragraph.text.find(clean_text))
                    _set_paragraph_numbering(paragraph, num_id, ilvl=2)
                    converted_l3 += 1
                else:
                    _set_paragraph_numbering(paragraph, num_id, ilvl=2)
            elif current_style.lower() == "heading 3" and num_id is not None:
                _set_paragraph_numbering(paragraph, num_id, ilvl=2)
        elif re_l2.match(text) and is_heading_or_candidate:
            if paragraph.style != level2_style:
                paragraph.style = level2_style
                adjusted += 1
            if level2_format == "arabic" and num_id is not None:
                m_l2 = re_l2_pattern.match(text)
                if m_l2:
                    clean_text = m_l2.group(3).strip().lstrip("、.．").strip()
                    strip_text_prefix(paragraph._p, paragraph.text.find(clean_text))
                    _set_paragraph_numbering(paragraph, num_id, ilvl=1)
                    converted_l2 += 1
                else:
                    _set_paragraph_numbering(paragraph, num_id, ilvl=1)
            elif current_style.lower() == "heading 2" and num_id is not None:
                _set_paragraph_numbering(paragraph, num_id, ilvl=1)
        elif re_l1.match(text) and is_heading_or_candidate:
            if paragraph.style != level1_style:
                paragraph.style = level1_style
                adjusted += 1

            m_cn = re_cn_chapter.match(text)
            m_dig = re_digit_chapter.match(text)
            m_ara = re_arabic_chapter.match(text)
            rest = ""
            if m_cn:
                rest = m_cn.group(2).strip().lstrip("、.．").strip()
            elif m_dig:
                rest = m_dig.group(2).strip().lstrip("、.．").strip()
            elif m_ara:
                rest = m_ara.group(2).strip().lstrip("、.．").strip()

            if rest and num_id is not None and level1_format != "keep":
                strip_text_prefix(paragraph._p, paragraph.text.find(rest))
                _set_paragraph_numbering(paragraph, num_id, ilvl=0)
                converted_l1 += 1
            elif num_id is not None and level1_format != "keep" and current_style.lower() == "heading 1":
                _set_paragraph_numbering(paragraph, num_id, ilvl=0)
        elif current_style.lower() == "heading 1" and num_id is not None and level1_format != "keep":
            if re_special.match(text):
                _remove_paragraph_numbering(paragraph)
            else:
                _set_paragraph_numbering(paragraph, num_id, ilvl=0)
        elif current_style.lower() == "heading 2" and num_id is not None and level2_format != "keep":
            _set_paragraph_numbering(paragraph, num_id, ilvl=1)
        elif current_style.lower() == "heading 3" and num_id is not None and level3_format != "keep":
            _set_paragraph_numbering(paragraph, num_id, ilvl=2)

    if adjusted:
        changes.append(f"Normalized {adjusted} heading paragraph style(s) according to chapter/section numbering")
    if converted_l1:
        changes.append(f"Bound {converted_l1} chapter heading(s) to native Word '{level1_format}' multilevel numbering")
    if converted_l2:
        changes.append(f"Bound {converted_l2} level-2 section heading(s) to native Word multilevel numbering")
    if converted_l3:
        changes.append(f"Bound {converted_l3} level-3 subsection heading(s) to native Word multilevel numbering")
    return changes


def apply_page(document: Document, page: dict[str, Any]) -> list[str]:
    changes: list[str] = []
    for section in document.sections:
        orientation = page.get("orientation")
        size = str(page.get("size") or "")
        preset_size = PAGE_SIZES_MM.get(size)
        width, height = preset_size if preset_size else (page.get("width_mm"), page.get("height_mm"))
        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            if width and height:
                section.page_width = Mm(max(float(width), float(height)))
                section.page_height = Mm(min(float(width), float(height)))
        elif orientation == "portrait":
            section.orientation = WD_ORIENT.PORTRAIT
            if width and height:
                section.page_width = Mm(min(float(width), float(height)))
                section.page_height = Mm(max(float(width), float(height)))
        elif width and height:
            section.page_width = Mm(float(width))
            section.page_height = Mm(float(height))
        mapping = {
            "margin_top_mm": "top_margin",
            "margin_right_mm": "right_margin",
            "margin_bottom_mm": "bottom_margin",
            "margin_left_mm": "left_margin",
            "gutter_mm": "gutter",
            "header_distance_mm": "header_distance",
            "footer_distance_mm": "footer_distance",
        }
        for token, attr in mapping.items():
            if page.get(token) is not None:
                setattr(section, attr, Mm(float(page[token])))
    if document.sections and page:
        changes.append(f"Applied page geometry to {len(document.sections)} section(s)")
    return changes


def iter_table_paragraphs(table) -> Iterable:
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from iter_table_paragraphs(nested)


def iter_all_paragraphs(document: Document) -> Iterable:
    yield from document.paragraphs
    for table in document.tables:
        yield from iter_table_paragraphs(table)
    for section in document.sections:
        for container in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            yield from container.paragraphs
            for table in container.tables:
                yield from iter_table_paragraphs(table)




def apply_reference_style(document: Document, body: dict[str, Any], references: dict[str, Any]) -> list[str]:
    if not references:
        return []
    reference_tokens = deep_merge(body, references)
    hanging_indent = reference_tokens.pop("hanging_indent_mm", None)
    if hanging_indent is not None:
        reference_tokens["left_indent_mm"] = float(reference_tokens.get("left_indent_mm", 0.0))
        reference_tokens["left_indent_chars"] = float(reference_tokens.get("left_indent_chars", 0.0))
        reference_tokens["first_line_indent_mm"] = -float(hanging_indent)
        reference_tokens["hanging_indent_mm"] = float(hanging_indent)
    bibliography = resolve_paragraph_style(document, references, "Bibliography")
    if bibliography.base_style is None:
        bibliography.base_style = get_or_create_paragraph_style(document, "Normal")
    set_style_font(bibliography, reference_tokens)
    set_paragraph_format(bibliography, reference_tokens)

    numbering_mode = references.get("numbering_mode")
    num_id = None
    if numbering_mode == "word-numbering":
        style_key = "decimal-bracket"
        num_fmt, level_text = LIST_NUMBERING_STYLES[style_key]
        numbering = _get_numbering_element(document, create=True)
        abstract_id = _numbering_max_id(numbering, "abstractNum", "abstractNumId") + 1
        num_id = _numbering_max_id(numbering, "num", "numId") + 1

        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        nsid = OxmlElement("w:nsid")
        nsid.set(qn("w:val"), f"{(abstract_id * 2654435761) & 0xFFFFFFFF:08X}")
        abstract.append(nsid)
        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "singleLevel")
        abstract.append(multi)
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), "0")
        start_node = OxmlElement("w:start")
        start_node.set(qn("w:val"), "1")
        lvl.append(start_node)
        num_fmt_node = OxmlElement("w:numFmt")
        num_fmt_node.set(qn("w:val"), num_fmt)
        lvl.append(num_fmt_node)
        text_node = OxmlElement("w:lvlText")
        text_node.set(qn("w:val"), level_text)
        lvl.append(text_node)
        justification = OxmlElement("w:lvlJc")
        justification.set(qn("w:val"), "left")
        lvl.append(justification)

        pPr_lvl = OxmlElement("w:pPr")
        ind_lvl = OxmlElement("w:ind")
        ind_lvl.set(qn("w:left"), "0")
        ind_lvl.set(qn("w:leftChars"), "0")
        ind_lvl.set(qn("w:hanging"), "420")
        ind_lvl.set(qn("w:hangingChars"), "200")
        pPr_lvl.append(ind_lvl)
        lvl.append(pPr_lvl)

        rPr_lvl = OxmlElement("w:rPr")
        rFonts_lvl = OxmlElement("w:rFonts")
        latin = reference_tokens.get("font_latin") or "Times New Roman"
        east_asia = reference_tokens.get("font_east_asia") or "宋体"
        rFonts_lvl.set(qn("w:ascii"), latin)
        rFonts_lvl.set(qn("w:hAnsi"), latin)
        rFonts_lvl.set(qn("w:eastAsia"), east_asia)
        rPr_lvl.append(rFonts_lvl)
        size_pt = float(reference_tokens.get("font_size_pt") or 10.5)
        sz_lvl = OxmlElement("w:sz")
        sz_lvl.set(qn("w:val"), str(int(round(size_pt * 2))))
        rPr_lvl.append(sz_lvl)
        lvl.append(rPr_lvl)

        abstract.append(lvl)
        numbering.append(abstract)

        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_ref = OxmlElement("w:abstractNumId")
        abstract_ref.set(qn("w:val"), str(abstract_id))
        num.append(abstract_ref)
        numbering.append(num)

    heading_pattern = re.compile(
        r"^\s*(?:references|bibliography|works cited|参考文献|参考资料)(?:\s*[\(（].*?[\)）])?\s*$",
        re.I,
    )
    known_style_names = {
        bibliography.name.lower(),
        "bibliography",
        "references",
        "reference",
        "参考文献",
    }
    in_reference_section = False
    styled = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        current_style = paragraph.style.name.lower() if paragraph.style else ""
        if heading_pattern.match(text):
            in_reference_section = True
            continue
        if in_reference_section and current_style.startswith("heading"):
            in_reference_section = False
        if not text:
            continue
        if in_reference_section or current_style in known_style_names:
            paragraph.style = bibliography
            if numbering_mode == "word-numbering" and num_id is not None:
                clean_text = re.sub(r"^\s*\[\s*\d+\s*\]\s*|^\s*\d+[\.、]\s*", "", text).strip()
                from word_format.properties import strip_text_prefix
                strip_text_prefix(paragraph._p, paragraph.text.find(clean_text))
                for run in paragraph.runs:
                    set_run_font(run, reference_tokens)
                _set_paragraph_numbering(paragraph, num_id, ilvl=0)
            elif numbering_mode == "none":
                _remove_paragraph_numbering(paragraph)
            else:
                for run in paragraph.runs:
                    set_run_font(run, reference_tokens)
            styled += 1
    return [f"Updated Bibliography style and assigned it to {styled} reference paragraph(s)"]


def _numbering_max_id(numbering, tag: str, attribute: str) -> int:
    values = []
    for element in numbering.findall(qn(f"w:{tag}")):
        value = element.get(qn(f"w:{attribute}"))
        if value is not None and str(value).isdigit():
            values.append(int(value))
    return max(values, default=-1)


def _num_format_for_num_id(numbering, num_id: str | None) -> str | None:
    if not num_id:
        return None
    num = next((item for item in numbering.findall(qn("w:num")) if item.get(qn("w:numId")) == num_id), None)
    if num is None:
        return None
    abstract_id = num.find(qn("w:abstractNumId"))
    if abstract_id is None:
        return None
    abstract = next(
        (item for item in numbering.findall(qn("w:abstractNum")) if item.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val"))),
        None,
    )
    level = abstract.find(qn("w:lvl")) if abstract is not None else None
    num_fmt = level.find(qn("w:numFmt")) if level is not None else None
    return num_fmt.get(qn("w:val")) if num_fmt is not None else None


def _get_numbering_element(document: Document, create: bool = False):
    try:
        return document.part.numbering_part.element
    except (KeyError, NotImplementedError):
        if not create:
            return None
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.oxml import parse_xml
        from docx.parts.numbering import NumberingPart

        numbering_elm = parse_xml(b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        part = NumberingPart(
            PackURI('/word/numbering.xml'),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
            numbering_elm,
            document.part.package,
        )
        document.part.relate_to(part, RT.NUMBERING)
        return numbering_elm


def _has_numbered_list_marker(document: Document, paragraph) -> bool:
    style_name = paragraph.style.name.lower() if paragraph.style else ""
    if (
        style_name.startswith("heading")
        or style_name.startswith("标题")
        or style_name in {"title", "subtitle", "toc", "wfm toc heading", "bibliography", "references", "reference", "参考文献"}
    ):
        return False
    numbering = _get_numbering_element(document, create=False)
    ppr = paragraph._p.pPr
    num_pr = ppr.find(qn("w:numPr")) if ppr is not None else None
    if num_pr is not None and numbering is not None:
        num_id = num_pr.find(qn("w:numId"))
        fmt = _num_format_for_num_id(numbering, num_id.get(qn("w:val")) if num_id is not None else None)
        return fmt not in {None, "bullet", "picture"}
    return any(marker in style_name for marker in ("list number", "numbered", "编号", "list paragraph"))


def _append_numbering_definition(document: Document, style: str, start: int) -> int:
    num_fmt, level_text = LIST_NUMBERING_STYLES[style]
    numbering = _get_numbering_element(document, create=True)
    abstract_id = _numbering_max_id(numbering, "abstractNum", "abstractNumId") + 1
    num_id = _numbering_max_id(numbering, "num", "numId") + 1
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    nsid = OxmlElement("w:nsid")
    nsid.set(qn("w:val"), f"{(abstract_id * 2654435761) & 0xFFFFFFFF:08X}")
    abstract.append(nsid)
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start_node = OxmlElement("w:start")
    start_node.set(qn("w:val"), str(start))
    lvl.append(start_node)
    num_fmt_node = OxmlElement("w:numFmt")
    num_fmt_node.set(qn("w:val"), num_fmt)
    lvl.append(num_fmt_node)
    text_node = OxmlElement("w:lvlText")
    text_node.set(qn("w:val"), level_text)
    lvl.append(text_node)
    justification = OxmlElement("w:lvlJc")
    justification.set(qn("w:val"), "left")
    lvl.append(justification)
    paragraph_properties = OxmlElement("w:pPr")
    indentation = OxmlElement("w:ind")
    indentation.set(qn("w:left"), "720")
    indentation.set(qn("w:hanging"), "360")
    paragraph_properties.append(indentation)
    lvl.append(paragraph_properties)
    abstract.append(lvl)
    numbering.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_reference = OxmlElement("w:abstractNumId")
    abstract_reference.set(qn("w:val"), str(abstract_id))
    num.append(abstract_reference)
    if start != 1:
        override = OxmlElement("w:lvlOverride")
        override.set(qn("w:ilvl"), "0")
        override_start = OxmlElement("w:startOverride")
        override_start.set(qn("w:val"), str(start))
        override.append(override_start)
        num.append(override)
    numbering.append(num)
    return num_id


def _set_managed_num_pr(paragraph, num_id: int) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn("w:numPr"))
    if old is not None:
        ppr.remove(old)
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num_pr.append(ilvl)
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.append(num)
    ppr.append(num_pr)


def apply_numbered_lists(document: Document, tokens: dict[str, Any]) -> list[str]:
    numbered = (tokens or {}).get("numbered", {})
    if not numbered:
        return []
    style = str(numbered.get("style") or "decimal-period")
    if style not in LIST_NUMBERING_STYLES:
        raise ValueError(f"lists.numbered.style is unsupported: {style}")
    start = int(numbered.get("start", 1))
    left_indent = numbered.get("left_indent_mm")
    hanging_indent = numbered.get("hanging_indent_mm")
    if start < 1:
        raise ValueError("lists.numbered.start must be at least 1")
    paragraphs = list(document.paragraphs)
    targets = [paragraph for paragraph in paragraphs if _has_numbered_list_marker(document, paragraph)]
    if not targets:
        return [f"Registered managed list numbering style {style}; no numbered list paragraphs were found"]
    changes: list[str] = []
    target_ids = {id(paragraph) for paragraph in targets}
    numbering = _get_numbering_element(document, create=True)
    index = 0
    while index < len(paragraphs):
        if id(paragraphs[index]) not in target_ids:
            index += 1
            continue
        end = index
        while end + 1 < len(paragraphs) and id(paragraphs[end + 1]) in target_ids:
            end += 1
        num_id = _append_numbering_definition(document, style, start)
        for paragraph in paragraphs[index : end + 1]:
            _set_managed_num_pr(paragraph, num_id)
            p_format = paragraph.paragraph_format
            if left_indent is not None:
                p_format.left_indent = Mm(float(left_indent))
            if hanging_indent is not None:
                p_format.first_line_indent = Mm(-float(hanging_indent))
        changes.append(f"Applied managed {style} numbering to {end - index + 1} list paragraph(s)")
        index = end + 1
    return changes






def validate_spec(spec: dict[str, Any]) -> list[str]:
    from word_format.spec import validate_spec as validate_contract
    try:
        validate_contract(spec)
        return []
    except (ValueError, TypeError, OverflowError, AttributeError) as exc:
        return [str(exc)]


def apply_document(
    input_path: str | Path,
    output_path: str | Path,
    spec: dict[str, Any],
    *,
    clear_direct: bool = False,
    verification: dict[str, Any] | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    from word_format.api import apply_spec as execute_spec
    report = execute_spec(
        input_path, output_path, normalize_spec(spec),
        clear_direct=clear_direct, verification=verification, overwrite=overwrite,
    )
    if verification and verification.get("visual_execution_mode") == "dashboard-direct":
        report["verification"]["visual"] = {
            "status": "skipped",
            "reason": "本地下载模式未连接视觉 AI，请运行显式渲染验收",
        }
    report["clear_direct_font_formatting"] = clear_direct
    return report


def main() -> int:
    from cli_output import configure_cli_output
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source DOCX")
    parser.add_argument("output", help="Destination DOCX; must not equal input")
    parser.add_argument("--handoff", required=True, help="Confirmed Dashboard HANDOFF.json")
    parser.add_argument(
        "--spec",
        help="Task-local spec that may add validated citations to the confirmed handoff",
    )
    parser.add_argument("--report", help="Write application report JSON")
    args = parser.parse_args()
    try:
        handoff = load_confirmed_handoff(args.handoff, args.input, args.output)
        if args.spec:
            spec = normalize_spec(load_json(args.spec))
            if not isinstance(spec, dict):
                raise ValueError("Task-local specification must be a JSON object")
            if _without_citations(spec) != _without_citations(handoff["spec"]):
                raise ValueError("Task-local specification may only add validated citations to the confirmed handoff")
        else:
            spec = copy.deepcopy(handoff["spec"])
        report = apply_document(
            args.input,
            args.output,
            spec,
            clear_direct=bool(handoff.get("clear_direct_font_formatting", False)),
            verification=handoff.get("verification") if isinstance(handoff.get("verification"), dict) else None,
        )
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        Path(args.report).write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
