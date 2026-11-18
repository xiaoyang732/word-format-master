"""Complex OOXML writers owned by the format core and invoked via its registry."""

from __future__ import annotations

import base64
import copy
import json
import re
from typing import Any, Iterable

from config import LIST_NUMBERING_STYLES
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from lxml import etree

from .constants import (
    BIBLIOGRAPHY_STYLE_ALIASES,
    DEFAULT_CAPTION_LABELS,
    DEFAULT_CHINESE_BASE_CONFIG,
    DEFAULT_LIST_HANGING_INDENT_TWIPS,
    DEFAULT_LIST_LEFT_INDENT_TWIPS,
    STYLE_BIBLIOGRAPHY,
    STYLE_WFM_PAGE_NUMBER,
    STYLE_WFM_TOC_BREAK,
    STYLE_WFM_TOC_FIELD,
    STYLE_WFM_TOC_HEADING,
    THREE_LINE_HEADER_BORDER_EIGHTH_PT,
    THREE_LINE_OUTER_BORDER_EIGHTH_PT,
    TOC_MANAGED_STYLES,
    UniversalChineseBaseConfig,
    build_caption_pattern,
    is_builtin_paragraph_style,
    section_content_width_twips,
    wfm_header_footer_style,
)
from .namespaces import qn
from .properties import paragraph_style_name
from .ooxml import insert_child

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def get_or_create_paragraph_style(document: Document, name: str):
    try:
        return document.styles[name]
    except KeyError:
        return document.styles.add_style(
            name, WD_STYLE_TYPE.PARAGRAPH, builtin=is_builtin_paragraph_style(name)
        )


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
    from .properties import element, write_font
    from .spec import FONT, token_operations

    for action, params in token_operations({k: v for k, v in tokens.items() if k in FONT}):
        key = next(k for k, v in FONT.items() if v == action)
        write_font(element(style.element, "rPr"), key, params["value"])


def set_paragraph_format(style, tokens: dict[str, Any], document=None) -> None:
    from .properties import element, write_effective_tabs, write_paragraph
    from .spec import INDENTS, PARA, token_operations

    selected = {
        k: v for k, v in tokens.items() if k in PARA or any(k.startswith(n + "_") for n in INDENTS)
    }
    if "font_size_pt" in tokens:
        selected["font_size_pt"] = tokens["font_size_pt"]
    for action, params in token_operations(selected):
        if not action.startswith("paragraph."):
            continue
        key = action[len("paragraph.") : -len(".set")]
        key = {"spacing_before": "space_before_pt", "spacing_after": "space_after_pt"}.get(key, key)
        value = params if key == "line_spacing" or key.endswith("indent") else params["value"]
        node = style._p if hasattr(style, "_p") else style.element
        if key == "tabs":
            if hasattr(style, "_p"):
                document = document or style.part.document
                paragraph = node
            else:
                if document is None:
                    from .contracts import FormatError

                    raise FormatError("Replacing style tabs requires the owning document")
                from .inspect import _style_paragraph

                paragraph = _style_paragraph(style.name, document)
            write_effective_tabs(document, paragraph, element(node, "pPr"), value)
            continue
        write_paragraph(element(node, "pPr"), key, value)


def set_run_font(run, tokens: dict[str, Any]) -> None:
    from .properties import element, write_font
    from .spec import FONT, token_operations

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
    from .fields import complex_field_runs

    for run in complex_field_runs(instruction, cached_text):
        paragraph._p.append(run)


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
            insert_child(tblPr, parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
        elif alignment in ("left", "right"):
            existing_jc = tblPr.find(qn("w:jc"))
            if existing_jc is not None:
                tblPr.remove(existing_jc)
            insert_child(tblPr, parse_xml(f'<w:jc {nsdecls("w")} w:val="{alignment}"/>'))

        snapshot_key = "{urn:word-format-master:format}originalBorders"
        cells = [tc for tr in table._tbl.findall(qn("w:tr")) for tc in tr.findall(qn("w:tc"))]
        if three_line and snapshot_key not in tblPr.attrib:
            border_nodes = [tblPr] + cells
            snapshot = []
            seen_cells = set()
            for node in border_nodes:
                owner = node if node is tblPr else node.get_or_add_tcPr()
                if owner in seen_cells:
                    continue
                seen_cells.add(owner)
                border = owner.find(qn("w:tblBorders" if node is tblPr else "w:tcBorders"))
                snapshot.append(
                    base64.b64encode(etree.tostring(border)).decode()
                    if border is not None
                    else None
                )
            tblPr.set(snapshot_key, json.dumps(snapshot))
        if "three_line_table" in tables and not three_line and snapshot_key in tblPr.attrib:
            snapshot = json.loads(tblPr.attrib.pop(snapshot_key))
            owners = [tblPr] + [tc.get_or_add_tcPr() for tc in cells]
            if len(snapshot) != len(owners):
                # Older snapshots used the grid view, which omits continuation
                # cells of vertical merges. Those cells were never changed.
                legacy = [tblPr]
                for row in table.rows:
                    for cell in row.cells:
                        owner = cell._tc.get_or_add_tcPr()
                        if owner not in legacy:
                            legacy.append(owner)
                if len(snapshot) == len(legacy):
                    owners = legacy
            if len(snapshot) != len(owners):
                raise ValueError("Table structure changed; saved borders cannot be restored")
            for index, (owner, saved) in enumerate(zip(owners, snapshot)):
                border = owner.find(qn("w:tblBorders" if index == 0 else "w:tcBorders"))
                if border is not None:
                    owner.remove(border)
                if saved:
                    insert_child(owner, parse_xml(base64.b64decode(saved)))
        if three_line:
            existing_borders = tblPr.find(qn("w:tblBorders"))
            if existing_borders is not None:
                tblPr.remove(existing_borders)
            three_line_xml = parse_xml(
                f"<w:tblBorders {nsdecls('w')}>"
                f'<w:top w:val="single" w:sz="{THREE_LINE_OUTER_BORDER_EIGHTH_PT}" w:space="0" w:color="auto"/>'
                f'<w:left w:val="none"/>'
                f'<w:bottom w:val="single" w:sz="{THREE_LINE_OUTER_BORDER_EIGHTH_PT}" w:space="0" w:color="auto"/>'
                f'<w:right w:val="none"/>'
                f'<w:insideH w:val="none"/>'
                f'<w:insideV w:val="none"/>'
                f"</w:tblBorders>"
            )
            insert_child(tblPr, three_line_xml)

            rows = table._tbl.findall(qn("w:tr"))
            for row_index, row in enumerate(rows):
                for cell in row.findall(qn("w:tc")):
                    tcPr = cell.get_or_add_tcPr()
                    existing_tc_borders = tcPr.find(qn("w:tcBorders"))
                    if existing_tc_borders is not None:
                        tcPr.remove(existing_tc_borders)
                    tc_border_xml = OxmlElement("w:tcBorders")
                    # Direct cell borders override table and conditional-style
                    # borders. Cover actual cells, including merge continuations.
                    for edge in (
                        "top",
                        "start",
                        "left",
                        "bottom",
                        "end",
                        "right",
                        "insideH",
                        "insideV",
                        "tl2br",
                        "tr2bl",
                    ):
                        border = OxmlElement("w:" + edge)
                        size = (
                            THREE_LINE_OUTER_BORDER_EIGHTH_PT
                            if (edge == "top" and row_index == 0)
                            or (edge == "bottom" and row_index == len(rows) - 1)
                            else THREE_LINE_HEADER_BORDER_EIGHTH_PT
                            if edge == "bottom" and row_index == 0
                            else None
                        )
                        border.set(qn("w:val"), "single" if size else "nil")
                        if size:
                            border.set(qn("w:sz"), str(size))
                            border.set(qn("w:space"), "0")
                            border.set(qn("w:color"), "auto")
                        tc_border_xml.append(border)
                    insert_child(tcPr, tc_border_xml)

        if "repeat_header_row" in tables and not repeat_header and table.rows:
            trPr = table.rows[0]._tr.trPr
            header = trPr.find(qn("w:tblHeader")) if trPr is not None else None
            if header is not None:
                trPr.remove(header)
        if repeat_header and table.rows:
            header_tr = table.rows[0]._tr
            trPr = header_tr.get_or_add_trPr()
            if trPr.find(qn("w:tblHeader")) is None:
                trPr.append(parse_xml(f"<w:tblHeader {nsdecls('w')}/>"))

        if latin or east_asia or font_size_pt is not None:
            from .properties import element, write_font

            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        ppr_rpr = element(element(paragraph._p, "pPr"), "rPr")
                        if font_size_pt is not None:
                            write_font(ppr_rpr, "font_size_pt", float(font_size_pt))
                        if latin:
                            write_font(ppr_rpr, "font_latin", latin)
                        if east_asia:
                            write_font(ppr_rpr, "font_east_asia", east_asia)
                        for run in paragraph.runs:
                            rpr = element(run._r, "rPr")
                            if font_size_pt is not None:
                                write_font(rpr, "font_size_pt", float(font_size_pt))
                            if latin:
                                write_font(rpr, "font_latin", latin)
                            if east_asia:
                                write_font(rpr, "font_east_asia", east_asia)

    desc = []
    if three_line:
        desc.append("three-line borders")
    if repeat_header:
        desc.append("repeat header row")
    if alignment:
        desc.append(f"alignment={alignment}")
    if font_size_pt:
        desc.append(f"font_size={font_size_pt}pt")
    changes.append(
        f"Applied table formatting ({', '.join(desc)}) to {len(document.tables)} table(s)"
    )
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


def _move_caption_next_to_target(
    document: Document, caption_element, target_element, position: str
) -> bool:
    if target_element is None or target_element is caption_element:
        return False
    body = document.element.body
    children = list(body)
    caption_index = children.index(caption_element)
    target_index = children.index(target_element)
    desired_index = target_index if position == "above" else target_index + 1
    if (
        caption_index == desired_index
        or (position == "above" and caption_index == target_index - 1)
        or (position == "below" and caption_index == target_index + 1)
    ):
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
    paragraphs = list(document.paragraphs)
    for role in ("figure", "table"):
        tokens = captions.get(role) or {}
        position = tokens.get("position")
        if position not in {"above", "below"}:
            continue
        labels = [str(tokens.get("label") or "")] + list(DEFAULT_CAPTION_LABELS[role])
        pattern = build_caption_pattern(labels)
        moved = 0
        for paragraph in paragraphs:
            if paragraph._p.getparent() is not document.element.body or not pattern.match(
                paragraph.text or ""
            ):
                continue
            previous = _neighboring_content_element(document, paragraph._p, -1)
            following = _neighboring_content_element(document, paragraph._p, 1)
            matches_role = (
                (lambda elem: elem.tag == qn("w:tbl"))
                if role == "table"
                else _is_figure_paragraph
            )
            primary = following if position == "above" else previous
            secondary = previous if position == "above" else following
            target = primary if (primary is not None and matches_role(primary)) else None
            if target is None and secondary is not None and matches_role(secondary):
                target = secondary
            if _move_caption_next_to_target(document, paragraph._p, target, position):
                moved += 1
            if target is not None:
                from .properties import element, write_paragraph

                owner = paragraph._p if position == "above" else target if role == "figure" else None
                if owner is not None:
                    write_paragraph(element(owner, "pPr"), "keep_with_next", True)
        if moved:
            changes.append(
                f"Moved {moved} {role} caption paragraph(s) to the {position} of adjacent object(s)"
            )
        else:
            changes.append(
                f"Checked {role} caption placement; no safe adjacent object move was required"
            )
    return changes


def caption_keep_nodes(document, params):
    """Only the proven caption/object pair may receive keepNext."""
    label = params.get("label") or DEFAULT_CAPTION_LABELS[params["role"]][0]
    pattern = build_caption_pattern([label])
    nodes = set()
    for paragraph in document.paragraphs:
        if not pattern.match(paragraph.text):
            continue
        for direction in (-1, 1):
            neighbor = _neighboring_content_element(document, paragraph._p, direction)
            if neighbor is None:
                continue
            matched = (neighbor.tag == qn("w:tbl") if params["role"] == "table"
                       else _is_figure_paragraph(neighbor))
            if matched:
                if params["position"] == "above":
                    nodes.add(paragraph._p)
                elif params["role"] == "figure":
                    nodes.add(neighbor)
                break
    return nodes


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
            if (paragraph_style_name(document, paragraph._p) in style_names) or (
                "WFM Header" in style_names
                and paragraph._p.get("{urn:word-format-master}managed-header") == "true"
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


def _standalone_page_number_paragraph(container):
    from .fields import page_number_paragraphs

    nodes = set(page_number_paragraphs(container._element))
    for paragraph in container.paragraphs:
        if paragraph._p in nodes:
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
        changes.append(
            f"Updated first-page header/footer behavior in {len(document.sections)} section(s)"
        )
    if tokens.get("different_odd_even") is not None and hasattr(
        document.settings, "odd_and_even_pages_header_footer"
    ):
        document.settings.odd_and_even_pages_header_footer = bool(tokens["different_odd_even"])
        changes.append("Updated odd/even header/footer behavior")
    for kind in ("header", "footer"):
        config = tokens.get(kind) or {}
        if config.get("enabled") is False or (
            config.get("enabled") and not str(config.get("text", "")).strip()
        ):
            cleared = clear_managed_header_footer_paragraphs(
                document, {wfm_header_footer_style(kind)}
            )
            changes.append(f"Removed managed {kind} text from {cleared} part(s)")
            continue
        if not config.get("enabled"):
            continue
        if kind == "header":
            from .headers import content_params, set_content
            from .inspect import DocumentIndex
            from .registry import REGISTRY
            from .spec import token_operations

            index = DocumentIndex(document)
            for oid, record in index.records.items():
                if record["kind"] != "section":
                    continue
                ref = index.ref(oid)
                set_content(index, ref, content_params({"mode": "text", "text": config["text"]}))
                for action, params in token_operations(
                    {k: v for k, v in config.items() if k not in {"enabled", "text"}}
                ):
                    action = "header." + (
                        action if action.startswith("font.") else action.removeprefix("paragraph.")
                    )
                    cap = REGISTRY[action]
                    cap.handler(index, ref, cap.validator(params))
            changes.append("Updated header through registered section header handlers")
            continue
        style_name = wfm_header_footer_style(kind)
        style = get_or_create_paragraph_style(document, style_name)
        set_style_font(style, config)
        set_paragraph_format(style, config, document)
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


def _find_body_section_index(document: Document) -> int:
    from .body_sections import body_section_index

    return body_section_index(document)


def apply_page_numbers(document: Document, tokens: dict[str, Any]) -> list[str]:
    if not tokens:
        return []
    if not tokens.get("enabled"):
        cleared = clear_managed_header_footer_paragraphs(document, {STYLE_WFM_PAGE_NUMBER})
        return [f"Removed managed page-number fields from {cleared} part(s)"]
    from .contracts import FormatError
    from .fields import page_number_paragraphs
    from .headers import detach
    from .properties import element

    if tokens.get("scope") == "body" and "start_section" not in tokens:
        from .body_sections import ensure_body_boundary

        ensure_body_boundary(document)
    sections = list(document.sections)
    body_idx = _find_body_section_index(document)
    start_idx = int(tokens["start_section"]) - 1 if "start_section" in tokens else (
        body_idx if tokens.get("scope", "all") == "body" else 0
    )
    if not 0 <= start_idx < len(sections):
        raise FormatError("start_section is outside the document")
    if (tokens.get("scope") == "body" and "start_section" not in tokens
            and body_idx == 0 and len(sections) > 1):
        raise FormatError("Body section is ambiguous; specify start_section", "needs_clarification")
    location = tokens.get("location", "footer")
    style_name = STYLE_WFM_PAGE_NUMBER
    style = get_or_create_paragraph_style(document, style_name)
    set_style_font(style, tokens)
    set_paragraph_format(style, tokens, document)
    start = int(tokens.get("start", 1))

    # Freeze all variants before changing linked definitions, including explicit
    # references sharing the same part. First/even front matter must stay blank.
    variants = (location, f"first_page_{location}", f"even_page_{location}")
    for section in reversed(sections):
        for attr in variants:
            detach(getattr(section, attr))
    cleared = 0
    for section in sections[:start_idx]:
        for attr in variants:
            container = getattr(section, attr)
            for paragraph in page_number_paragraphs(container._element):
                paragraph.getparent().remove(paragraph)
                cleared += 1
    reset_idx = start_idx or (body_idx if "show_on_body_first_page" in tokens else 0)
    for i, section in enumerate(sections[start_idx:], start_idx):
        page_type = element(section._sectPr, "pgNumType")
        page_type.set(qn("w:fmt"), "decimal")
        if i in {start_idx, reset_idx}:
            page_type.set(qn("w:start"), str(start))
        else:
            page_type.attrib.pop(qn("w:start"), None)
    containers = []
    for section in sections[start_idx:]:
        containers.append(getattr(section, location))
        if document.settings.odd_and_even_pages_header_footer:
            containers.append(getattr(section, f"even_page_{location}"))
    for i, section in enumerate(sections[start_idx:], start_idx):
        show_first = tokens.get("show_on_first_page", True)
        if i >= body_idx and "show_on_body_first_page" in tokens:
            show_first = tokens["show_on_body_first_page"]
        if not show_first:
            section.different_first_page_header_footer = True
            container = getattr(section, f"first_page_{location}")
            # Freeze the inherited content before other variants receive PAGE
            # fields; retain images, relationships and unrelated footer text.
            for paragraph in page_number_paragraphs(container._element):
                paragraph.getparent().remove(paragraph)
        elif section.different_first_page_header_footer:
            containers.append(getattr(section, f"first_page_{location}"))
    seen: set[int] = set()
    reused = 0
    for container in containers:
        identity = id(container._element)
        if identity in seen:
            continue
        seen.add(identity)
        owned = page_number_paragraphs(container._element)
        # Collapse duplicate imported page-only paragraphs deterministically.
        for extra in owned[1:]:
            extra.getparent().remove(extra)
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


def _remove_managed_toc(document: Document) -> int:
    removed = 0
    for paragraph in list(document.paragraphs):
        if paragraph_style_name(document, paragraph._p) in TOC_MANAGED_STYLES:
            if paragraph._p.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
                continue
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
    re_abstract_cn = re.compile(
        r"^\s*(?:[【\[]\s*(?:摘\s*要|中文摘要|内容摘要|提\s*要)\s*[】\]]\s*(?:[:：]\s*)?.*|"
        r"(?:摘\s*要|中文摘要|内容摘要|提\s*要)(?:\s*[:：].*)?)$",
        re.I,
    )
    re_abstract_en = re.compile(
        r"^\s*(?:[【\[]\s*(?:abstracts?|英文摘要)\s*[】\]]\s*(?:[:：]\s*)?.*|"
        r"(?:abstracts?|英文摘要)(?:\s*[:：].*)?)$",
        re.I,
    )
    re_keywords_cn = re.compile(
        r"^\s*(?:[【\[]\s*(?:关\s*键\s*词|主\s*题\s*词)\s*[】\]]\s*(?:[:：]\s*)?.*|"
        r"(?:关\s*键\s*词|主\s*题\s*词)(?:\s*[:：].*)?)$",
        re.I,
    )
    re_keywords_en = re.compile(
        r"^\s*(?:[【\[]\s*(?:key\s*words?|index\s+terms?)\s*[】\]]\s*(?:[:：]\s*)?.*|"
        r"(?:key\s*words?|index\s+terms?)(?:\s*[:：].*)?)$",
        re.I,
    )
    re_toc = re.compile(r"^\s*(?:目\s*录|目\s*次|table\s+of\s+contents|contents)\s*$", re.I)
    re_next_module = re.compile(
        r"^\s*(?:第\s*[一二三四五六七八九十0-9]+\s*章|\d+(?:[、.\s]|\s*)[^\s\d.]|"
        r"主要符号(?:对照)?表|符号(?:与|及)?缩略语说明|符号说明|缩略语(?:对照)?表|符号表|术语表|"
        r"symbols?|abbreviations?|nomenclature|glossary|"
        r"参考文献|参考资料|引用文献|致\s*谢|谢\s*辞|后\s*记|附\s*录|"
        r"references|bibliography|works\s+cited|acknowledg(?:e)?ments?|appendix)",
        re.I,
    )

    paragraphs = [
        p for p in document.paragraphs if p.text.strip() or p._p.find(qn("w:drawing")) is not None
    ]
    if not paragraphs:
        return None, None

    placeholder_toc_p = None
    for p in paragraphs:
        if paragraph_style_name(document, p._p) in TOC_MANAGED_STYLES:
            continue
        if re_toc.match(p.text.strip()):
            placeholder_toc_p = p
            break

    if placeholder_toc_p is not None:
        return placeholder_toc_p, placeholder_toc_p

    def _upcoming_abstract_lang(start_idx: int) -> str | None:
        """Look ahead across 1-2 short abstract_title lines for a following Abstract/摘要."""
        limit = min(len(paragraphs), start_idx + 3)
        for j in range(start_idx, limit):
            cand = paragraphs[j]
            if paragraph_style_name(document, cand._p) in TOC_MANAGED_STYLES:
                continue
            ct = cand.text.strip()
            if not ct:
                continue
            if re_abstract_en.match(ct):
                return "en"
            if re_abstract_cn.match(ct):
                return "zh"
            c_style = paragraph_style_name(document, cand._p).lower()
            c_h1 = (
                c_style.startswith("heading 1")
                or c_style.startswith("标题 1")
                or c_style == "heading"
            )
            if (
                len(ct) > 200
                or c_h1
                or re_toc.match(ct)
                or re_next_module.match(ct)
                or re.search(r"[。；;！？!?]", ct)
            ):
                return None
        return None

    state = "SEEKING_ABSTRACT"
    anchor_p = None

    for idx, p in enumerate(paragraphs):
        if paragraph_style_name(document, p._p) in TOC_MANAGED_STYLES:
            continue
        t = p.text.strip()
        if not t:
            continue
        style_name = paragraph_style_name(document, p._p).lower()
        is_h1 = (
            style_name.startswith("heading 1")
            or style_name.startswith("标题 1")
            or style_name == "heading"
        )

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
            elif _upcoming_abstract_lang(idx) == "en":
                continue
            else:
                anchor_p = p
                break
        elif state == "IN_ENGLISH_ABSTRACT":
            if re_keywords_en.match(t) or re_keywords_cn.match(t):
                state = "IN_ENGLISH_KEYWORDS"
            elif re_abstract_cn.match(t):
                state = "IN_CHINESE_ABSTRACT"
            elif re_toc.match(t) or re_next_module.match(t) or is_h1:
                anchor_p = p
                break
        elif state == "IN_ENGLISH_KEYWORDS":
            if re_abstract_cn.match(t):
                state = "IN_CHINESE_ABSTRACT"
            elif _upcoming_abstract_lang(idx) == "zh":
                continue
            else:
                anchor_p = p
                break

    if anchor_p is None and state == "SEEKING_ABSTRACT":
        anchor_p = paragraphs[0] if paragraphs else None

    # If anchor_p was found in a subsequent section (e.g. Chapter 1 in body),
    # check if there is an intervening section break paragraph before anchor_p.
    # Inserting before the section break keeps TOC inside the front-matter section!
    if anchor_p is not None:
        prev_xml = anchor_p._p.getprevious()
        while prev_xml is not None and prev_xml.tag == qn("w:p"):
            sect_pr = prev_xml.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
            if sect_pr is not None:
                from docx.text.paragraph import Paragraph as DocxParagraph

                anchor_p = DocxParagraph(prev_xml, document)
                break
            if prev_xml.find(".//" + qn("w:t")) is not None:
                break
            prev_xml = prev_xml.getprevious()

    return anchor_p, placeholder_toc_p


def apply_table_of_contents(document: Document, tokens: dict[str, Any]) -> list[str]:
    if not tokens:
        return []
    removed = _remove_managed_toc(document)
    if not tokens.get("enabled"):
        return [f"Removed managed Word TOC block ({removed} paragraph(s))"]

    title = str(tokens.get("title") or "目录")
    max_level = int(tokens.get("max_heading_level", 3))
    from .contracts import FormatError
    from .properties import outline_level
    from .toc import entry_styles, headings, imported_fields

    for node in imported_fields(document):
        node.getparent().remove(node)
        removed += 1
    entries = headings(document, max_level)
    if not entries:
        raise FormatError(
            "Table of contents (TOC) has no confirmed headings", "needs_clarification"
        )
    # Preserve the template's font, numbering and paragraph style. A real
    # outline level makes conservative text-inferred headings searchable by
    # Word's TOC \\u switch without imposing built-in Heading formatting.
    from .properties import element

    for paragraph, level in entries:
        if outline_level(document, paragraph._p) is None:
            pr = element(paragraph._p, "pPr")
            element(pr, "outlineLvl").set(qn("w:val"), str(level - 1))
    cached_entries = ["  " * (level - 1) + p.text.strip() for p, level in entries]
    heading_style = get_or_create_paragraph_style(document, STYLE_WFM_TOC_HEADING)
    field_style = get_or_create_paragraph_style(document, STYLE_WFM_TOC_FIELD)
    break_style = get_or_create_paragraph_style(document, STYLE_WFM_TOC_BREAK)
    for style in (heading_style, field_style, break_style):
        if style.base_style is None:
            style.base_style = get_or_create_paragraph_style(document, "Normal")
    for style, key in (
        (heading_style, "title_format"),
        (field_style, "field_format"),
        (break_style, "break_format"),
    ):
        appearance = tokens.get(key, tokens.get("title_style", {}) if key == "title_format" else {})
        set_style_font(style, appearance)
        set_paragraph_format(style, appearance, document)
    if appearance := tokens.get("entry_format"):
        for style in entry_styles(document, max_level, create=True):
            set_style_font(style, appearance)
            set_paragraph_format(style, appearance, document)
    # The TOC title must never index itself, regardless of its visual appearance.
    from .properties import element

    element(element(heading_style.element, "pPr"), "outlineLvl").set(qn("w:val"), "9")

    anchor_p, placeholder_toc_p = _find_toc_insertion_anchor(document)
    anchor_xml = anchor_p._p if anchor_p is not None else None
    reusing_boundary = (
        anchor_p is not None and paragraph_style_name(document, anchor_xml) == STYLE_WFM_TOC_BREAK
        and anchor_xml.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None
    )

    if placeholder_toc_p is not None:
        title_paragraph = placeholder_toc_p
        saved_rpr = None
        for r in title_paragraph.runs:
            rpr = r._r.find(qn("w:rPr"))
            if rpr is not None:
                saved_rpr = copy.deepcopy(rpr)
                break
        clear_paragraph_content(title_paragraph)
        title_paragraph.style = heading_style
        title_run = title_paragraph.add_run(title)
        if saved_rpr is not None:
            title_run._r.insert(0, saved_rpr)
        title_appearance = tokens.get("title_format", tokens.get("title_style", {}))
        if title_appearance:
            set_run_font(title_run, title_appearance)
        insert_target = title_paragraph._p.getnext()
        field_paragraph = _insert_body_paragraph_before(document, insert_target, field_style)
        append_field(
            field_paragraph,
            f'TOC \\o "1-{max_level}" \\h \\z \\u',
            cached_text="\n".join(cached_entries) or "目录将在 Word 中更新",
        )
        if tokens.get("page_break_after", True) and not reusing_boundary:
            break_paragraph = _insert_body_paragraph_before(
                document, field_paragraph._p.getnext(), break_style
            )
            break_paragraph.add_run().add_break(WD_BREAK.PAGE)
    else:
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
            prev_p = (
                anchor_xml.getprevious()
                if anchor_xml is not None
                else (document.paragraphs[-1]._p if document.paragraphs else None)
            )
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
        append_field(
            field_paragraph,
            f'TOC \\o "1-{max_level}" \\h \\z \\u',
            cached_text="\n".join(cached_entries) or "目录将在 Word 中更新",
        )
        if tokens.get("page_break_after", True) and not reusing_boundary:
            break_paragraph = _insert_body_paragraph_before(document, anchor_xml, break_style)
            break_paragraph.add_run().add_break(WD_BREAK.PAGE)

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
        insert_child(pPr, num_pr)
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
        insert_child(pPr, num_pr)
    else:
        num_pr.clear()
    num_id_node = OxmlElement("w:numId")
    num_id_node.set(qn("w:val"), "0")
    num_pr.append(num_id_node)


def apply_reference_style(
    document: Document, body: dict[str, Any], references: dict[str, Any]
) -> list[str]:
    if not references:
        return []
    reference_tokens = deep_merge(body, references)
    hanging_indent = reference_tokens.pop("hanging_indent_mm", None)
    if hanging_indent is not None:
        reference_tokens["left_indent_mm"] = float(reference_tokens.get("left_indent_mm", 0.0))
        reference_tokens["left_indent_chars"] = float(
            reference_tokens.get("left_indent_chars", 0.0)
        )
        reference_tokens["first_line_indent_mm"] = -float(hanging_indent)
        reference_tokens["hanging_indent_mm"] = float(hanging_indent)
    bibliography = resolve_paragraph_style(document, references, STYLE_BIBLIOGRAPHY)
    if bibliography.base_style is None:
        bibliography.base_style = get_or_create_paragraph_style(document, "Normal")
    set_style_font(bibliography, reference_tokens)
    set_paragraph_format(bibliography, reference_tokens, document)

    numbering_mode = references.get("numbering_mode")
    num_id = None
    if numbering_mode == "word-numbering":
        style_key = "decimal-bracket"
        num_fmt, level_text = LIST_NUMBERING_STYLES[style_key]
        numbering = _get_numbering_element(document, create=True)
        abstract_id = _numbering_max_id(numbering, "abstractNum", "abstractNumId") + 1
        # numId=0 means "suppress numbering", even in an otherwise empty part.
        num_id = max(0, _numbering_max_id(numbering, "num", "numId")) + 1

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
        suff_node = OxmlElement("w:suff")
        suff_node.set(qn("w:val"), "space")
        lvl.append(suff_node)
        text_node = OxmlElement("w:lvlText")
        text_node.set(qn("w:val"), level_text)
        lvl.append(text_node)
        justification = OxmlElement("w:lvlJc")
        justification.set(qn("w:val"), "left")
        lvl.append(justification)

        # Marker typography and indentation inherit unless explicitly supplied.
        # The chosen template already applies reference paragraph properties.
        from .properties import element, write_font
        from .spec import FONT

        for property_name in FONT:
            if property_name in reference_tokens:
                write_font(element(lvl, "rPr"), property_name, reference_tokens[property_name])

        abstract.append(lvl)
        insert_child(numbering, abstract)

        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_ref = OxmlElement("w:abstractNumId")
        abstract_ref.set(qn("w:val"), str(abstract_id))
        num.append(abstract_ref)
        insert_child(numbering, num)

    from .inspect import DocumentIndex

    reference_nodes = DocumentIndex(document).reference_nodes()
    styled = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if paragraph._p in reference_nodes:
            paragraph.style = bibliography
            if numbering_mode == "word-numbering" and num_id is not None:
                clean_text = re.sub(r"^\s*\[\s*\d+\s*\]\s*|^\s*\d+[\.、]\s*", "", text).strip()
                from .properties import strip_text_prefix

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


def _num_format_for_num_id(numbering, num_id: str | None, ilvl: int = 0) -> str | None:
    if not num_id:
        return None
    num = next(
        (item for item in numbering.findall(qn("w:num")) if item.get(qn("w:numId")) == num_id), None
    )
    if num is None:
        return None
    abstract_id = num.find(qn("w:abstractNumId"))
    if abstract_id is None:
        return None
    abstract = next(
        (
            item
            for item in numbering.findall(qn("w:abstractNum"))
            if item.get(qn("w:abstractNumId")) == abstract_id.get(qn("w:val"))
        ),
        None,
    )
    level = (
        next(
            (node for node in abstract.findall(qn("w:lvl")) if node.get(qn("w:ilvl")) == str(ilvl)),
            None,
        )
        if abstract is not None
        else None
    )
    override = next(
        (node for node in num.findall(qn("w:lvlOverride")) if node.get(qn("w:ilvl")) == str(ilvl)),
        None,
    )
    if override is not None and override.find(qn("w:lvl")) is not None:
        level = override.find(qn("w:lvl"))
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
        from docx.parts.numbering import NumberingPart

        numbering_elm = parse_xml(
            b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )
        part = NumberingPart(
            PackURI("/word/numbering.xml"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            numbering_elm,
            document.part.package,
        )
        document.part.relate_to(part, RT.NUMBERING)
        return numbering_elm


def _has_numbered_list_marker(document: Document, paragraph, *, index=None) -> bool:
    from .inspect import DocumentIndex

    index = index or DocumentIndex(document)
    oid = index.ids_by_element.get(paragraph._p)
    record = index.records.get(oid, {})
    if not record.get("top_level") or record.get("semantic_role") not in {"body_text", "list"}:
        return False
    style_name = paragraph_style_name(document, paragraph._p).lower()
    if (
        style_name.startswith("heading")
        or style_name.startswith("标题")
        or style_name
        in {
            "title",
            "subtitle",
            "toc",
            STYLE_WFM_TOC_HEADING.lower(),
            *BIBLIOGRAPHY_STYLE_ALIASES,
        }
    ):
        return False
    from .numbering import effective_numbering
    from .properties import read_paragraph

    if read_paragraph(document, paragraph._p, "numbering_suppressed"):
        return False
    found = effective_numbering(document, paragraph._p)
    if found is not None and found[2] is not None:
        fmt = _num_format_for_num_id(_get_numbering_element(document), found[0], found[1])
        return fmt not in {None, "bullet", "picture"}
    return any(
        marker in style_name for marker in ("list number", "numbered", "编号", "list paragraph")
    )


def _append_numbering_definition(document: Document, style: str, start: int) -> int:
    num_fmt, level_text = LIST_NUMBERING_STYLES[style]
    numbering = _get_numbering_element(document, create=True)
    abstract_id = _numbering_max_id(numbering, "abstractNum", "abstractNumId") + 1
    num_id = max(0, _numbering_max_id(numbering, "num", "numId")) + 1
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
    suff_node = OxmlElement("w:suff")
    suff_node.set(qn("w:val"), "space")
    lvl.append(suff_node)
    text_node = OxmlElement("w:lvlText")
    text_node.set(qn("w:val"), level_text)
    lvl.append(text_node)
    justification = OxmlElement("w:lvlJc")
    justification.set(qn("w:val"), "left")
    lvl.append(justification)
    paragraph_properties = OxmlElement("w:pPr")
    indentation = OxmlElement("w:ind")
    indentation.set(qn("w:left"), str(DEFAULT_LIST_LEFT_INDENT_TWIPS))
    indentation.set(qn("w:hanging"), str(DEFAULT_LIST_HANGING_INDENT_TWIPS))
    paragraph_properties.append(indentation)
    lvl.append(paragraph_properties)
    abstract.append(lvl)
    insert_child(numbering, abstract)
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
    insert_child(numbering, num)
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
    insert_child(ppr, num_pr)


def normalize_universal_chinese_base(
    document: Document,
    config: UniversalChineseBaseConfig = DEFAULT_CHINESE_BASE_CONFIG,
) -> Document:
    """Normalize a python-docx Document into a clean universal Chinese academic baseline.

    Cleans Office/python-docx default theme font language (ja-JP -> zh-CN),
    replaces Japanese/Western theme fonts (MS Mincho / MS Gothic / Cambria / Calibri)
    in theme1.xml with ``config.east_asia_body_font`` / ``config.latin_font``, neutralizes
    theme accent1/accent2 blue colors (#365F91 / #4F81BD) in styles.xml, and standardizes
    docDefaults, Normal, Heading 1..9, Title, Bibliography, toc 1..9, header, and footer styles.
    """
    from .properties import element, write_font, write_paragraph

    line_spacing_spec = {
        "kind": "multiple",
        "value": config.line_spacing_multiple,
        "unit": "multiple",
    }
    legacy_theme_fonts = {"ＭＳ 明朝", "ＭＳ ゴシック", "Cambria", "Calibri"}
    toc_right_tab_twips = section_content_width_twips(
        document.sections[0] if document.sections else None
    )

    # 1. settings.xml: w:themeFontLang -> zh-CN / en-US
    try:
        settings_el = document.settings.element
        lang_el = settings_el.find(qn("w:themeFontLang"))
        if lang_el is None:
            lang_el = OxmlElement("w:themeFontLang")
            settings_el.append(lang_el)
        lang_el.set(qn("w:val"), "en-US")
        lang_el.set(qn("w:eastAsia"), "zh-CN")
        lang_el.set(qn("w:bidi"), "ar-SA")
    except AttributeError:
        pass

    # 2. theme1.xml: replace ja-JP MS Mincho/Gothic and Cambria/Calibri with configured fonts
    for part in document.part.package.parts:
        if str(part.partname) == "/word/theme/theme1.xml":
            blob_str = part.blob.decode("utf-8")
            blob_str = re.sub(
                r'typeface="ＭＳ 明朝"', f'typeface="{config.east_asia_body_font}"', blob_str
            )
            blob_str = re.sub(
                r'typeface="ＭＳ ゴシック"', f'typeface="{config.east_asia_body_font}"', blob_str
            )
            blob_str = re.sub(r'typeface="Cambria"', f'typeface="{config.latin_font}"', blob_str)
            blob_str = re.sub(r'typeface="Calibri"', f'typeface="{config.latin_font}"', blob_str)
            blob_str = re.sub(
                r'(<a:font\s+script="Hans"\s+typeface=")[^"]*(")',
                rf"\g<1>{config.east_asia_body_font}\2",
                blob_str,
            )
            blob_str = re.sub(
                r'(<a:font\s+script="Jpan"\s+typeface=")[^"]*(")',
                rf"\g<1>{config.east_asia_body_font}\2",
                blob_str,
            )
            part._blob = blob_str.encode("utf-8")
            if hasattr(document.part, "_wfm_theme_part"):
                delattr(document.part, "_wfm_theme_part")

    # 3. styles.xml: docDefaults -> configured body font, size, color, and line spacing
    styles_el = document.styles.element
    doc_defaults = element(styles_el, "docDefaults")
    rpr_default = element(doc_defaults, "rPrDefault")
    rpr_def_inner = element(rpr_default, "rPr")
    write_font(rpr_def_inner, "font_east_asia", config.east_asia_body_font)
    write_font(rpr_def_inner, "font_latin", config.latin_font)
    write_font(rpr_def_inner, "font_size_pt", config.body_font_size_pt)
    write_font(rpr_def_inner, "color", config.text_color)
    lang_def = element(rpr_def_inner, "lang")
    lang_def.set(qn("w:val"), "en-US")
    lang_def.set(qn("w:eastAsia"), "zh-CN")

    ppr_default = element(doc_defaults, "pPrDefault")
    ppr_def_inner = element(ppr_default, "pPr")
    write_paragraph(ppr_def_inner, "space_before_pt", 0.0)
    write_paragraph(ppr_def_inner, "space_after_pt", 0.0)
    write_paragraph(ppr_def_inner, "line_spacing", line_spacing_spec)

    # 4. Neutralize themefonts, blue colors, and tblStylePr shading across all styles in styles.xml
    for st in styles_el.findall(qn("w:style")):
        for rpr in st.iter(qn("w:rPr")):
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                    rf.attrib.pop(qn("w:" + attr), None)
                rf.set(qn("w:ascii"), config.latin_font)
                rf.set(qn("w:hAnsi"), config.latin_font)
                if not rf.get(qn("w:eastAsia")) or rf.get(qn("w:eastAsia")) in legacy_theme_fonts:
                    rf.set(qn("w:eastAsia"), config.east_asia_body_font)
            col = rpr.find(qn("w:color"))
            if col is not None:
                for attr in ("themeColor", "themeTint", "themeShade"):
                    col.attrib.pop(qn("w:" + attr), None)
                col.set(qn("w:val"), config.text_color)
        for tbl_style_pr in list(st.findall(qn("w:tblStylePr"))):
            for shd in list(tbl_style_pr.iter(qn("w:shd"))):
                shd.getparent().remove(shd)

    # 5. Configure Normal, Heading 1..9, Title, Bibliography, toc 1..9, header, footer
    normal_st = get_or_create_paragraph_style(document, "Normal")
    normal_rpr = element(normal_st.element, "rPr")
    write_font(normal_rpr, "font_east_asia", config.east_asia_body_font)
    write_font(normal_rpr, "font_latin", config.latin_font)
    write_font(normal_rpr, "font_size_pt", config.body_font_size_pt)
    write_font(normal_rpr, "color", config.text_color)
    normal_ppr = element(normal_st.element, "pPr")
    write_paragraph(normal_ppr, "alignment", "justify")
    write_paragraph(normal_ppr, "space_before_pt", 0.0)
    write_paragraph(normal_ppr, "space_after_pt", 0.0)
    write_paragraph(normal_ppr, "line_spacing", line_spacing_spec)

    heading_specs: dict[str, tuple[str, float, bool, str, int]] = {
        "Heading 1": (config.east_asia_heading_font, config.h1_font_size_pt, True, "center", 0),
        "Heading 2": (config.east_asia_heading_font, config.h2_font_size_pt, True, "left", 1),
        "Heading 3": (config.east_asia_heading_font, config.h3_plus_font_size_pt, True, "left", 2),
    }
    heading_specs.update(
        {
            f"Heading {level}": (
                config.east_asia_heading_font,
                config.h3_plus_font_size_pt,
                True,
                "left",
                level - 1,
            )
            for level in range(4, 10)
        }
    )
    for h_name, (ea_font, sz_pt, is_bold, align, out_lvl) in heading_specs.items():
        st = get_or_create_paragraph_style(document, h_name)
        rpr = element(st.element, "rPr")
        write_font(rpr, "font_east_asia", ea_font)
        write_font(rpr, "font_latin", config.latin_font)
        write_font(rpr, "font_size_pt", sz_pt)
        write_font(rpr, "bold", is_bold)
        write_font(rpr, "italic", False)
        write_font(rpr, "color", config.text_color)
        ppr = element(st.element, "pPr")
        write_paragraph(ppr, "alignment", align)
        write_paragraph(ppr, "space_before_pt", 0.0)
        write_paragraph(ppr, "space_after_pt", 0.0)
        write_paragraph(ppr, "line_spacing", line_spacing_spec)
        write_paragraph(ppr, "first_line_indent", {"value": 0, "unit": "chars"})
        write_paragraph(ppr, "left_indent", {"value": 0, "unit": "chars"})
        write_paragraph(ppr, "outline_level", out_lvl)

    for style_name, east_asia, size in (
        ("Title", config.east_asia_heading_font, config.title_font_size_pt),
        (STYLE_BIBLIOGRAPHY, config.east_asia_body_font, config.reference_font_size_pt),
    ):
        style = get_or_create_paragraph_style(document, style_name)
        rpr = element(style.element, "rPr")
        for name, value in (
            ("font_east_asia", east_asia),
            ("font_latin", config.latin_font),
            ("font_size_pt", size),
            ("color", config.text_color),
        ):
            write_font(rpr, name, value)
        language = element(rpr, "lang")
        language.set(qn("w:val"), "en-US")
        language.set(qn("w:eastAsia"), "zh-CN")
        if style_name == "Title":
            ppr = element(style.element, "pPr")
            write_paragraph(ppr, "alignment", "center")
            write_paragraph(ppr, "first_line_indent", {"value": 0, "unit": "chars"})
            border = ppr.find(qn("w:pBdr"))
            if border is not None:
                ppr.remove(border)

    for lvl in range(1, 10):
        st = None
        for cand_name in (f"toc {lvl}", f"TOC {lvl}", f"目录 {lvl}"):
            try:
                st = document.styles[cand_name]
                break
            except KeyError:
                pass
        if st is None:
            st = document.styles.add_style(f"toc {lvl}", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
            st.base_style = normal_st
        rpr = element(st.element, "rPr")
        write_font(rpr, "font_east_asia", config.east_asia_body_font)
        write_font(rpr, "font_latin", config.latin_font)
        write_font(rpr, "font_size_pt", config.body_font_size_pt)
        write_font(rpr, "bold", False)
        write_font(rpr, "italic", False)
        write_font(rpr, "color", config.text_color)
        ppr = element(st.element, "pPr")
        write_paragraph(ppr, "alignment", "left")
        write_paragraph(ppr, "space_before_pt", 0.0)
        write_paragraph(ppr, "space_after_pt", 0.0)
        write_paragraph(ppr, "line_spacing", line_spacing_spec)
        write_paragraph(ppr, "first_line_indent", {"value": 0, "unit": "chars"})
        write_paragraph(
            ppr,
            "left_indent",
            {"value": float((lvl - 1) * config.toc_indent_step_chars), "unit": "chars"},
        )
        write_paragraph(
            ppr,
            "tabs",
            [
                {
                    "alignment": "right",
                    "position": toc_right_tab_twips,
                    "unit": "twips",
                    "leader": "dot",
                }
            ],
        )

    for hf_name in ("Header", "Footer", "header", "footer"):
        try:
            st = document.styles[hf_name]
        except KeyError:
            continue
        rpr = element(st.element, "rPr")
        write_font(rpr, "font_east_asia", config.east_asia_body_font)
        write_font(rpr, "font_latin", config.latin_font)
        write_font(rpr, "font_size_pt", config.header_footer_font_size_pt)
        write_font(rpr, "color", config.text_color)

    return document


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
    from .inspect import DocumentIndex

    semantic_index = DocumentIndex(document)
    targets = [paragraph for paragraph in paragraphs
               if _has_numbered_list_marker(document, paragraph, index=semantic_index)]
    if not targets:
        return [
            f"Registered managed list numbering style {style}; no numbered list paragraphs were found"
        ]
    changes: list[str] = []
    target_ids = {id(paragraph) for paragraph in targets}
    _get_numbering_element(document, create=True)
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
            from .properties import write_paragraph

            ppr = paragraph._p.get_or_add_pPr()
            if left_indent is not None:
                write_paragraph(ppr, "left_indent", {"value": float(left_indent), "unit": "mm"})
            if hanging_indent is not None:
                write_paragraph(
                    ppr, "hanging_indent", {"value": float(hanging_indent), "unit": "mm"}
                )
        changes.append(f"Applied managed {style} numbering to {end - index + 1} list paragraph(s)")
        index = end + 1
    return changes
