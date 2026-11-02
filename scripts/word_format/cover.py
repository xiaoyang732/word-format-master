"""Render or clone cover layouts from explicit data; appearance belongs to templates."""

from __future__ import annotations

import copy
import json
import re
from pathlib import Path

from document_structure import classify_module
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.shared import Mm

from .contracts import FormatError, text_value
from .doc_loader import SUPPORTED_WORD_SUFFIXES, open_word_document
from .namespaces import qn
from .properties import (
    element,
    outline_level,
    paragraph_style,
    read_font,
    read_paragraph,
    rpr_sources,
    write_font,
    write_paragraph,
)
from .structures import set_paragraph_format, set_run_font

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_OFFICE_NS = "urn:schemas-microsoft-com:office:office"


def apply_east_asia_font(run, font_name: str, ascii_font: str):
    write_font(element(run._r, "rPr"), "font_east_asia", font_name)
    write_font(element(run._r, "rPr"), "font_latin", ascii_font)


def _paragraph(container, config, data):
    value = data.get(config["field"])
    if not value:
        return
    paragraph = container.add_paragraph()
    set_paragraph_format(paragraph, config["tokens"])
    set_run_font(paragraph.add_run(str(value)), config["tokens"])


def render_cover(doc, layout, data, *, fields=None, use_section_break=True):
    if layout.get("kind") != "word-format-cover-layout" or layout.get("version") != "1.0":
        raise FormatError("Unsupported cover layout")
    text_value(data.get("title_cn"), "cover title")
    if "different_first_page" in layout:
        doc.sections[0].different_first_page_header_footer = layout["different_first_page"]
    for block in layout["blocks"]:
        if not block.get("when") or data.get(block["when"]):
            _paragraph(doc, block, data)
    if fields is None:
        fields = [(label, data[key]) for label, key in layout["labels"] if data.get(key)]
    if fields:
        config = layout["field_table"]
        grid = doc.add_table(rows=len(fields), cols=2)
        grid.alignment = {
            "left": WD_TABLE_ALIGNMENT.LEFT,
            "center": WD_TABLE_ALIGNMENT.CENTER,
            "right": WD_TABLE_ALIGNMENT.RIGHT,
        }[config["alignment"]]
        grid.autofit = False
        element(grid._tbl.tblPr, "tblLayout").set(qn("w:type"), "fixed")
        borders = element(grid._tbl.tblPr, "tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element(borders, side).set(qn("w:val"), "nil")
        for col, width in zip(grid.columns, config["widths_mm"], strict=True):
            col.width = Mm(width)
        for row, (label, value) in zip(grid.rows, fields, strict=True):
            for i, (content, key) in enumerate(((label, "label"), (value, "value"))):
                cell = row.cells[i]
                cell.width = Mm(config["widths_mm"][i])
                paragraph = cell.paragraphs[0]
                set_paragraph_format(paragraph, config[key])
                set_run_font(paragraph.add_run(str(content)), config[key])
            if config.get("value_border"):
                border = element(element(row.cells[1]._tc.get_or_add_tcPr(), "tcBorders"), "bottom")
                for key, value in config["value_border"].items():
                    border.set(qn("w:" + key), str(value))
    if layout.get("date"):
        _paragraph(doc, layout["date"], data)
    if use_section_break:
        section = doc.add_section(WD_SECTION_START.NEW_PAGE)
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
    else:
        doc.add_page_break()


def add_standard_thesis_cover(
    doc, *, title_cn, fields=None, use_section_break=True, layout=None, **data
):
    """Compatibility entry point selecting the independent standard layout explicitly."""
    if layout is None:
        path = Path(__file__).resolve().parents[2] / "templates/builtin/thesis-standard/cover.json"
        layout = json.loads(path.read_text(encoding="utf-8"))
    render_cover(
        doc,
        layout,
        {"title_cn": title_cn, **data},
        fields=fields,
        use_section_break=use_section_break,
    )


_SP = r"[\s\u3000]*"
_FIELD_LABEL_CORES = {
    "title": (
        rf"(?:(?:毕业|学位|课程)?(?:论文|设计)?(?:[（(](?:设计|论文)[）)])?{_SP}题{_SP}目|"
        rf"(?:课题|论文|设计|项目){_SP}名{_SP}称)"
    ),
    "title_en": rf"(?:英文|外文){_SP}(?:论文|设计)?{_SP}题{_SP}目|english\s+title",
    "subtitle": rf"副{_SP}题{_SP}目|副{_SP}标{_SP}题",
    "student": rf"(?:学生|作者|研究生|申请人)?{_SP}姓{_SP}名|作{_SP}者",
    "student_id": rf"(?:学生|考生)?{_SP}(?:学{_SP}号|考{_SP}号)",
    "college": (
        rf"(?:所属|所在|二级)?{_SP}(?:学{_SP}院(?:{_SP}[（(]{_SP}系{_SP}[）)])?|"
        rf"院{_SP}[（(]{_SP}系{_SP}[）)]|院{_SP}系|系{_SP}别|学{_SP}部|培养单位)"
    ),
    "major": rf"(?:所学|学科)?{_SP}(?:专{_SP}业(?:名称|学位类别|领域)?|学科专业)",
    "direction": rf"(?:研究|专业){_SP}方{_SP}向",
    "class": rf"(?:专业|所在|行政)?{_SP}班{_SP}级",
    "grade": rf"年{_SP}级|届{_SP}别",
    "advisor": rf"(?:校内|第一)?{_SP}(?:指导{_SP}(?:教师|老师)|导{_SP}师(?:姓名)?)",
    "co_advisor": rf"(?:企业|校外|第二|合作|副){_SP}(?:指导{_SP}(?:教师|老师)|导{_SP}师)",
    "advisor_title": rf"(?:指导教师|导师)?{_SP}职{_SP}称",
    "school_code": rf"(?:学校|单位){_SP}代{_SP}码",
    "classification": rf"(?:中图)?{_SP}分{_SP}类{_SP}号",
    "udc": r"udc",
    "security": rf"密{_SP}级",
    "date": rf"(?:提交|完成|答辩|写作)?{_SP}日{_SP}期|时{_SP}间",
    "signature": rf"(?:本人)?{_SP}(?:签{_SP}名|签{_SP}字)|声{_SP}明{_SP}人",
    "signature_date": rf"(?:声{_SP}明|签{_SP}名|签{_SP}字)?{_SP}日{_SP}期",
}

_FIELD_ALIASES = {
    "signature_date": "signature_date",
    "declaration_date": "signature_date",
    "签名日期": "signature_date",
    "声明日期": "signature_date",
    "signature": "signature",
    "本人签名": "signature",
    "签名": "signature",
    "签字": "signature",
    "声明人": "signature",
    "title": "title",
    "subject": "title",
    "title_cn": "title",
    "题目": "title",
    "论文题目": "title",
    "设计题目": "title",
    "课题名称": "title",
    "title_en": "title_en",
    "英文题目": "title_en",
    "外文题目": "title_en",
    "subtitle": "subtitle",
    "副标题": "subtitle",
    "student": "student",
    "student_name": "student",
    "author": "student",
    "author_name": "student",
    "学生姓名": "student",
    "姓名": "student",
    "作者": "student",
    "student_id": "student_id",
    "学号": "student_id",
    "学生学号": "student_id",
    "college": "college",
    "school": "college",
    "department": "college",
    "二级学院": "college",
    "学院": "college",
    "院系": "college",
    "院（系）": "college",
    "院(系)": "college",
    "系别": "college",
    "学部": "college",
    "major": "major",
    "专业": "major",
    "所学专业": "major",
    "专业名称": "major",
    "学科专业": "major",
    "direction": "direction",
    "research_direction": "direction",
    "研究方向": "direction",
    "class": "class",
    "class_name": "class",
    "班级": "class",
    "专业班级": "class",
    "grade": "grade",
    "年级": "grade",
    "届别": "grade",
    "advisor": "advisor",
    "supervisor": "advisor",
    "指导教师": "advisor",
    "指导老师": "advisor",
    "导师": "advisor",
    "校内导师": "advisor",
    "co_advisor": "co_advisor",
    "enterprise_advisor": "co_advisor",
    "企业导师": "co_advisor",
    "校外导师": "co_advisor",
    "第二导师": "co_advisor",
    "合作导师": "co_advisor",
    "advisor_title": "advisor_title",
    "职称": "advisor_title",
    "导师职称": "advisor_title",
    "school_code": "school_code",
    "学校代码": "school_code",
    "单位代码": "school_code",
    "classification": "classification",
    "分类号": "classification",
    "udc": "udc",
    "UDC": "udc",
    "security": "security",
    "密级": "security",
    "date": "date",
    "日期": "date",
    "完成日期": "date",
    "提交日期": "date",
}

# Backwards-compatible mapping of canonical keys to inline regex strings.
_COVER_FIELD_PATTERNS = {
    alias: rf"^(?:{_FIELD_LABEL_CORES[canonical]}){_SP}[:：]"
    for alias, canonical in _FIELD_ALIASES.items()
}

_STANDALONE_DATE_RE = re.compile(
    r"^[\s\u3000]*(?:[\d〇○零一二三四五六七八九十][\s\u3000]*){2,4}年.*月(?:.*日)?[\s\u3000]*$"
    r"|^[\s\u3000]*年[\s\u3000]*月[\s\u3000]*日[\s\u3000]*$"
)
_STANDALONE_TITLE_PLACEHOLDER_RE = re.compile(
    r"^[（(【\[]?\s*(?:在此(?:处)?输入)?(?:毕业)?(?:论文|设计)?题\s*目\s*[）)】\]]?$"
)
_COVER_BANNER_RE = re.compile(
    r"(?:毕业(?:设计|论文)|学位论文|课程(?:设计|论文|报告)|实验报告|实习报告|开题报告|期末作业|封面)"
)
_EXPLICIT_CHAPTER_START_RE = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十百零〇0-9]+\s*[章节篇]|[一二三四五六七八九十]+、|\d+(?:[、.．]\s*|\s+)\S)"
)
_NON_COVER_MODULES = frozenset(
    {
        "declaration",
        "abstract",
        "keywords",
        "toc",
        "symbols",
        "references",
        "acknowledgements",
        "appendix",
    }
)
_OPAQUE_RUN_TAGS = frozenset(
    {
        qn("w:drawing"),
        qn("w:pict"),
        qn("w:object"),
        qn("w:fldChar"),
        qn("w:instrText"),
    }
)


def _custom_label_core(label: str) -> str:
    cleaned = re.sub(r"[\s\u3000]+", "", label.strip().rstrip(":：").strip())
    if not cleaned:
        raise FormatError("Cover field key must contain a non-empty label")
    parts = []
    for ch in cleaned:
        if ch in {"(", "（"}:
            parts.append(r"[（(]")
        elif ch in {")", "）"}:
            parts.append(r"[）)]")
        else:
            parts.append(re.escape(ch))
    return _SP.join(parts)


def _field_patterns(
    key: str,
) -> tuple[str | None, re.Pattern[str], re.Pattern[str], re.Pattern[str]]:
    validated = text_value(key, "cover field key")
    canonical = _FIELD_ALIASES.get(validated) or _FIELD_ALIASES.get(validated.lower())
    exact_core = _custom_label_core(validated)
    exact_cell_re = re.compile(rf"^[\s\u3000]*(?:{exact_core}){_SP}[:：]?[\s\u3000]*$", re.I)
    core = _FIELD_LABEL_CORES[canonical] if canonical else exact_core
    cell_re = re.compile(rf"^[\s\u3000]*(?:{core}){_SP}[:：]?[\s\u3000]*$", re.I)
    inline_re = re.compile(rf"(?:^|(?<=[\s\u3000]))(?:{core}){_SP}[:：]", re.I)
    return canonical, exact_cell_re, cell_re, inline_re


def _build_any_inline_label_re(extra_keys: list[str]) -> re.Pattern[str]:
    cores = list(_FIELD_LABEL_CORES.values())
    for key in extra_keys:
        canonical = _FIELD_ALIASES.get(key) or _FIELD_ALIASES.get(key.lower())
        if not canonical:
            cores.append(_custom_label_core(key))
    joined = "|".join(f"(?:{c})" for c in cores)
    return re.compile(rf"[\s\u3000]+(?:{joined}){_SP}[:：]", re.I)


def _half_width_len(value: str) -> int:
    return sum(2 if ord(ch) > 127 else 1 for ch in value)


def _has_active_underline(rpr) -> bool:
    if rpr is None:
        return False
    u = rpr.find(qn("w:u"))
    return u is not None and u.get(qn("w:val"), "single") not in {"none", "0", "false"}


def _format_slot_value(value: str, old_slot: str, val_rpr) -> tuple[str, object | None]:
    rpr_copy = copy.deepcopy(val_rpr) if val_rpr is not None else None
    is_underscore_slot = bool(
        old_slot and re.fullmatch(r"[\s\u3000]*[_＿]{2,}[\s\u3000]*", old_slot)
    )
    if is_underscore_slot:
        if rpr_copy is None:
            rpr_copy = OxmlElement("w:rPr")
        u = element(rpr_copy, "u")
        if u.get(qn("w:val")) in {None, "none", "0", "false"}:
            u.set(qn("w:val"), "single")
    has_underline = _has_active_underline(rpr_copy)
    has_padding = bool(
        old_slot
        and (
            is_underscore_slot
            or old_slot != old_slot.strip(" \t\u3000")
            or (not old_slot.strip(" \t\u3000") and len(old_slot) >= 2)
        )
    )
    if value and has_underline and has_padding:
        slot_width = _half_width_len(old_slot)
        val_width = _half_width_len(value)
        if val_width < slot_width:
            pad_total = slot_width - val_width
            left_pad = pad_total // 2
            right_pad = pad_total - left_pad
            return (" " * left_pad) + value + (" " * right_pad), rpr_copy
    return value, rpr_copy


def _make_text_run(text: str, rpr) -> object:
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")
    t.text = text
    run.append(t)
    return run


def _is_opaque_run(run) -> bool:
    if any(child.tag in _OPAQUE_RUN_TAGS for child in run.iter()):
        return True
    return any(
        br.get(qn("w:type")) == "page" for br in run.findall(qn("w:br"))
    )


def _replace_paragraph_text_pattern(p, pattern: str | re.Pattern[str], replacement: str) -> bool:
    runs = [r for r in p.findall(qn("w:r")) if not _is_opaque_run(r)]
    if not runs:
        return False
    spans = []
    offset = 0
    for r in runs:
        t_nodes = r.findall(qn("w:t"))
        r_text = "".join(t.text or "" for t in t_nodes)
        spans.append((r, offset, offset + len(r_text), r_text))
        offset += len(r_text)
    full_text = "".join(s[3] for s in spans)
    m = re.search(pattern, full_text) if isinstance(pattern, str) else pattern.search(full_text)
    if not m:
        return False
    start, end = m.span()
    target_runs = [item for item in spans if item[2] > start and item[1] < end]
    if not target_runs:
        return False
    first_r, f_start, f_end, f_text = target_runs[0]
    last_r, l_start, l_end, l_text = target_runs[-1]
    prefix = f_text[: start - f_start]
    suffix = l_text[end - l_start :]
    for r, _, _, _ in target_runs[1:]:
        p.remove(r)
    for t in first_r.findall(qn("w:t")):
        first_r.remove(t)
    new_t = OxmlElement("w:t")
    new_text = prefix + replacement + suffix
    if new_text != new_text.strip() or "  " in new_text:
        new_t.set(XML_SPACE, "preserve")
    new_t.text = new_text
    first_r.append(new_t)
    return True


def _set_labeled_paragraph_value(
    p,
    value: str,
    *,
    inline_re: re.Pattern[str] | None = None,
    any_label_re: re.Pattern[str] | None = None,
) -> bool:
    text_runs = [r for r in p.findall(qn("w:r")) if not _is_opaque_run(r)]
    if not text_runs:
        return False
    spans: list[tuple[object, int, int, str]] = []
    offset = 0
    for r in text_runs:
        r_text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        spans.append((r, offset, offset + len(r_text), r_text))
        offset += len(r_text)
    full_text = "".join(item[3] for item in spans)
    if inline_re is not None:
        match = inline_re.search(full_text)
        if not match:
            return False
        val_start = match.end()
    else:
        colon_pos = next((i for i, ch in enumerate(full_text) if ch in {":", "："}), None)
        if colon_pos is None:
            return False
        val_start = colon_pos + 1

    val_end = len(full_text)
    if any_label_re is not None and val_start < len(full_text):
        next_match = any_label_re.search(full_text[val_start:])
        if next_match:
            val_end = val_start + next_match.start()
    if val_start < len(full_text):
        inline_date = re.search(
            r"[\s\u3000]{2,}(?=(?:(?:20\d\d|\d{2,4})\s*年|年\s*月\s*日))",
            full_text[val_start:],
        )
        if inline_date:
            val_end = min(val_end, val_start + inline_date.start())

    old_slot = full_text[val_start:val_end]
    val_rpr = None
    for r, s, e, _ in spans:
        if e > val_start and s < val_end and s >= val_start:
            if r.find(qn("w:rPr")) is not None:
                val_rpr = r.find(qn("w:rPr"))
                break
    if val_rpr is None:
        for r, s, e, _ in spans:
            if e > val_start and s < val_end and r.find(qn("w:rPr")) is not None:
                val_rpr = r.find(qn("w:rPr"))
                break
    if val_rpr is None:
        for r, s, e, _ in spans:
            if s < val_start <= e and r.find(qn("w:rPr")) is not None:
                val_rpr = r.find(qn("w:rPr"))
                break
    if val_rpr is None:
        ppr_rpr = p.find("./" + qn("w:pPr") + "/" + qn("w:rPr"))
        val_rpr = ppr_rpr if ppr_rpr is not None else text_runs[0].find(qn("w:rPr"))

    formatted_val, final_rpr = _format_slot_value(str(value), old_slot, val_rpr)

    # Locate anchor run holding the character immediately before val_start (the colon).
    anchor_idx = next(
        (idx for idx, (_, s, e, _) in enumerate(spans) if s < val_start <= e),
        0,
    )
    anchor_run, a_start, a_end, a_text = spans[anchor_idx]
    anchor_rpr = anchor_run.find(qn("w:rPr"))
    prefix_text = a_text[: val_start - a_start]
    same_run_suffix = a_text[val_end - a_start :] if val_end < a_end else ""

    # Determine trailing suffix if val_end lands inside a later run.
    later_suffix = ""
    later_rpr = anchor_rpr
    remove_runs = []
    for r, s, e, r_text in spans[anchor_idx + 1 :]:
        if s >= val_end:
            break
        remove_runs.append(r)
        if val_end < e:
            later_suffix = r_text[val_end - s :]
            later_rpr = r.find(qn("w:rPr"))
            break

    # Rewrite anchor_run to contain only prefix_text.
    t_nodes = anchor_run.findall(qn("w:t"))
    for t in t_nodes[1:]:
        anchor_run.remove(t)
    if t_nodes:
        t_nodes[0].text = prefix_text
        if prefix_text != prefix_text.strip() or "  " in prefix_text:
            t_nodes[0].set(XML_SPACE, "preserve")
    elif prefix_text:
        t0 = OxmlElement("w:t")
        t0.text = prefix_text
        anchor_run.append(t0)

    for r in remove_runs:
        p.remove(r)

    cursor = anchor_run
    if formatted_val:
        val_run = _make_text_run(formatted_val, final_rpr)
        cursor.addnext(val_run)
        cursor = val_run
    trailing_text = same_run_suffix or later_suffix
    if trailing_text:
        tail_rpr = anchor_rpr if same_run_suffix else later_rpr
        tail_run = _make_text_run(trailing_text, tail_rpr)
        cursor.addnext(tail_run)
    return True


def _paragraph_top_text(p) -> str:
    return "".join(
        t.text or ""
        for t in p.iter(qn("w:t"))
        if not any(
            a.tag in {qn("w:txbxContent"), qn("w:del")} for a in t.iterancestors() if a is not p
        )
    )


def _cell_text(tc) -> str:
    return "".join(
        t.text or ""
        for t in tc.iter(qn("w:t"))
        if not any(a.tag == qn("w:del") for a in t.iterancestors())
    ).strip()


def _set_table_cell_value(value_tc, value: str, *, fallback_tc=None) -> None:
    paragraphs = value_tc.findall(qn("w:p"))
    if not paragraphs:
        p0 = OxmlElement("w:p")
        value_tc.append(p0)
        paragraphs = [p0]
    p0 = paragraphs[0]
    val_rpr = None
    old_slot = _cell_text(value_tc)
    for r in p0.findall(qn("w:r")):
        if not _is_opaque_run(r) and r.find(qn("w:rPr")) is not None:
            val_rpr = r.find(qn("w:rPr"))
            break
    if val_rpr is None:
        val_rpr = p0.find("./" + qn("w:pPr") + "/" + qn("w:rPr"))
    if val_rpr is None and fallback_tc is not None:
        # Inherit font family, size, and color from the label cell without forcing label bold.
        for fb_r in fallback_tc.iter(qn("w:r")):
            fb_rpr = fb_r.find(qn("w:rPr"))
            if fb_rpr is not None:
                val_rpr = OxmlElement("w:rPr")
                for tag_name in ("rFonts", "sz", "szCs", "color"):
                    sub = fb_rpr.find(qn("w:" + tag_name))
                    if sub is not None:
                        val_rpr.append(copy.deepcopy(sub))
                break
    for extra_p in paragraphs[1:]:
        value_tc.remove(extra_p)
    for r in list(p0.findall(qn("w:r"))):
        if not _is_opaque_run(r):
            p0.remove(r)
    formatted_val, final_rpr = _format_slot_value(str(value), old_slot, val_rpr)
    if formatted_val:
        p0.append(_make_text_run(formatted_val, final_rpr))


def _clear_continuation_cell(tc) -> None:
    paragraphs = tc.findall(qn("w:p"))
    for extra_p in paragraphs[1:]:
        tc.remove(extra_p)
    if paragraphs:
        for r in list(paragraphs[0].findall(qn("w:r"))):
            if not _is_opaque_run(r):
                paragraphs[0].remove(r)


def _is_vmerge_continue(tc) -> bool:
    vmerge = tc.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
    if vmerge is None:
        return False
    val = vmerge.get(qn("w:val"))
    return val in {None, "", "continue"}


def _fill_cover_tables(nodes, compiled_fields, all_cell_res) -> set[str]:
    matched_keys: set[str] = set()
    tables = []
    for node in nodes:
        if node.tag == qn("w:tbl"):
            tables.append(node)
        tables.extend(t for t in node.iter(qn("w:tbl")) if t is not node)
    for tbl in tables:
        rows = tbl.findall(qn("w:tr"))
        for r_idx, row in enumerate(rows):
            cells = row.findall(qn("w:tc"))
            if len(cells) < 2:
                continue
            c_idx = 0
            while c_idx + 1 < len(cells):
                label_tc = cells[c_idx]
                raw_label = _cell_text(label_tc)
                matched = False
                if raw_label:
                    hit = next(
                        (item for item in compiled_fields if item[3].match(raw_label)),
                        None,
                    )
                    if hit is None:
                        hit = next(
                            (item for item in compiled_fields if item[4].match(raw_label)),
                            None,
                        )
                    if hit is not None:
                        key, val = hit[0], hit[1]
                        value_tc = cells[c_idx + 1]
                        _set_table_cell_value(value_tc, str(val), fallback_tc=label_tc)
                        matched_keys.add(key)
                        matched = True
                        # Clear multi-row vMerge continuation slots (e.g. 2-line title cells).
                        next_r = r_idx + 1
                        while next_r < len(rows):
                            next_cells = rows[next_r].findall(qn("w:tc"))
                            if (
                                c_idx + 1 < len(next_cells)
                                and _is_vmerge_continue(next_cells[c_idx])
                                and not _cell_text(next_cells[c_idx])
                                and not any(
                                    cre.match(_cell_text(next_cells[c_idx + 1]))
                                    for cre in all_cell_res
                                )
                            ):
                                _clear_continuation_cell(next_cells[c_idx + 1])
                                next_r += 1
                            else:
                                break
                c_idx += 2 if matched else 1
    return matched_keys


def _remap_cover_relationships(src_part, dst_part, nodes):
    copied_parts = {}
    used_names = {str(p.partname) for p in dst_part.package.parts}
    counter = 1
    for root in nodes:
        for sect in root.iter(qn("w:sectPr")):
            for child in list(sect):
                if child.tag in {qn("w:headerReference"), qn("w:footerReference")}:
                    sect.remove(child)
                elif child.tag == qn("w:type") and child.get(qn("w:val")) == "continuous":
                    sect.remove(child)
            element(sect, "titlePg")
        for elem in root.iter():
            for attr_name, attr_val in list(elem.attrib.items()):
                if not (
                    attr_name.startswith(f"{{{_REL_NS}}}")
                    or attr_name == f"{{{_OFFICE_NS}}}relid"
                ):
                    continue
                if attr_val not in src_part.rels:
                    continue
                rel = src_part.rels[attr_val]
                if rel.is_external:
                    new_rid = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    src_target = rel.target_part
                    src_key = str(src_target.partname)
                    if src_key not in copied_parts:
                        suffix = Path(src_key).suffix or ".bin"
                        folder = "/word/embeddings" if "/embeddings/" in src_key else "/word/media"
                        while True:
                            candidate = f"{folder}/wfm_cover_{counter}{suffix}"
                            counter += 1
                            if candidate not in used_names:
                                used_names.add(candidate)
                                break
                        copied_parts[src_key] = Part(
                            PackURI(candidate),
                            src_target.content_type,
                            src_target.blob,
                            dst_part.package,
                        )
                    new_rid = dst_part.relate_to(copied_parts[src_key], rel.reltype)
                elem.set(attr_name, new_rid)


def _materialize_rpr(src_doc, p, r, rpr) -> None:
    ea = read_font(src_doc, p, r, "font_east_asia")
    if isinstance(ea, str) and ea and not ea.startswith("theme:"):
        write_font(rpr, "font_east_asia", ea)
    lat = read_font(src_doc, p, r, "font_latin")
    if isinstance(lat, str) and lat and not lat.startswith("theme:"):
        write_font(rpr, "font_latin", lat)
    elif isinstance(lat, dict):
        fonts = element(rpr, "rFonts")
        for attr in ("ascii", "hAnsi"):
            val = lat.get(attr)
            if isinstance(val, str) and val and not val.startswith("theme:"):
                fonts.set(qn("w:" + attr), val)
                fonts.attrib.pop(qn("w:" + attr + "Theme"), None)
    sz = read_font(src_doc, p, r, "font_size_pt")
    if isinstance(sz, (int, float)):
        write_font(rpr, "font_size_pt", float(sz))
    write_font(rpr, "bold", bool(read_font(src_doc, p, r, "bold")))
    write_font(rpr, "italic", bool(read_font(src_doc, p, r, "italic")))
    col = read_font(src_doc, p, r, "color")
    if isinstance(col, str) and col and not col.startswith("theme:"):
        write_font(rpr, "color", col)
    if rpr.find(qn("w:u")) is None:
        for pr in rpr_sources(src_doc, p, r):
            if pr is not None and pr.find(qn("w:u")) is not None:
                rpr.append(copy.deepcopy(pr.find(qn("w:u"))))
                break


def _materialize_cover_fonts(src_doc, nodes) -> None:
    paragraphs = []
    for node in nodes:
        if node.tag == qn("w:p"):
            paragraphs.append(node)
        paragraphs.extend(p for p in node.iter(qn("w:p")) if p is not node)
    for p in paragraphs:
        ppr = element(p, "pPr")
        existing_spacing = ppr.find(qn("w:spacing")) is not None
        for prop in ("alignment", "space_before_pt", "space_after_pt", "line_spacing"):
            if (prop == "alignment" and ppr.find(qn("w:jc")) is not None) or (
                prop != "alignment" and existing_spacing
            ):
                continue
            val = read_paragraph(src_doc, p, prop)
            if val is not None and not (isinstance(val, dict) and val.get("kind") == "auto"):
                write_paragraph(ppr, prop, val)
        _materialize_rpr(src_doc, p, None, element(ppr, "rPr"))
        for r in p.findall(qn("w:r")):
            if not _is_opaque_run(r):
                _materialize_rpr(src_doc, p, r, element(r, "rPr"))


def _has_top_level_page_break(p) -> bool:
    return any(
        br.get(qn("w:type")) == "page"
        and not any(a.tag == qn("w:txbxContent") for a in br.iterancestors())
        for br in p.iter(qn("w:br"))
    )


def _strip_top_level_page_breaks(p) -> None:
    for br in list(p.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page" and not any(
            a.tag == qn("w:txbxContent") for a in br.iterancestors()
        ):
            parent = br.getparent()
            if parent is not None:
                parent.remove(br)


def _is_empty_paragraph(p) -> bool:
    if _paragraph_top_text(p).strip():
        return False
    if p.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
        return False
    return not any(
        node.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")} for node in p.iter()
    )


def _has_declaration_module(doc) -> bool:
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style_name = p.style.name if p.style else ""
        mod = classify_module(text, style_name, None)
        if mod == "declaration":
            return True
        if mod in {"abstract", "toc", "chapters"}:
            break
    return False


def _is_cover_boundary_paragraph(
    doc,
    p,
    *,
    has_cover_content: bool,
    saw_page_break: bool,
    include_declaration: bool = False,
) -> bool:
    text = _paragraph_top_text(p).strip()
    if (
        not text
        or _STANDALONE_DATE_RE.match(text)
        or _STANDALONE_TITLE_PLACEHOLDER_RE.match(text)
    ):
        return False
    style = paragraph_style(doc, p)
    style_name = style.name if style is not None else ""
    outline = outline_level(doc, p)
    mod = classify_module(text, style_name, outline)
    non_cover = (
        _NON_COVER_MODULES - {"declaration"}
        if include_declaration
        else _NON_COVER_MODULES
    )
    if mod in non_cover:
        return True
    if mod == "chapters":
        if _COVER_BANNER_RE.search(text) and not _EXPLICIT_CHAPTER_START_RE.match(text):
            return False
        explicit_h1_style = bool(
            (outline is not None and int(outline) == 0)
            or re.search(r"(?:heading|标题)1$", re.sub(r"\s+", "", style_name).lower())
        )
        page_break_before = bool(read_paragraph(doc, p, "page_break_before"))
        if has_cover_content and (
            saw_page_break
            or page_break_before
            or explicit_h1_style
            or _EXPLICIT_CHAPTER_START_RE.match(text)
        ):
            return True
    return False


def _extract_cover_nodes(
    src_doc, *, use_section_break: bool = True, include_declaration: bool = True
) -> list:
    cloned = []
    found_sect = False
    has_cover_content = False
    saw_page_break = False
    src_has_decl = _has_declaration_module(src_doc) if include_declaration else False
    seen_declaration = False

    for child in src_doc.element.body:
        if child.tag == qn("w:sectPr"):
            break
        text = _paragraph_top_text(child).strip() if child.tag == qn("w:p") else ""
        if child.tag == qn("w:p") and _is_cover_boundary_paragraph(
            src_doc,
            child,
            has_cover_content=has_cover_content,
            saw_page_break=saw_page_break,
            include_declaration=src_has_decl,
        ):
            break
        node = copy.deepcopy(child)
        cloned.append(node)
        if child.tag == qn("w:tbl"):
            has_cover_content = True
        elif child.tag == qn("w:p"):
            style = paragraph_style(src_doc, child)
            if classify_module(text, style.name if style is not None else "",
                               outline_level(src_doc, child)) == "declaration":
                seen_declaration = True
            if text or any(
                n.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")} for n in child.iter()
            ):
                has_cover_content = True
            if _has_top_level_page_break(child):
                saw_page_break = True
            if child.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
                if not src_has_decl or seen_declaration:
                    found_sect = True
                    break

    if not cloned:
        raise FormatError("Cover template contains no cover elements")

    if use_section_break:
        if not found_sect:
            # Remove redundant trailing page break before synthesizing nextPage sectPr.
            if cloned[-1].tag == qn("w:p") and _has_top_level_page_breaks_in(cloned[-1]):
                _strip_top_level_page_breaks(cloned[-1])
                if len(cloned) > 1 and _is_empty_paragraph(cloned[-1]):
                    cloned.pop()
            body_sect = src_doc.element.body.find(qn("w:sectPr"))
            sect_p = OxmlElement("w:p")
            ppr = OxmlElement("w:pPr")
            ppr.append(
                copy.deepcopy(body_sect) if body_sect is not None else OxmlElement("w:sectPr")
            )
            sect_p.append(ppr)
            cloned.append(sect_p)
    else:
        if found_sect and cloned[-1].tag == qn("w:p"):
            ppr = cloned[-1].find(qn("w:pPr"))
            sect = ppr.find(qn("w:sectPr")) if ppr is not None else None
            if sect is not None:
                ppr.remove(sect)
            if len(cloned) > 1 and _is_empty_paragraph(cloned[-1]):
                cloned.pop()
        if not (cloned[-1].tag == qn("w:p") and _has_top_level_page_break(cloned[-1])):
            br_p = OxmlElement("w:p")
            br_r = OxmlElement("w:r")
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            br_r.append(br)
            br_p.append(br_r)
            cloned.append(br_p)
    return cloned


def _has_top_level_page_breaks_in(p) -> bool:
    return _has_top_level_page_break(p)


def _remove_existing_cover_from_target(
    target_doc, *, include_declaration: bool = True
) -> bool:
    body = target_doc.element.body
    children = [c for c in body if c.tag != qn("w:sectPr")]
    if not children:
        return False
    has_cover_content = False
    saw_page_break = False
    target_has_decl = _has_declaration_module(target_doc) if include_declaration else False
    seen_declaration = False
    cut_end = 0
    for idx, child in enumerate(children):
        if child.tag == qn("w:p") and _is_cover_boundary_paragraph(
            target_doc,
            child,
            has_cover_content=has_cover_content,
            saw_page_break=saw_page_break,
            include_declaration=target_has_decl,
        ):
            cut_end = idx
            break
        if child.tag == qn("w:tbl"):
            has_cover_content = True
        elif child.tag == qn("w:p"):
            text = _paragraph_top_text(child).strip()
            style = paragraph_style(target_doc, child)
            if classify_module(text, style.name if style is not None else "",
                               outline_level(target_doc, child)) == "declaration":
                seen_declaration = True
            if text or any(
                n.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")} for n in child.iter()
            ):
                has_cover_content = True
            if _has_top_level_page_break(child):
                saw_page_break = True
            if (
                child.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None
                and idx + 1 < len(children)
            ):
                if not target_has_decl or seen_declaration:
                    cut_end = idx + 1
                    break
    if cut_end <= 0:
        return False
    for node in children[:cut_end]:
        body.remove(node)
    return True


def _set_whole_paragraph_text(p, value: str) -> bool:
    text_runs = [r for r in p.findall(qn("w:r")) if not _is_opaque_run(r)]
    if not text_runs:
        return False
    r0 = text_runs[0]
    rpr = r0.find(qn("w:rPr"))
    for extra in text_runs[1:]:
        p.remove(extra)
    t_nodes = r0.findall(qn("w:t"))
    for t in t_nodes[1:]:
        r0.remove(t)
    if t_nodes:
        t0 = t_nodes[0]
    else:
        t0 = OxmlElement("w:t")
        r0.append(t0)
    t0.text = str(value)
    if str(value) != str(value).strip() or "  " in str(value):
        t0.set(XML_SPACE, "preserve")
    if rpr is not None:
        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            for sc in rpr.findall(qn("w:szCs")):
                sc.set(qn("w:val"), sz.get(qn("w:val")))
    return True


def clone_template_cover(
    template_path: str | Path,
    target_doc: Document | None = None,
    *,
    fields: dict[str, str] | None = None,
    replacements: dict[str, str] | None = None,
    replace_existing_cover: bool = False,
    use_section_break: bool = True,
    include_declaration: bool = True,
) -> Document:
    """Clone the cover page (tables, VML textboxes, images, underlines, layout) from a template."""
    src_path = Path(template_path).resolve()
    if src_path.suffix.lower() not in SUPPORTED_WORD_SUFFIXES:
        raise FormatError(
            "Cover template must be a DOCX, DOTX, DOC, DOT, or WPS file", "unsupported"
        )
    src_doc = open_word_document(src_path)
    if target_doc is None:
        target_doc = Document()
        for child in list(target_doc.element.body):
            if child.tag != qn("w:sectPr"):
                target_doc.element.body.remove(child)
    elif replace_existing_cover:
        _remove_existing_cover_from_target(
            target_doc, include_declaration=include_declaration
        )

    cloned = _extract_cover_nodes(
        src_doc, use_section_break=use_section_break, include_declaration=include_declaration
    )
    _materialize_cover_fonts(src_doc, cloned)
    _remap_cover_relationships(src_doc.part, target_doc.part, cloned)

    fields = fields or {}
    if not isinstance(fields, dict):
        raise FormatError("Cover fields must be an object mapping field names to text values")
    compiled_fields = []
    for key, val in fields.items():
        if val is None:
            continue
        canonical, exact_cell_re, cell_re, inline_re = _field_patterns(key)
        compiled_fields.append((key, str(val), canonical, exact_cell_re, cell_re, inline_re))
    if not any(item[2] == "signature" for item in compiled_fields):
        student_val = next((item[1] for item in compiled_fields if item[2] == "student"), None)
        if student_val:
            canonical, exact_cell_re, cell_re, inline_re = _field_patterns("本人签名")
            compiled_fields.append(("signature", student_val, canonical, exact_cell_re, cell_re, inline_re))

    replacements = replacements or {}
    if not isinstance(replacements, dict):
        raise FormatError("Cover replacements must be an object mapping old strings to new strings")

    all_cell_res = [
        re.compile(rf"^[\s\u3000]*(?:{core}){_SP}[:：]?[\s\u3000]*$", re.I)
        for core in _FIELD_LABEL_CORES.values()
    ] + [item[4] for item in compiled_fields]
    any_inline_re = _build_any_inline_label_re([item[0] for item in compiled_fields])

    matched_keys = _fill_cover_tables(cloned, compiled_fields, all_cell_res)

    all_paragraphs = []
    for node in cloned:
        if node.tag == qn("w:p"):
            all_paragraphs.append(node)
        all_paragraphs.extend(p for p in node.iter(qn("w:p")) if p is not node)

    inline_fields = sorted(compiled_fields, key=lambda item: (item[2] is not None,))

    for p in all_paragraphs:
        raw_text = "".join(
            t.text or "" for r in p.findall(qn("w:r")) if not _is_opaque_run(r) for t in r.findall(qn("w:t"))
        ).strip()
        if not raw_text:
            continue
        handled_whole_line = False
        for key, val, canonical, *_ in compiled_fields:
            if canonical == "date" and _STANDALONE_DATE_RE.match(raw_text):
                if _set_whole_paragraph_text(p, val):
                    matched_keys.add(key)
                    handled_whole_line = True
                break
        if not handled_whole_line:
            for key, val, _, _, _, inline_re in inline_fields:
                current_text = "".join(
                    t.text or ""
                    for r in p.findall(qn("w:r"))
                    if not _is_opaque_run(r)
                    for t in r.findall(qn("w:t"))
                )
                if inline_re.search(current_text):
                    if _set_labeled_paragraph_value(
                        p, val, inline_re=inline_re, any_label_re=any_inline_re
                    ):
                        matched_keys.add(key)
        # If date is provided and paragraph has declaration signature with a date placeholder, update it.
        has_repl_for_date = any("年" in old and "月" in old for old in (replacements or {}))
        sig_date_val = next((item[1] for item in compiled_fields if item[2] == "signature_date"), None)
        if not sig_date_val and not has_repl_for_date:
            sig_date_val = next((item[1] for item in compiled_fields if item[2] == "date"), None)
        if sig_date_val and not has_repl_for_date:
            cur_p_text = "".join(
                t.text or "" for r in p.findall(qn("w:r")) if not _is_opaque_run(r) for t in r.findall(qn("w:t"))
            )
            if re.search(r"本人签名|声明人|签字|签名", cur_p_text):
                _replace_paragraph_text_pattern(
                    p,
                    r"(?:20\d\d|\d{2,4})\s*年\s*月\s*日|(?:20\d\d|\d{2,4})\s*年月日|年\s*月\s*日",
                    sig_date_val,
                )
        if replacements:
            for old, new in replacements.items():
                replaced_in_single_run = False
                for r in p.findall(qn("w:r")):
                    if _is_opaque_run(r):
                        continue
                    for t in r.findall(qn("w:t")):
                        if t.text and old in t.text:
                            t.text = t.text.replace(old, str(new))
                            replaced_in_single_run = True
                if not replaced_in_single_run:
                    _replace_paragraph_text_pattern(p, re.escape(old), str(new))

    # Fallback for standalone centered title placeholders (e.g. "论文题目" without colon or table).
    for key, val, canonical, *_ in compiled_fields:
        if canonical == "title" and key not in matched_keys:
            for node in cloned:
                if node.tag != qn("w:p"):
                    continue
                raw_text = _paragraph_top_text(node).strip()
                if _STANDALONE_TITLE_PLACEHOLDER_RE.match(raw_text):
                    if _set_whole_paragraph_text(node, val):
                        matched_keys.add(key)
                        break

    body = target_doc.element.body
    for insert_idx, node in enumerate(cloned):
        body.insert(insert_idx, node)
    return target_doc


def apply_template_cover_file(
    template_path: str | Path,
    destination: str | Path,
    *,
    into_path: str | Path | None = None,
    fields: dict[str, str] | None = None,
    replacements: dict[str, str] | None = None,
    replace_existing_cover: bool = False,
    use_section_break: bool = True,
    include_declaration: bool = True,
    overwrite: bool = False,
) -> dict:
    """Safely clone a template cover to ``destination`` while protecting all inputs."""
    import os
    import uuid

    from .contracts import SCHEMA_VERSION, output_path, sha256, source_path

    tpl_path = Path(template_path).resolve()
    if not tpl_path.is_file():
        raise FormatError(f"Cover template does not exist: {tpl_path}")
    tpl_hash = sha256(tpl_path)

    out_path = Path(destination).resolve()
    if out_path.suffix.lower() != ".docx":
        raise FormatError("Cover output path must end with .docx")
    if out_path == tpl_path or (out_path.exists() and out_path.samefile(tpl_path)):
        raise FormatError("Cover output must not overwrite the template file")

    draft_doc = None
    draft_record = None
    if into_path is not None:
        draft_path = source_path(into_path)
        out_path = output_path(draft_path, out_path)
        draft_hash = sha256(draft_path)
        draft_doc = Document(draft_path)
        draft_record = {"path": str(draft_path), "sha256": draft_hash}

    if out_path.exists() and not overwrite:
        raise FormatError(
            "Output already exists; choose a new filename or explicitly request overwrite"
        )

    result_doc = clone_template_cover(
        tpl_path,
        draft_doc,
        fields=fields,
        replacements=replacements,
        replace_existing_cover=replace_existing_cover,
        use_section_break=use_section_break,
        include_declaration=include_declaration,
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = out_path.with_name(f".{out_path.stem}.{uuid.uuid4().hex}.docx")
    try:
        result_doc.save(temporary)
        if sha256(tpl_path) != tpl_hash:
            raise FormatError("Template changed during cover cloning; output not published")
        if draft_record and sha256(draft_record["path"]) != draft_record["sha256"]:
            raise FormatError("Source draft changed during cover cloning; output not published")
        if out_path.exists() and not overwrite:
            raise FormatError("Output appeared during execution; output not replaced")
        os.replace(temporary, out_path)
    finally:
        if temporary.exists():
            temporary.unlink()

    return {
        "schema_version": SCHEMA_VERSION,
        "kind": "word-format-cover-report",
        "status": "passed",
        "template": {"path": str(tpl_path), "sha256": tpl_hash},
        "input": draft_record,
        "output": str(out_path),
        "output_sha256": sha256(out_path),
        "replace_existing_cover": bool(replace_existing_cover),
        "fields": sorted((fields or {}).keys()),
    }

