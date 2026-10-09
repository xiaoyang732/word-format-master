"""Registered write handlers. Entry points never contain their own format code."""

from __future__ import annotations

import copy

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table

from .contracts import FormatError
from .inspect import _style_paragraph
from .properties import element, simple_range, text_runs, write_font, write_paragraph


def ensure_style(index, ref):
    name=ref["id"][len("style:"):]
    try:
        style=index.document.styles[name]
    except KeyError:
        style=index.document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style=index.document.styles["Normal"]
    index.elements[ref["id"]]=style.element
    if style.type!=WD_STYLE_TYPE.PARAGRAPH:
        raise FormatError("Only paragraph styles are supported", "unsupported")
    index.records.setdefault(ref["id"],{"id":ref["id"],"kind":"style","part":"word/styles.xml","name":name})
    return style


def set_font_property(index, ref, params, name):
    if ref["kind"]=="style":
        pr=element(ensure_style(index,ref).element,"rPr")
        write_font(pr,name,params["value"])
        return
    p=index.elements[ref["id"]]
    runs=simple_range(p,ref["start"],ref["end"],True) if ref["kind"]=="text_range" else list(text_runs(p))
    for run in runs:
        # Empty runs may contain bookmarks/fields. Do not create formatting on them.
        if any(t.text for t in run.findall(qn("w:t"))):
            write_font(element(run,"rPr"),name,params["value"])
    if not any(t.text for run in runs for t in run.findall(qn("w:t"))):
        write_font(element(element(p,"pPr"),"rPr"),name,params["value"])


def set_east_asia_font(index,ref,params): set_font_property(index,ref,params,"font_east_asia")
def set_latin_font(index,ref,params): set_font_property(index,ref,params,"font_latin")
def set_font_size(index,ref,params): set_font_property(index,ref,params,"font_size_pt")
def set_bold(index,ref,params): set_font_property(index,ref,params,"bold")
def set_italic(index,ref,params): set_font_property(index,ref,params,"italic")
def set_font_color(index,ref,params): set_font_property(index,ref,params,"color")


def set_paragraph_property(index,ref,params,name):
    node=ensure_style(index,ref).element if ref["kind"]=="style" else index.elements[ref["id"]]
    value=params if name=="line_spacing" or name.endswith("indent") else params["value"]
    write_paragraph(element(node,"pPr"),name,value)


def set_alignment(index,ref,params): set_paragraph_property(index,ref,params,"alignment")
def set_spacing_before(index,ref,params): set_paragraph_property(index,ref,params,"space_before_pt")
def set_spacing_after(index,ref,params): set_paragraph_property(index,ref,params,"space_after_pt")
def set_line_spacing(index,ref,params): set_paragraph_property(index,ref,params,"line_spacing")
def set_first_line_indent(index,ref,params): set_paragraph_property(index,ref,params,"first_line_indent")
def set_left_indent(index,ref,params): set_paragraph_property(index,ref,params,"left_indent")
def set_right_indent(index,ref,params): set_paragraph_property(index,ref,params,"right_indent")
def set_hanging_indent(index,ref,params): set_paragraph_property(index,ref,params,"hanging_indent")
def set_keep_with_next(index,ref,params): set_paragraph_property(index,ref,params,"keep_with_next")
def set_keep_together(index,ref,params): set_paragraph_property(index,ref,params,"keep_together")
def set_page_break_before(index,ref,params): set_paragraph_property(index,ref,params,"page_break_before")
def set_widow_control(index,ref,params): set_paragraph_property(index,ref,params,"widow_control")


def assign_style(index,ref,params):
    style_ref={"id":"style:"+params["name"],"kind":"style","part":"word/styles.xml"}
    style=ensure_style(index,style_ref)
    node=element(element(index.elements[ref["id"]],"pPr"),"pStyle")
    node.set(qn("w:val"),style.style_id)


def set_page_property(index,ref,params,attr):
    node=index.elements[ref["id"]]
    tag="pgSz" if attr in {"w","h","orient"} else "pgMar"
    value=params["value"]
    element(node,tag).set(qn("w:"+attr), value if attr=="orient" else str(int(round(value*1440/25.4))))


def set_page_width(index,ref,params): set_page_property(index,ref,params,"w")
def set_page_height(index,ref,params): set_page_property(index,ref,params,"h")
def set_page_orientation(index,ref,params): set_page_property(index,ref,params,"orient")
def set_top_margin(index,ref,params): set_page_property(index,ref,params,"top")
def set_bottom_margin(index,ref,params): set_page_property(index,ref,params,"bottom")
def set_left_margin(index,ref,params): set_page_property(index,ref,params,"left")
def set_right_margin(index,ref,params): set_page_property(index,ref,params,"right")
def set_gutter(index,ref,params): set_page_property(index,ref,params,"gutter")
def set_header_distance(index,ref,params): set_page_property(index,ref,params,"header")
def set_footer_distance(index,ref,params): set_page_property(index,ref,params,"footer")


def set_repeat_header(index,ref,params):
    tbl=index.elements[ref["id"]]
    rows=tbl.findall(qn("w:tr"))
    if not rows: raise FormatError("Table has no header row", "unsupported")
    pr=element(rows[0],"trPr")
    node=pr.find(qn("w:tblHeader"))
    if params["value"]:
        element(pr,"tblHeader").set(qn("w:val"),"1")
    elif node is not None:
        pr.remove(node)


def set_table_alignment(index,ref,params):
    element(element(index.elements[ref["id"]],"tblPr"),"jc").set(qn("w:val"),params["value"])


def apply_three_line_table(index,ref,params):
    from apply_spec import apply_tables
    class Scope:
        tables=[Table(index.elements[ref["id"]],index.document._body)]
    apply_tables(Scope(),{}, {"three_line_table":params["value"]})


def configure_headers(index,ref,params):
    from apply_spec import apply_headers_footers
    return apply_headers_footers(index.document,params)


def configure_page_numbers(index,ref,params):
    from apply_spec import apply_page_numbers
    return apply_page_numbers(index.document,params)


def configure_toc(index,ref,params):
    from apply_spec import apply_table_of_contents
    return apply_table_of_contents(index.document,params)


def configure_heading_numbering(index,ref,params):
    from apply_spec import normalize_heading_styles
    return normalize_heading_styles(index.document,params)


def configure_lists(index,ref,params):
    from apply_spec import apply_numbered_lists
    return apply_numbered_lists(index.document,{"numbered":params})


def convert_reference_numbering(index,ref,params):
    from apply_spec import apply_reference_style
    return apply_reference_style(index.document,{},params)


def set_caption_position(index,ref,params):
    from apply_spec import apply_caption_positions
    return apply_caption_positions(index.document,{params["role"]:{"label":params.get("label",""),"position":params["position"]}})


def insert_citations(index,ref,params):
    from citation_workflow import apply_citations
    return apply_citations(index.document,params)
