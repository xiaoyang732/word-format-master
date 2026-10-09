"""Executable capabilities: validation, sole writer, independent reader and QA."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Callable

from docx.oxml.ns import qn

from . import operations as op
from .contracts import FormatError, boolean, number, require_keys, text_value
from .inspect import _style_paragraph
from .properties import read_font, read_paragraph


@dataclass(frozen=True)
class Capability:
    action: str
    property: str
    targets: frozenset
    handler: Callable
    validator: Callable
    reader: Callable
    verifier: Callable
    parameters: dict
    structural: bool = False
    effects: tuple = ()

    def public(self):
        return {"action":self.action,"property":self.property,"targets":sorted(self.targets),
                "handler":self.handler.__name__,"validator":self.validator.__name__,"reader":self.reader.__name__,"verifier":self.verifier.__name__,
                "parameters":self.parameters,"structural":self.structural,"effects":list(self.effects),
                "test":"tests/test_headers.py" if self.property.startswith("header:") or self.property=="section_break" else "tests/test_format_core.py"}


REGISTRY: dict[str,Capability] = {}


def scalar(kind, low=None, high=None, unit=None, enum=None):
    def validate(params):
        require_keys(params,{"value","unit"} if unit else {"value"},{"value"},"params")
        value=params["value"]
        if kind=="number": number(value,low,high)
        elif kind=="boolean": boolean(value)
        elif kind=="text": text_value(value)
        if enum and value not in enum: raise FormatError(f"value must be one of {list(enum)}")
        if unit and params.get("unit",unit)!=unit: raise FormatError(f"unit must be {unit}")
        if kind=="number" and unit=="pt" and low==5 and abs(value*2-round(value*2))>1e-9:
            raise FormatError("Word font sizes require half-point precision")
        if kind=="text" and not enum and "color" in validate.__name__: pass
        return {"value":value,**({"unit":unit} if unit else {})}
    return validate


def indent(params):
    require_keys(params,{"value","unit"},{"value","unit"},"indent params")
    if params["unit"] not in {"mm","chars"}: raise FormatError("indent unit must be mm or chars")
    number(params["value"],0,100 if params["unit"]=="chars" else 500)
    return dict(params)


def line_spacing(params):
    require_keys(params,{"kind","value","unit"},{"kind","value"},"line spacing")
    kind=params["kind"]
    if kind not in {"multiple","exact","at_least"}: raise FormatError("line spacing kind must be multiple, exact or at_least")
    unit="multiple" if kind=="multiple" else "pt"
    if params.get("unit",unit)!=unit: raise FormatError("Line spacing mode and unit disagree")
    number(params["value"],0.1 if unit=="multiple" else 1,20 if unit=="multiple" else 500)
    return {"kind":kind,"value":params["value"],"unit":unit}


def color(params):
    require_keys(params,{"value"},{"value"},"color")
    value=text_value(params["value"]).lstrip("#")
    if value!="auto" and not re.fullmatch(r"[0-9a-fA-F]{6}",value): raise FormatError("Color must be RGB hex or auto")
    return {"value":value.upper() if value!="auto" else value}


def read_atomic(index,ref,cap):
    if cap.property in {"font_east_asia","font_latin","font_size_pt","bold","italic","color"}:
        return index.font_values(ref["id"],cap.property,ref.get("start"),ref.get("end"))
    p=index.elements[ref["id"]]
    if ref["kind"]=="style":
        p=_style_paragraph(ref["id"][6:],index.document)
    if cap.property in {"alignment","space_before_pt","space_after_pt","line_spacing","first_line_indent","left_indent","right_indent","hanging_indent","keep_with_next","keep_together","page_break_before","widow_control"}:
        return read_paragraph(index.document,p,cap.property)
    if cap.property.startswith("page:"):
        attr=cap.property[5:]
        tag="pgSz" if attr in {"w","h","orient"} else "pgMar"
        child=p.find(qn("w:"+tag))
        raw=child.get(qn("w:"+attr)) if child is not None else None
        return (raw or "portrait") if attr=="orient" else float(raw)*25.4/1440 if raw else None
    if cap.property=="style":
        child=p.find("./"+qn("w:pPr")+"/"+qn("w:pStyle"))
        sid=child.get(qn("w:val")) if child is not None else "Normal"
        return next((s.name for s in index.document.styles if s.style_id==sid),sid)
    if cap.property=="table_alignment":
        child=p.find("./"+qn("w:tblPr")+"/"+qn("w:jc"))
        return child.get(qn("w:val")) if child is not None else "left"
    if cap.property=="repeat_header":
        child=p.find("./"+qn("w:tr")+"/"+qn("w:trPr")+"/"+qn("w:tblHeader"))
        return child is not None and child.get(qn("w:val"),"1") not in {"0","false"}
    if cap.property=="three_line":
        borders=p.find("./"+qn("w:tblPr")+"/"+qn("w:tblBorders"))
        def val(tag,attr):
            node=borders.find(qn("w:"+tag)) if borders is not None else None
            return node.get(qn("w:"+attr)) if node is not None else None
        header=p.find("./"+qn("w:tr")+"/"+qn("w:tc")+"/"+qn("w:tcPr")+"/"+qn("w:tcBorders")+"/"+qn("w:bottom"))
        return val("top","sz")=="12" and val("bottom","sz")=="12" and val("insideV","val")=="none" and header is not None and header.get(qn("w:sz"))=="6"
    raise FormatError(f"No reader for {cap.property}")


def equal_value(actual,expected,tolerance=0.025):
    if isinstance(actual,dict) and isinstance(expected,dict):
        return actual.keys()==expected.keys() and all(equal_value(actual[k],expected[k],tolerance) for k in actual)
    if isinstance(actual,bool) or isinstance(expected,bool): return actual is expected
    if isinstance(actual,(int,float)) and isinstance(expected,(int,float)): return abs(actual-expected)<=tolerance
    return actual==expected


def verify_atomic(actual,params,cap):
    expected=params if cap.property=="line_spacing" or cap.property.endswith("indent") else params.get("value",params.get("name"))
    if isinstance(actual,list): return bool(actual) and all(equal_value(v,expected) for v in actual)
    if cap.property.endswith("indent") and isinstance(actual,dict) and params["value"]==0:
        return abs(actual["value"])<0.025
    return equal_value(actual,expected)


def add(action,prop,targets,handler,validator,parameters,structural=False,effects=()):
    if action in REGISTRY: raise RuntimeError(f"Duplicate capability {action}")
    REGISTRY[action]=Capability(action,prop,frozenset(targets),handler,validator,read_atomic,verify_atomic,parameters,structural,effects)


for action,prop,handler,validator,params in (
    ("font.east_asia.set","font_east_asia",op.set_east_asia_font,scalar("text"),{"value":"font name"}),
    ("font.latin.set","font_latin",op.set_latin_font,scalar("text"),{"value":"font name"}),
    ("font.size.set","font_size_pt",op.set_font_size,scalar("number",5,96,"pt"),{"value":"5..96 in half-point steps","unit":"pt"}),
    ("font.bold.set","bold",op.set_bold,scalar("boolean"),{"value":"boolean"}),
    ("font.italic.set","italic",op.set_italic,scalar("boolean"),{"value":"boolean"}),
    ("font.color.set","color",op.set_font_color,color,{"value":"RGB hex or auto"}),
): add(action,prop,{"paragraph","text_range","style"},handler,validator,params)

for action,prop,handler,validator,params in (
    ("paragraph.alignment.set","alignment",op.set_alignment,scalar("text",enum=("left","center","right","justify")),{"value":["left","center","right","justify"]}),
    ("paragraph.spacing_before.set","space_before_pt",op.set_spacing_before,scalar("number",0,500,"pt"),{"value":"0..500","unit":"pt"}),
    ("paragraph.spacing_after.set","space_after_pt",op.set_spacing_after,scalar("number",0,500,"pt"),{"value":"0..500","unit":"pt"}),
    ("paragraph.line_spacing.set","line_spacing",op.set_line_spacing,line_spacing,{"kind":["multiple","exact","at_least"],"value":"positive number","unit":"multiple or pt"}),
    ("paragraph.first_line_indent.set","first_line_indent",op.set_first_line_indent,indent,{"value":"nonnegative number","unit":["mm","chars"]}),
    ("paragraph.left_indent.set","left_indent",op.set_left_indent,indent,{"value":"nonnegative number","unit":["mm","chars"]}),
    ("paragraph.right_indent.set","right_indent",op.set_right_indent,indent,{"value":"nonnegative number","unit":["mm","chars"]}),
    ("paragraph.hanging_indent.set","hanging_indent",op.set_hanging_indent,indent,{"value":"nonnegative number","unit":["mm","chars"]}),
    ("paragraph.keep_with_next.set","keep_with_next",op.set_keep_with_next,scalar("boolean"),{"value":"boolean"}),
    ("paragraph.keep_together.set","keep_together",op.set_keep_together,scalar("boolean"),{"value":"boolean"}),
    ("paragraph.page_break_before.set","page_break_before",op.set_page_break_before,scalar("boolean"),{"value":"boolean"}),
    ("paragraph.widow_control.set","widow_control",op.set_widow_control,scalar("boolean"),{"value":"boolean"}),
): add(action,prop,{"paragraph","style"},handler,validator,params)

for action,attr,handler in (
    ("page.width.set","w",op.set_page_width),("page.height.set","h",op.set_page_height),
    ("page.orientation.set","orient",op.set_page_orientation),("page.margin_top.set","top",op.set_top_margin),
    ("page.margin_bottom.set","bottom",op.set_bottom_margin),("page.margin_left.set","left",op.set_left_margin),
    ("page.margin_right.set","right",op.set_right_margin),("page.gutter.set","gutter",op.set_gutter),
    ("page.header_distance.set","header",op.set_header_distance),("page.footer_distance.set","footer",op.set_footer_distance),
): add(action,"page:"+attr,{"section"},handler,scalar("text",enum=("portrait","landscape")) if attr=="orient" else scalar("number",1 if attr in {"w","h"} else 0,2000,"mm"),{"value":["portrait","landscape"]} if attr=="orient" else {"value":"0..2000 mm (dimensions positive)","unit":"mm"})

def style_params(params):
    require_keys(params,{"name"},{"name"},"style params")
    return {"name":text_value(params["name"],"style name")}

add("paragraph.style.assign","style",{"paragraph"},op.assign_style,style_params,{"name":"paragraph style"})
add("table.alignment.set","table_alignment",{"table"},op.set_table_alignment,scalar("text",enum=("left","center","right")),{"value":["left","center","right"]})
add("table.header_repeat.set","repeat_header",{"table"},op.set_repeat_header,scalar("boolean"),{"value":"boolean"})
add("table.three_line.apply","three_line",{"table"},op.apply_three_line_table,scalar("boolean"),{"value":"true to apply; false restores only WFM-saved borders"},True,("table borders",))


def complex_validator(section):
    def validate(params):
        from .spec import validate_section
        validate_section(section,params)
        if section=="headers_footers" and (set(params)&{"even_header","first_header","sections","section_number"} or
            set(params.get("header",{}))&{"mode","prefix","linked_to_previous"}):
            raise FormatError("Use registered header.* operations or a spec for section/page-variant header configuration", "unsupported")
        return dict(params)
    return validate


def read_complex(index,ref,cap):
    from .verification import read_complex_state
    return read_complex_state(index,cap.property)


def verify_complex(actual,params,cap):
    from .verification import verify_complex_state
    return verify_complex_state(actual,params,cap.property)


for action,prop,handler,section,effects in (
    ("header_footer.configure","headers_footers",op.configure_headers,"headers_footers",("selected header/footer content", "first/odd-even page flags")),
    ("page_number.configure","page_numbers",op.configure_page_numbers,"page_numbers",("managed/reused PAGE fields","first-page flag","page numbering start")),
    ("toc.configure","table_of_contents",op.configure_toc,"table_of_contents",("managed TOC paragraphs","field update setting")),
    ("heading.numbering.configure","heading_numbering",op.configure_heading_numbering,"heading_numbering",("heading styles and numbering","explicit manual number prefixes")),
    ("list.numbering.configure","numbered",op.configure_lists,"numbered",("existing numbered lists and definitions",)),
    ("reference.numbering.convert","reference_numbering",op.convert_reference_numbering,"reference_numbering",("reference numbering and explicit prefixes",)),
    ("caption.position.set","caption_position",op.set_caption_position,"caption_position",("adjacent top-level caption order",)),
    ("citation.insert","citations",op.insert_citations,"citations",("validated citation placement text",)),
):
    REGISTRY[action]=Capability(action,prop,frozenset({"document"}),handler,complex_validator(section),read_complex,verify_complex,
                               {"contract":section,"scope":"document; use atomic properties for local formatting"},True,effects)


def capabilities():
    return {"schema_version":"2.0","operations":[c.public() for c in REGISTRY.values()],
            "unsupported":["page-based selectors without renderer anchors","complex text ranges","automatic prose rewriting"],
            "legacy_spec":"1.0 via the same planner and executor"}


from . import headers
from . import section_breaks

REGISTRY["section.break.next_page.insert"] = Capability(
    "section.break.next_page.insert", "section_break", frozenset({"paragraph"}),
    section_breaks.insert, section_breaks.validate, section_breaks.read, section_breaks.verify,
    {}, True, ("next-page section boundary after selected top-level paragraph; pagination changes",))


def header_variant(validator):
    return lambda params: headers.variant_params(params, validator)


for action, prop, targets, handler, validator, parameters in (
    ("header.content.set", "content", {"section"}, headers.set_content, headers.content_params,
     {"mode": ["text", "chapter_title"], "text": "literal or empty", "style_name": "existing heading style", "prefix": "optional literal"}),
    ("header.link_previous.set", "link", {"section"}, headers.set_link, header_variant(scalar("boolean")), {"value": "boolean"}),
    ("header.first_page_different.set", "first_page", {"section"}, headers.set_first_page, scalar("boolean"), {"value": "boolean"}),
    ("header.odd_even_different.set", "odd_even", {"document"}, headers.set_odd_even, scalar("boolean"), {"value": "boolean"}),
    ("header.font.east_asia.set", "font_east_asia", {"section"}, headers.set_east_asia, header_variant(scalar("text")), {"value": "font name"}),
    ("header.font.latin.set", "font_latin", {"section"}, headers.set_latin, header_variant(scalar("text")), {"value": "font name"}),
    ("header.font.size.set", "font_size_pt", {"section"}, headers.set_size, header_variant(scalar("number", 5, 96, "pt")), {"value": "5..96 in half-point steps", "unit": "pt"}),
    ("header.alignment.set", "alignment", {"section"}, headers.set_alignment, header_variant(scalar("text", enum=("left", "center", "right"))), {"value": ["left", "center", "right"]}),
    ("header.font.bold.set", "bold", {"section"}, headers.set_bold, header_variant(scalar("boolean")), {"value":"boolean"}),
    ("header.font.italic.set", "italic", {"section"}, headers.set_italic, header_variant(scalar("boolean")), {"value":"boolean"}),
    ("header.font.color.set", "color", {"section"}, headers.set_color, header_variant(color), {"value":"RGB or auto"}),
    ("header.spacing_before.set", "space_before_pt", {"section"}, headers.set_spacing_before, header_variant(scalar("number",0,500,"pt")), {"value":"0..500","unit":"pt"}),
    ("header.spacing_after.set", "space_after_pt", {"section"}, headers.set_spacing_after, header_variant(scalar("number",0,500,"pt")), {"value":"0..500","unit":"pt"}),
    ("header.line_spacing.set", "line_spacing", {"section"}, headers.set_line_spacing, header_variant(line_spacing), {"kind":["multiple","exact","at_least"],"value":"positive","unit":"multiple or pt"}),
    ("header.first_line_indent.set", "first_line_indent", {"section"}, headers.set_first_indent, header_variant(indent), {"value":"nonnegative","unit":["chars","mm"]}),
    ("header.left_indent.set", "left_indent", {"section"}, headers.set_left_indent, header_variant(indent), {"value":"nonnegative","unit":["chars","mm"]}),
    ("header.right_indent.set", "right_indent", {"section"}, headers.set_right_indent, header_variant(indent), {"value":"nonnegative","unit":["chars","mm"]}),
    ("header.hanging_indent.set", "hanging_indent", {"section"}, headers.set_hanging_indent, header_variant(indent), {"value":"nonnegative","unit":["chars","mm"]}),
    ("header.keep_with_next.set", "keep_with_next", {"section"}, headers.set_keep_next, header_variant(scalar("boolean")), {"value":"boolean"}),
    ("header.keep_together.set", "keep_together", {"section"}, headers.set_keep_together, header_variant(scalar("boolean")), {"value":"boolean"}),
    ("header.page_break_before.set", "page_break_before", {"section"}, headers.set_page_break, header_variant(scalar("boolean")), {"value":"boolean"}),
    ("header.widow_control.set", "widow_control", {"section"}, headers.set_widow, header_variant(scalar("boolean")), {"value":"boolean"}),
):
    REGISTRY[action] = Capability(action, "header:" + prop, frozenset(targets), handler, validator,
                                 headers.read_header, headers.verify_header,
                                 {**parameters, **({"variant": ["default", "first", "even"]} if prop not in {"first_page", "odd_even"} else {})},
                                 True, ("selected header variant; following inherited header isolated",) if prop not in {"first_page", "odd_even"} else
                                 ("Word page flag affects both header and footer display",))
