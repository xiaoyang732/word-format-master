"""Legacy Format Spec validation and expansion into registered operations."""

from __future__ import annotations

import copy
import re

from docx.oxml.ns import qn
from config import LIST_NUMBERING_STYLES
from document_structure import validate_module_order
from .contracts import FormatError, boolean, number, require_keys, text_value

FONT = {"font_east_asia":"font.east_asia.set","font_latin":"font.latin.set","font_size_pt":"font.size.set",
        "bold":"font.bold.set","italic":"font.italic.set","color":"font.color.set"}
PARA = {"alignment":"paragraph.alignment.set","space_before_pt":"paragraph.spacing_before.set",
        "space_after_pt":"paragraph.spacing_after.set","line_spacing":"paragraph.line_spacing.set",
        "keep_with_next":"paragraph.keep_with_next.set","keep_together":"paragraph.keep_together.set",
        "page_break_before":"paragraph.page_break_before.set","widow_control":"paragraph.widow_control.set"}
INDENTS = ("first_line_indent","left_indent","right_indent","hanging_indent")
META = {"style_name","source_style_id","numbering"}
TOKENS = set(FONT)|set(PARA)|{f"{n}_{unit}" for n in INDENTS for unit in ("mm","chars")}
PAGE = {"width_mm":"page.width.set","height_mm":"page.height.set","orientation":"page.orientation.set",
        "margin_top_mm":"page.margin_top.set","margin_bottom_mm":"page.margin_bottom.set",
        "margin_left_mm":"page.margin_left.set","margin_right_mm":"page.margin_right.set","gutter_mm":"page.gutter.set",
        "header_distance_mm":"page.header_distance.set","footer_distance_mm":"page.footer_distance.set"}


def token_operations(tokens):
    from .registry import REGISTRY
    for key,value in tokens.items():
        if value is None or key in META: continue
        if key in FONT or key in PARA:
            action=(FONT|PARA)[key]
            if key=="line_spacing":
                require_keys(value,{"kind","value","value_pt","unit"},{"kind"},"line_spacing")
                kind=value["kind"]
                if kind in {"single","1.5","double"}:
                    params={"kind":"multiple","value":{"single":1,"1.5":1.5,"double":2}[kind]}
                else:
                    params={"kind":kind,"value":value.get("value") if kind=="multiple" else value.get("value_pt",value.get("value"))}
            else:
                params={"value":"justify" if key=="alignment" and value=="both" else value}
            yield action,REGISTRY[action].validator(params)
    for name in INDENTS:
        chars=tokens.get(name+"_chars"); mm=tokens.get(name+"_mm")
        if chars is None and mm is None: continue
        # Old snapshots can contain both representations. Only equivalent values
        # are accepted; edited contradictory pairs must be resolved by the caller.
        if chars is not None and mm is not None:
            number(chars,0,500,name+"_chars")
            number(mm,-500 if name=="first_line_indent" else 0,500,name+"_mm")
            pt=tokens.get("font_size_pt",12)
            if abs(chars*pt*25.4/72-mm)>0.1:
                raise FormatError(f"Conflicting units for {name}; supply only chars or mm")
        value=chars if chars is not None else mm
        effective=name
        if name=="first_line_indent" and value<0:
            effective="hanging_indent";value=-value
        # Zero snapshot values for the opposite indent are not active requests.
        other="hanging_indent" if name=="first_line_indent" else "first_line_indent"
        if name in {"first_line_indent","hanging_indent"} and value==0 and any(tokens.get(other+"_"+u,0) for u in ("mm","chars")):
            continue
        if name=="hanging_indent" and value==0 and any(tokens.get("first_line_indent_"+u) is not None for u in ("mm","chars")):
            continue
        action="paragraph."+effective+".set"
        params={"value":value,"unit":"chars" if chars is not None else "mm"}
        yield action,REGISTRY[action].validator(params)


def validate_tokens(tokens,extra=frozenset()):
    require_keys(tokens,TOKENS|META|set(extra),label="format tokens")
    for key in META & tokens.keys():
        if key!="numbering" and tokens[key] is not None: text_value(tokens[key],key)
    for key in TOKENS & tokens.keys():
        if tokens[key] is None: continue # Legacy snapshots use null for absent evidence.
        if key.endswith("_mm") or key.endswith("_chars"):
            number(tokens[key],-500 if key.startswith("first_line") else 0,500,key)
    list(token_operations(tokens))


def validate_section(section,data):
    from .registry import REGISTRY
    if section=="page":
        require_keys(data,set(PAGE)|{"size"},label="page")
        if data.get("size") not in {None,"A4","Letter","Custom"}: raise FormatError("Unsupported page size")
        for key,action in PAGE.items():
            if data.get(key) is not None: REGISTRY[action].validator({"value":data[key]})
    elif section in {"body","references"}:
        validate_tokens(data,{"citation_system","citation_mode","numbering_mode"} if section=="references" else set())
        if section=="references" and "numbering_mode" in data and data["numbering_mode"] not in {"none","word-numbering"}:
            raise FormatError("references.numbering_mode must be none or word-numbering")
    elif section=="headers_footers":
        require_keys(data,{"preserve_existing","different_first_page","different_odd_even","header","footer","detected_parts"},label=section)
        for key in ("preserve_existing","different_first_page","different_odd_even"):
            if key in data: boolean(data[key],key)
        for key in ("header","footer"):
            if key in data:
                config=data[key];validate_tokens(config,{"enabled","text"})
                if "enabled" in config: boolean(config["enabled"],key+".enabled")
                if "text" in config: text_value(config["text"],key+".text",True)
    elif section=="page_numbers":
        validate_tokens(data,{"enabled","location","format","start","show_on_first_page","prefix","total_separator"})
        for key in ("enabled","show_on_first_page"):
            if key in data: boolean(data[key],key)
        for key,choices in (("location",{"header","footer"}),("format",{"number","page-number","page-x-of-y"})):
            if key in data and data[key] not in choices: raise FormatError(f"Invalid page_numbers.{key}")
        if "start" in data: number(data["start"],0,100000,"start",True)
        for key in ("prefix","total_separator"):
            if key in data: text_value(data[key],key,True)
    elif section=="table_of_contents":
        require_keys(data,{"enabled","title","max_heading_level","page_break_after"},label=section)
        for key in ("enabled","page_break_after"):
            if key in data: boolean(data[key],key)
        if "title" in data: text_value(data["title"],"title")
        if "max_heading_level" in data: number(data["max_heading_level"],1,9,"max_heading_level",True)
    elif section=="numbered":
        require_keys(data,{"style","start","left_indent_mm","hanging_indent_mm"},label=section)
        if "style" in data and data["style"] not in LIST_NUMBERING_STYLES: raise FormatError("Unsupported list style")
        if "start" in data: number(data["start"],1,100000,"start",True)
        for key in ("left_indent_mm","hanging_indent_mm"):
            if key in data: number(data[key],0,500,key)
    elif section=="heading_numbering":
        require_keys(data,{"level1","level2","level3"},label=section)
        for level,value in data.items():
            if value not in ({"keep","chinese","arabic"} if level=="level1" else {"keep","arabic"}):
                raise FormatError(f"Invalid {level} heading numbering")
    elif section=="reference_numbering":
        require_keys(data,{"numbering_mode"},{"numbering_mode"},section)
        if data["numbering_mode"] not in {"none","word-numbering"}: raise FormatError("Invalid reference numbering mode")
    elif section=="caption_position":
        require_keys(data,{"role","label","position"},{"role","position"},section)
        if data["role"] not in {"figure","table"} or data["position"] not in {"above","below"}: raise FormatError("Invalid caption role/position")
        if "label" in data: text_value(data["label"],"label",True)
    elif section=="citations":
        require_keys(data,{"placements"},{"placements"},section)
        if not isinstance(data["placements"],list): raise FormatError("citations.placements must be an array")
        for placement in data["placements"]:
            require_keys(placement,{"paragraph_index","paragraph_sha256","reference_ids","citation_text","insert","spacing","confidence","reason","char_offset"},
                         {"paragraph_index","paragraph_sha256","reference_ids","citation_text"},"citation placement")
            number(placement["paragraph_index"],0,1000000,"paragraph_index",True)
            if not isinstance(placement["reference_ids"],list) or not placement["reference_ids"]: raise FormatError("reference_ids must be a nonempty array")
            for value in placement["reference_ids"]: text_value(value,"reference ID")
            text_value(placement["paragraph_sha256"],"paragraph_sha256");text_value(placement["citation_text"],"citation_text")
            if "confidence" in placement: number(placement["confidence"],0,1,"confidence")
    else: raise FormatError(f"Unknown section {section}")


def validate_spec(spec):
    require_keys(spec,{"schema_version","id","name","mode","template_required","page","body","headings","lists","captions","tables",
                       "headers_footers","page_numbers","table_of_contents","references","document_structure","citations","authority","notes","unsupported","ai_analysis_summary","headings_numbering_format"},label="spec")
    if spec.get("schema_version","1.0")!="1.0": raise FormatError("Unsupported legacy spec version")
    if spec.get("template_required"): raise FormatError("Official templates require their publisher workflow", "unsupported")
    for section in ("page","body","headers_footers","page_numbers","table_of_contents","references","citations"):
        if section in spec: validate_section(section,spec[section])
    headings=spec.get("headings",[])
    if not isinstance(headings,list): raise FormatError("headings must be an array")
    levels=set()
    for heading in headings:
        validate_tokens(heading,{"level"});level=number(heading.get("level"),1,9,"heading.level",True)
        if level in levels: raise FormatError("Duplicate heading levels")
        levels.add(level)
    lists=spec.get("lists",{})
    require_keys(lists,{"headings","numbered","source_definitions"},label="lists")
    if "headings" in lists: validate_section("heading_numbering",lists["headings"])
    if "numbered" in lists: validate_section("numbered",lists["numbered"])
    if "headings_numbering_format" in spec: validate_section("heading_numbering",spec["headings_numbering_format"])
    require_keys(spec.get("captions",{}),{"figure","table"},label="captions")
    for tokens in spec.get("captions",{}).values():
        validate_tokens(tokens,{"label","position","numbering_mode"})
        if tokens.get("position") not in {None,"above","below"}: raise FormatError("Caption position must be above/below")
    tables=spec.get("tables",{})
    require_keys(tables,{"font_size_pt","font_east_asia","font_latin","alignment","three_line_table","repeat_header_row"},label="tables")
    validate_tokens({k:v for k,v in tables.items() if k not in {"three_line_table","repeat_header_row"}})
    for key in ("three_line_table","repeat_header_row"):
        if key in tables: boolean(tables[key],key)
    structure=spec.get("document_structure",{})
    require_keys(structure,{"schema_version","strict_order","ordered_modules","order_errors"},label="document_structure")
    if "strict_order" in structure: boolean(structure["strict_order"],"strict_order")
    modules=structure.get("ordered_modules",[])
    if not isinstance(modules,list) or any(not isinstance(m,dict) for m in modules): raise FormatError("ordered_modules must be an array of objects")
    errors=validate_module_order(modules,strict=structure.get("strict_order",True))
    if errors: raise FormatError("; ".join(errors))
    for module in modules:
        for role in module.get("style_roles",{}).values():
            validate_tokens(role.get("tokens",{}))


def spec_request(index,spec,clear_direct=False):
    """Expand a full legacy spec once; execution always uses REGISTRY."""
    validate_spec(spec)
    operations=[]
    def add(action,ids,params):
        if not ids: return
        kind=index.records[ids[0]]["kind"] if ids[0] in index.records else "style"
        target={"type":kind,"ids":ids} if kind!="style" else {"type":"style","name":ids[0][6:]}
        if kind=="document": target={"type":"document"}
        operations.append({"action":action,"target":target,"params":params})
    # Resolve per-object intent first: module values provide defaults and explicit
    # global role values win. Never modify a shared style for a module-specific rule.
    effective={}
    modules=spec.get("document_structure",{}).get("ordered_modules",[])
    for oid,r in index.records.items():
        if r["kind"]!="paragraph" or r["part"]!="word/document.xml": continue
        matching=[m for m in modules if m.get("kind")==r.get("module")]
        occurrence=max(0,r.get("module_occurrence",1)-1)
        module=matching[min(occurrence,len(matching)-1)] if matching else None
        if module:
            role="heading" if r.get("role")=="heading" else "title" if r.get("role")=="title" else "keywords" if r.get("module")=="keywords" else "body"
            tokens=module.get("style_roles",{}).get(role,{}).get("tokens",{})
            effective[oid]=dict(tokens)
        role=r.get("role")
        tokens={}
        style_name=None
        if role=="body" or (role=="module_body" and r.get("module") in {"abstract","appendix","acknowledgements"}):
            tokens=spec.get("body",{})
        elif role=="heading":
            tokens=next((h for h in spec.get("headings",[]) if h["level"]==r.get("heading_level")),{})
            if tokens: style_name=f"Heading {tokens['level']}"
        elif role=="reference":
            tokens=spec.get("references",{});style_name="Bibliography" if tokens else None
        elif role=="caption":
            for label_kind,defaults in (("figure",("图","Figure","Fig.")),("table",("表","Table"))):
                cap=spec.get("captions",{}).get(label_kind,{})
                labels=(cap.get("label"),)+defaults
                if cap and any(label and re.match(r"^\s*"+re.escape(label)+r"\s*\d",r.get("text",""),re.I) for label in labels):
                    tokens=cap;style_name=cap.get("style_name") or ("Figure Caption" if label_kind=="figure" else "Table Caption")
        if style_name: add("paragraph.style.assign",[oid],{"name":style_name})
        merged=effective.setdefault(oid,{})
        # Remove alternate unit representation when a higher priority source sets it.
        for name in INDENTS:
            if any(k in tokens for k in (name+"_mm",name+"_chars")):
                for unit in ("mm","chars"): merged.pop(name+"_"+unit,None)
        merged.update(tokens)
    for oid,tokens in effective.items():
        for action,params in token_operations(tokens): add(action,[oid],params)
    # Global named styles mirror the explicit global rules; direct target writes
    # above ensure that existing overrides cannot hide the requested format.
    for name,tokens in [("Normal",spec.get("body",{}))]+[(f"Heading {h['level']}",h) for h in spec.get("headings",[])]+[("Bibliography",spec.get("references",{}))]:
        for action,params in token_operations(tokens): add(action,["style:"+name],params)
    page=spec.get("page",{})
    dims={"A4":(210,297),"Letter":(215.9,279.4)}.get(page.get("size"))
    values={k:v for k,v in page.items() if k in PAGE and v is not None}
    if dims: values.update(width_mm=dims[0],height_mm=dims[1])
    if values.get("orientation")=="landscape" and "width_mm" in values and "height_mm" in values:
        values["width_mm"],values["height_mm"]=max(values["width_mm"],values["height_mm"]),min(values["width_mm"],values["height_mm"])
    for key,value in values.items(): add(PAGE[key],[oid for oid,r in index.records.items() if r["kind"]=="section"],{"value":value})
    tables=spec.get("tables",{})
    for oid,r in index.records.items():
        if r["kind"]=="table":
            for key,action in (("alignment","table.alignment.set"),("repeat_header_row","table.header_repeat.set"),("three_line_table","table.three_line.apply")):
                if key in tables: add(action,[oid],{"value":tables[key]})
        if r.get("role")=="table_text":
            for action,params in token_operations({k:v for k,v in tables.items() if k in FONT}): add(action,[oid],params)
    for section,action in (("headers_footers","header_footer.configure"),("page_numbers","page_number.configure"),("table_of_contents","toc.configure")):
        if spec.get(section): add(action,["document"],spec[section])
    lists=spec.get("lists",{})
    if lists.get("headings") or spec.get("headings_numbering_format"):
        add("heading.numbering.configure",["document"],lists.get("headings",spec.get("headings_numbering_format")))
    if lists.get("numbered"): add("list.numbering.configure",["document"],lists["numbered"])
    if spec.get("references",{}).get("numbering_mode"):
        add("reference.numbering.convert",["document"],{"numbering_mode":spec["references"]["numbering_mode"]})
    for role,tokens in spec.get("captions",{}).items():
        if tokens.get("position"): add("caption.position.set",["document"],{k:v for k,v in {"role":role,"label":tokens.get("label",""),"position":tokens["position"]}.items()})
    if spec.get("citations",{}).get("placements"): add("citation.insert",["document"],spec["citations"])
    # Citation anchors are source indices, and heading conversion can change their
    # text, so explicit citation insertions precede structural normalization.
    operations.sort(key=lambda item: 0 if item["action"]=="citation.insert" else 2 if item["action"]=="toc.configure" else 1)
    return {"schema_version":"2.0","origin":"dashboard","operations":operations,"verification_spec":copy.deepcopy(spec),
            "notes":["Legacy clear_direct is scoped to requested properties; unrelated direct formatting is preserved."] if clear_direct else []}


# Generated from executable actions, not a second independent supported-fields list.
def legacy_paths():
    paths={}
    for prefix in ("body","headings.*","references","captions.figure","captions.table","headers_footers.header","headers_footers.footer","page_numbers"):
        for key,action in (FONT|PARA).items(): paths[f"{prefix}.{key}"]=action
        for name in INDENTS:
            for unit in ("mm","chars"): paths[f"{prefix}.{name}_{unit}"]="paragraph."+name+".set"
    paths.update({"page."+k:v for k,v in PAGE.items()});paths["page.size"]="page.width.set + page.height.set"
    for prefix,fields,action in (
        ("lists.headings",("level1","level2","level3"),"heading.numbering.configure"),
        ("lists.numbered",("style","start","left_indent_mm","hanging_indent_mm"),"list.numbering.configure"),
        ("headers_footers",("preserve_existing","different_first_page","different_odd_even"),"header_footer.configure"),
        ("page_numbers",("enabled","location","format","start","show_on_first_page"),"page_number.configure"),
        ("table_of_contents",("enabled","title","max_heading_level","page_break_after"),"toc.configure"),
    ):
        paths.update({prefix+"."+field:action for field in fields})
    for role in ("figure","table"):
        paths[f"captions.{role}.position"]="caption.position.set"
        paths[f"captions.{role}.label"]="selector metadata (caption matching)"
    for key in ("font_east_asia","font_latin","font_size_pt"): paths["tables."+key]=FONT[key]
    paths.update({"tables.alignment":"table.alignment.set","tables.repeat_header_row":"table.header_repeat.set","tables.three_line_table":"table.three_line.apply",
                  "headers_footers.header.enabled":"header_footer.configure","headers_footers.header.text":"header_footer.configure",
                  "headers_footers.footer.enabled":"header_footer.configure","headers_footers.footer.text":"header_footer.configure",
                  "document_structure":"module role selectors expanded into registered properties",
                  "clear_direct_font_formatting":"requested property overrides only"})
    return paths
