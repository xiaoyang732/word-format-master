"""Deterministic, unambiguous selectors. Never infer a different target on failure."""

from __future__ import annotations

from docx.oxml.ns import qn

from .contracts import FormatError, number, require_keys, text_value
from .properties import simple_range


def resolve_targets(index, target, allowed):
    require_keys(target,{"type","id","ids","role","level","chapter_id","chapter_text","text","contains","ordinal","all","start","end","quote","occurrence","table_id","row","cell","name","part","module","number"},{"type"},"target")
    kind=target["type"]
    if kind == "text_range":
        if kind not in allowed: raise FormatError("Operation does not support text ranges", "unsupported")
        if "quote" in target and any(k in target for k in ("start","end")):
            raise FormatError("Choose quote or offsets for a text range")
        base={k:v for k,v in target.items() if k not in {"start","end","quote","occurrence"}}
        base["type"]="paragraph"
        refs=resolve_targets(index,base,{"paragraph"})
        if len(refs)!=1:
            raise FormatError("A text range must resolve to one paragraph", "needs_clarification")
        ref=refs[0]
        text=index.records[ref["id"]]["text"]
        if "quote" in target:
            quote=text_value(target["quote"],"quote")
            starts=[]; start=0
            while True:
                pos=text.find(quote,start)
                if pos<0: break
                starts.append(pos); start=pos+1
            if not starts:
                raise FormatError("Requested quote was not found", "needs_clarification")
            if len(starts)>1 and "occurrence" not in target:
                raise FormatError("Quote occurs more than once; specify occurrence", "needs_clarification")
            occurrence=int(number(target.get("occurrence",1),1,len(starts),"occurrence",True))
            start=starts[occurrence-1];end=start+len(quote)
        else:
            start=int(number(target.get("start"),0,len(text),"start",True))
            end=int(number(target.get("end"),start+1,len(text),"end",True))
        simple_range(index.elements[ref["id"]],start,end)
        return [{**ref,"kind":"text_range","start":start,"end":end,"quote":text[start:end]}]
    if kind not in allowed:
        raise FormatError(f"Operation does not support target type {kind!r}","unsupported")
    if kind=="document":
        if set(target)!={"type"}:
            raise FormatError("Document target does not accept filters")
        return [index.ref("document")]
    if kind=="style":
        if set(target)!={"type","name"}:
            raise FormatError("Style target requires only name")
        name=text_value(target["name"],"style name")
        oid="style:"+name
        if oid in index.records and index.records[oid].get("style_type")!=1:
            raise FormatError("Only paragraph styles are supported", "unsupported")
        return [index.ref(oid)] if oid in index.records else [{"id":oid,"kind":"style","part":"word/styles.xml","xml_sha256":None}]
    if "all" in target and not isinstance(target["all"],bool):
        raise FormatError("target.all must be boolean")
    allowed_fields={"paragraph":{"role","level","chapter_id","chapter_text","text","contains","ordinal","all","part","module","table_id","row","cell"},
                    "table":{"ordinal","all"},"section":{"number","all"}}
    extras=set(target)-({"type","id","ids"}|allowed_fields.get(kind,set()))
    if extras:
        raise FormatError(f"Invalid filters for {kind}: {sorted(extras)}")
    candidates=[r for r in index.records.values() if r["kind"]==kind]
    if "ids" in target:
        if not isinstance(target["ids"],list) or not target["ids"] or any(not isinstance(v,str) for v in target["ids"]) or len(set(target["ids"]))!=len(target["ids"]):
            raise FormatError("target.ids must be a nonempty list of unique IDs")
        candidates=[index.records.get(oid,{}) for oid in target["ids"]]
        if any(r.get("kind")!=kind for r in candidates):
            raise FormatError("One or more target IDs have the wrong kind or no longer exist", "needs_clarification")
    if "id" in target:
        candidates=[r for r in candidates if r["id"]==target["id"]]
    for field in ("role","module","part"):
        if field in target:
            candidates=[r for r in candidates if r.get(field)==target[field]]
    if "level" in target:
        level=number(target["level"],1,9,"level",True)
        candidates=[r for r in candidates if r.get("heading_level")==level]
    chapter=target.get("chapter_id")
    if "chapter_text" in target:
        matches=[r for r in index.records.values() if r.get("heading_level")==1 and r.get("text")==target["chapter_text"]]
        if len(matches)!=1:
            raise FormatError("Chapter title must identify exactly one chapter", "needs_clarification")
        chapter=matches[0]["id"]
    if chapter:
        if chapter not in index.records or index.records[chapter].get("heading_level")!=1:
            raise FormatError("chapter_id must reference a level-1 heading", "needs_clarification")
        candidates=[r for r in candidates if r.get("chapter_id")==chapter]
    if "table_id" in target:
        tid=target["table_id"]
        if index.records.get(tid,{}).get("kind")!="table":
            raise FormatError("table_id does not identify a table", "needs_clarification")
        tbl=index.elements[tid]
        rows=tbl.findall(qn("w:tr"))
        if "row" in target:
            rows=[rows[int(number(target["row"],1,len(rows),"row",True))-1]]
        cells=[cell for row in rows for cell in row.findall(qn("w:tc"))]
        if "cell" in target:
            if "row" not in target: raise FormatError("cell requires row")
            cells=[cells[int(number(target["cell"],1,len(cells),"cell",True))-1]]
        paragraphs={p for cell in cells for p in cell.iter(qn("w:p")) if next(p.iterancestors(qn("w:tbl")),None) is tbl}
        candidates=[r for r in candidates if index.elements[r["id"]] in paragraphs]
    elif "row" in target or "cell" in target:
        raise FormatError("row/cell require table_id")
    if kind=="paragraph" and "table_id" not in target and not any(k in target for k in ("id","ids","part")):
        candidates=[r for r in candidates if r.get("top_level")]
    for field in ("text","contains"):
        if field in target:
            value=text_value(target[field],field)
            candidates=[r for r in candidates if (r.get("text")==value if field=="text" else value in r.get("text",""))]
    if "number" in target:
        n=number(target["number"],1,len(index.document.sections),"section number",True)
        candidates=[r for r in candidates if r.get("number")==n]
    if "ordinal" in target:
        if kind=="paragraph": candidates=[r for r in candidates if r.get("text","").strip()]
        n=int(number(target["ordinal"],1,max(1,len(candidates)),"ordinal",True))
        candidates=candidates[n-1:n]
    if not candidates:
        raise FormatError("No objects match the requested target", "needs_clarification")
    if len(candidates)>1 and not target.get("all") and "ids" not in target:
        raise FormatError(f"Target matches {len(candidates)} objects; specify all, IDs or ordinal", "needs_clarification")
    if any(r.get("unsupported_container") for r in candidates):
        raise FormatError("Target is inside revisions, a text box or a content control", "unsupported")
    return [index.ref(r["id"]) for r in candidates]
