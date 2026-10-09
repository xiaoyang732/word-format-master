"""Stable source-bound object indices shared by selection, execution and QA."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from document_structure import MODULE_MARKERS, classify_module, heading_level
from .contracts import FormatError, json_hash, sha256, source_path
from .properties import read_font, read_paragraph, text_runs, xml_value


def paragraph_text(p):
    return "".join(n.text or "" for n in p.iter(qn("w:t")) if not any(
        a.tag in {qn("w:del"), qn("w:txbxContent")} for a in n.iterancestors() if a is not p))


class DocumentIndex:
    def __init__(self, document):
        self.document = document
        self.records: dict[str, dict] = {}
        self.elements: dict[str, object] = {}
        self.parents: dict[str, object] = {}
        self.parts = {str(p.partname).lstrip("/"): p for p in document.part.package.parts}
        self.top_ids = []
        self._add("document", "document", "word/document.xml", document.element, {})
        body = document.element.body
        for i, p in enumerate(body.iter(qn("w:p"))):
            top = p.getparent() is body
            style_node = p.find("./" + qn("w:pPr") + "/" + qn("w:pStyle"))
            style_id = style_node.get(qn("w:val")) if style_node is not None else "Normal"
            style = next((s for s in document.styles if s.style_id == style_id), None)
            style_name = style.name if style is not None else style_id
            text = paragraph_text(p)
            level = heading_level(text.strip(), style_name)
            special = classify_module(text.strip(), style_name)
            role = "heading" if level or special in MODULE_MARKERS else "body"
            normalized = style_name.lower().replace(" ", "")
            if normalized in {"title", "subtitle"}:
                role = "title"
            elif "caption" in normalized or re.match(r"^\s*(?:图|表|Figure|Fig\.|Table)\s*\d+", text, re.I):
                role = "caption"
            elif normalized.startswith("toc") or normalized.startswith("wfmtoc"):
                role = "toc"
            elif "list" in normalized or p.find("./" + qn("w:pPr") + "/" + qn("w:numPr")) is not None:
                role = "list" if not level else "heading"
            if not top:
                role = "table_text"
            unsafe = any(a.tag in {qn("w:sdt"), qn("w:ins"), qn("w:del"), qn("w:txbxContent")} for a in p.iterancestors())
            oid = f"p{i}"
            self._add(oid, "paragraph", "word/document.xml", p, {
                "text": text, "text_sha256": hashlib.sha256(text.encode()).hexdigest(),
                "style_id": style_id, "style_name": style_name, "role": role,
                "heading_level": level, "top_level": top, "unsupported_container": unsafe,
                "module": special if special in MODULE_MARKERS else None, "chapter_id": None,
            })
            if top:
                self.top_ids.append(oid)
        chapter = None
        module = None
        module_counts={};module_occurrence=0
        for oid in self.top_ids:
            record = self.records[oid]
            if record["module"]:
                module = record["module"]
                module_counts[module]=module_counts.get(module,0)+1
                module_occurrence=module_counts[module]
                chapter = None
            elif record["heading_level"] == 1:
                module = "chapters"
                chapter = oid
            record["module"] = module
            record["module_occurrence"]=module_occurrence
            record["chapter_id"] = chapter
            if module == "references" and record["role"] == "body":
                record["role"] = "reference"
            elif module in {"abstract", "keywords", "cover", "toc", "declaration", "symbols", "acknowledgements", "appendix"} and record["role"] == "body":
                record["role"] = "module_body"
        for i, tbl in enumerate(body.iter(qn("w:tbl"))):
            pid = next((oid for oid,e in self.elements.items() if e is tbl.getparent()), None)
            self._add(f"t{i}", "table", "word/document.xml", tbl, {"parent_id":pid,"top_level":tbl.getparent() is body})
        for i, section in enumerate(document.sections):
            self._add(f"s{i}", "section", "word/document.xml", section._sectPr, {"number":i+1})
        for name, part in list(self.parts.items()):
            if re.fullmatch(r"word/(header|footer)\d+\.xml", name):
                for i,p in enumerate(part.element.iter(qn("w:p"))):
                    self._add(f"{name}:p{i}", "paragraph", name, p, {
                        "text":paragraph_text(p), "role":"header" if "header" in name else "footer",
                        "text_sha256":hashlib.sha256(paragraph_text(p).encode()).hexdigest(),
                        "unsupported_container":False, "top_level":False,
                    })
        for style in document.styles:
            self._add(f"style:{style.name}", "style", "word/styles.xml", style.element, {"name":style.name,"style_type":int(style.type)})

    def _add(self, oid, kind, part, node, extra):
        self.records[oid] = {"id":oid,"kind":kind,"part":part,"xml_sha256":json_hash(xml_value(node)), **extra}
        self.elements[oid] = node

    def public(self):
        return list(self.records.values())

    def ref(self, oid):
        record = self.records[oid]
        return {key:record[key] for key in ("id","kind","part","xml_sha256")}

    def paragraph(self, oid):
        record = self.records[oid]
        part = self.parts[record["part"]]
        return Paragraph(self.elements[oid], part)

    def font_values(self, oid, name, start=None, end=None):
        p = self.elements[oid]
        if self.records[oid]["kind"] == "style":
            return read_font(self.document, _style_paragraph(self.records[oid]["name"], self.document), None, name)
        values = []
        offset = 0
        for r in text_runs(p):
            length = sum(len(t.text or "") for t in r.findall(qn("w:t")))
            if length and (start is None or (offset < end and offset + length > start)):
                value = read_font(self.document,p,r,name)
                if value not in values:
                    values.append(value)
            offset += length
        return values or [read_font(self.document,p,None,name)]


def _style_paragraph(name, document):
    from docx.oxml import OxmlElement
    p=OxmlElement("w:p")
    pr=OxmlElement("w:pPr")
    st=OxmlElement("w:pStyle")
    st.set(qn("w:val"), document.styles[name].style_id)
    pr.append(st); p.append(pr)
    return p


def inspect_document(path):
    source=source_path(path)
    document=Document(source)
    index=DocumentIndex(document)
    records=index.public()
    for record in records:
        if record["kind"] == "paragraph":
            p=index.elements[record["id"]]
            record["format"]={name:index.font_values(record["id"],name) for name in ("font_size_pt","font_east_asia","font_latin","bold","italic")}
            record["format"].update({name:read_paragraph(document,p,name) for name in ("alignment","space_before_pt","space_after_pt","line_spacing","first_line_indent")})
    return {"schema_version":"2.0","source":{"path":str(source),"sha256":sha256(source)},"objects":records,
            "counting_rule":"Paragraph ordinal is 1-based among non-empty paragraphs after role/chapter filters; table paragraphs excluded unless explicitly selected.",
            "limits":["Page-number selectors need renderer anchors; unsupported here.","Text ranges with fields, drawings, tabs or revisions are rejected."]}
