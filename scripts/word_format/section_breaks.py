"""Explicit next-page section boundaries; never invent chapter positions."""

import copy

from docx.oxml.ns import qn

from .contracts import FormatError, require_keys
from .properties import element, xml_value


def validate(params):
    require_keys(params, set(), label="next-page section break")
    return {}


def insert(index, ref, params):
    p = index.elements[ref["id"]]
    if p.getparent() is not index.document.element.body:
        raise FormatError("Section breaks require a top-level paragraph", "unsupported")
    pr = element(p, "pPr")
    if pr.find(qn("w:sectPr")) is not None:
        raise FormatError("Paragraph already ends a section", "needs_clarification")
    following = p.getnext()
    ending = None
    while following is not None:
        if following.tag == qn("w:sectPr"):
            ending = following
            break
        ending = following.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if ending is not None: break
        following = following.getnext()
    if ending is None:
        raise FormatError("No containing section found", "unsupported")
    boundary = copy.deepcopy(ending)
    pr.append(boundary)
    element(ending, "type").set(qn("w:val"), "nextPage")


def read(index, ref, cap):
    p = index.elements[ref["id"]]
    node = p.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
    following = p.getnext()
    while following is not None:
        ending = following if following.tag==qn("w:sectPr") else following.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if ending is not None:
            kind=ending.find(qn("w:type"))
            return {"present":node is not None,"type":kind.get(qn("w:val")) if kind is not None else "nextPage"}
        following=following.getnext()
    return {"present":node is not None,"type":None}


def verify(actual, params, cap):
    return actual == {"present": True, "type": "nextPage"}


def projection(index, refs):
    body = copy.deepcopy(index.document.element)
    selected = {index.elements[r["id"]] for r in refs}
    for original, clone in zip(index.document.element.body.iter(qn("w:p")), body.body.iter(qn("w:p"))):
        if original in selected:
            pr = clone.find(qn("w:pPr"))
            if pr is not None:
                node = pr.find(qn("w:sectPr"))
                if node is not None: pr.remove(node)
                if not len(pr) and not pr.attrib: clone.remove(pr)
            following=clone.getnext()
            while following is not None:
                ending=following if following.tag==qn("w:sectPr") else following.find("./"+qn("w:pPr")+"/"+qn("w:sectPr"))
                if ending is not None:
                    kind=ending.find(qn("w:type"))
                    if kind is not None: ending.remove(kind)
                    break
                following=following.getnext()
    return (xml_value(body), {str(p.partname): p.blob for p in index.document.part.package.parts
                              if str(p.partname) != "/word/document.xml"})
