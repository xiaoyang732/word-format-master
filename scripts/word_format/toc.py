"""Shared heading discovery and TOC result validation, independent of fonts."""

import re

from document_structure import MODULE_MARKERS, heading_level

from .namespaces import qn
from .properties import outline_level, paragraph_style_name


def imported_fields(document):
    """Only complete, standalone imported TOC fields can be safely replaced.

    A field spanning paragraphs or sharing text with prose needs an explicit
    structural decision. Never remove surrounding text as if it were a cache.
    """
    from .contracts import FormatError
    from .fields import field_codes

    result = []
    for p in document.paragraphs:
        if paragraph_style_name(document, p._p).startswith("WFM TOC"):
            continue
        if not any(re.match(r"^TOC(?:\s|$)", c, re.I) for c in field_codes(p._p)):
            continue
        depth = 0
        outside = False
        codes = []
        for node in p._p.iter():
            if node.tag == qn("w:fldChar"):
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    depth += 1
                    codes.append("")
                elif kind == "end":
                    if depth == 1 and codes and not re.match(r"^\s*TOC(?:\s|$)", codes[-1], re.I):
                        outside = True
                    if codes:
                        codes.pop()
                    depth -= 1
                    if depth < 0:
                        outside = True
            elif node.tag == qn("w:instrText") and codes:
                codes[-1] += node.text or ""
            elif node.tag == qn("w:fldSimple"):
                if not depth and not re.match(
                    r"^\s*TOC(?:\s|$)", node.get(qn("w:instr"), ""), re.I
                ):
                    outside = True
            elif node.tag in {qn("w:t"), qn("w:drawing"), qn("w:object"), qn("w:pict")}:
                simple = any(
                    a.tag == qn("w:fldSimple")
                    and re.match(r"^\s*TOC(?:\s|$)", a.get(qn("w:instr"), ""), re.I)
                    for a in node.iterancestors()
                )
                if not depth and not simple:
                    outside = True
        if depth or outside:
            raise FormatError(
                "Imported TOC spans paragraphs or shares unrelated content; map its structure explicitly",
                "needs_clarification",
            )
        if p._p.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            raise FormatError(
                "Imported TOC carries a section break; preserve its boundary explicitly",
                "needs_clarification",
            )
        result.append(p._p)
    return result


def headings(document, max_level=9):
    result = []
    for p in document.paragraphs:
        text = p.text.strip()
        style = paragraph_style_name(document, p._p)
        outline = outline_level(document, p._p)
        if (
            not text
            or outline == 9
            or style.lower().replace(" ", "").startswith(("toc", "wfmtoc"))
            or MODULE_MARKERS["toc"].fullmatch(text)
        ):
            continue
        level = heading_level(text, style, outline)
        if level and level <= max_level:
            result.append((p, level))
    return result


def result_status(document, max_level):
    expected = [p.text.strip() for p, _ in headings(document, max_level)]
    fields = [
        p for p in document.paragraphs if paragraph_style_name(document, p._p) == "WFM TOC Field"
    ]
    text = "\n".join("".join(n.text or "" for n in p._p.iter(qn("w:t"))) for p in fields)
    error = bool(
        re.search(
            r"错误[！!]?\s*[：:]?\s*未找到目录项|Error!\s*No table of contents entries found"
            r"|目录将在\s*Word\s*中更新",
            text,
            re.I,
        )
    )

    def compact(s):
        return re.sub(r"\s+", "", s)

    remaining = compact(text)
    missing = []
    for label in expected:
        label = compact(label)
        pos = remaining.find(label)
        if pos < 0:
            missing.append(label)
        else:
            remaining = remaining[pos + len(label) :]
    return {
        "valid": bool(expected and text.strip()) and not error and not missing,
        "heading_count": len(expected),
        "missing": missing,
        "error": error,
        "text": text,
    }
