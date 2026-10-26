"""Shared field classification for header content and page-number ownership."""

import re

from .namespaces import qn


def field_codes(root):
    if root is None:
        return []
    return [" ".join((node.text or "").split()) for node in root.iter(qn("w:instrText"))] + [
        " ".join(node.get(qn("w:instr"), "").split()) for node in root.iter(qn("w:fldSimple"))
    ]


def page_number_paragraphs(root):
    """Identify complete page-only paragraphs, including imported simple fields."""
    if root is None:
        return []
    result = []
    for paragraph in root.iter(qn("w:p")):
        codes = field_codes(paragraph)
        if not any(re.match(r"^PAGE(?:\s|$)", code, re.I) for code in codes):
            continue
        if any(not re.match(r"^(?:PAGE|NUMPAGES)(?:\s|$)", code, re.I) for code in codes):
            continue
        style = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:pStyle"))
        managed = style is not None and style.get(qn("w:val")) == "WFMPageNumber"
        visible = "".join(node.text or "" for node in paragraph.iter(qn("w:t"))).strip()
        if (
            managed
            or not visible
            or re.fullmatch(
                r"(?:Page\s*)?(?:第\s*)?\d+(?:\s*(?:of|/|页，共)\s*\d+\s*页?)?", visible, re.I
            )
        ):
            result.append(paragraph)
    return result
