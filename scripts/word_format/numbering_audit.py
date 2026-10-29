"""Read-only guard against typed structural numbers in generated documents."""

import re

from .inspect import DocumentIndex
from .namespaces import qn
from .numbering import effective_numbering, heading_prefix

_LIST = re.compile(
    r"^\s*(?:\d+[、．](?!\d)|\d+\.\s+|[（(]\d+[）)]"
    r"|[一二三四五六七八九十百]+[、．]|[（(][一二三四五六七八九十百]+[）)])"
)
_CAPTION = re.compile(r"^\s*(?:图|表|Figure\s*|Table\s*)\s*\d", re.I)
_CITATION = re.compile(r"[\[［]\s*\d+(?:\s*[,，、;；\-–]\s*\d+)*\s*[\]］]")


def audit_numbering(document):
    """Inspect actual XML, excluding field results rather than trusting their display text.

    Semantic labels (not every digit in prose/data) establish numbering intent.
    Ambiguous equations/identifiers still require the authoring review described
    in thesis-generation.md.
    """
    index = DocumentIndex(document)
    records = {
        index.elements[oid]: r for oid, r in index.records.items() if r["kind"] == "paragraph"
    }
    issues = []
    checked = 0
    for part in document.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            continue
        # Complex fields can span paragraphs (notably a generated TOC).
        depth = 0
        for position, paragraph in enumerate(root.iter(qn("w:p")), 1):
            checked += 1
            literal = []
            for node in paragraph.iter():
                if node.tag == qn("w:fldChar"):
                    kind = node.get(qn("w:fldCharType"))
                    if kind == "begin":
                        depth += 1
                    elif kind == "end":
                        depth = max(0, depth - 1)
                elif node.tag == qn("w:t") and not depth:
                    if not any(
                        a.tag in {qn("w:fldSimple"), qn("w:del")} for a in node.iterancestors()
                    ):
                        literal.append(node.text or "")
            text = "".join(literal)
            record = records.get(paragraph, {})
            kind = None
            if record.get("semantic_role") == "chapter_heading" and heading_prefix(text):
                kind = "heading"
            elif _CAPTION.match(text):
                kind = "caption"
            elif _CITATION.search(text):
                kind = "reference_or_citation"
            elif _LIST.match(text):
                kind = "list"
            elif re.match(r"^\s*(?:附录|Appendix)\s*[A-ZＡ-Ｚ0-9]", text, re.I):
                kind = "appendix"
            elif any(
                n.tag.endswith("}oMath") or n.tag.endswith("}oMathPara") for n in paragraph.iter()
            ) and re.search(r"[（(]\s*\d+(?:[.．\-–]\d+)*\s*[）)]\s*$", text):
                kind = "equation"
            elif (
                "header" in str(part.partname) or "footer" in str(part.partname)
            ) and re.fullmatch(r"\s*(?:第\s*)?\d+\s*(?:页)?\s*", text):
                kind = "page"
            if kind:
                issues.append(
                    {
                        "part": str(part.partname),
                        "paragraph": position,
                        "kind": kind,
                        "text": text[:160],
                        "reason": "Typed number; use Word numbering or a dynamic field",
                    }
                )
            elif record.get("semantic_role") == "chapter_heading":
                found = effective_numbering(document, paragraph)
                if found and found[2] is None:
                    issues.append(
                        {
                            "part": str(part.partname),
                            "paragraph": position,
                            "kind": "heading",
                            "text": text[:160],
                            "reason": "Numbering definition is missing",
                        }
                    )
    return {
        "status": "failed" if issues else "passed",
        "issues": issues,
        "checked_paragraphs": checked,
    }
