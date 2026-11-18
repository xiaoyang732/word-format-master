"""Resolve body boundaries from the common document semantics, never section ordinals."""

from __future__ import annotations

import copy
from typing import Any

from .contracts import FormatError
from .inspect import DocumentIndex
from .namespaces import qn
from .properties import element


def order_scoped_pagination(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Build the TOC before determining the body boundary in a combined plan."""
    def scoped(item: dict[str, Any]) -> bool:
        params = item.get("params")
        return bool(
            item.get("action") == "page_number.configure"
            and isinstance(params, dict)
            and params.get("enabled")
            and params.get("scope") == "body"
            and "start_section" not in params
        )

    if not all(isinstance(item, dict) for item in operations):
        return operations
    if not any(item.get("action") == "toc.configure" for item in operations):
        return operations
    ordinary = [item for item in operations if not scoped(item)]
    last_toc = next(
        i
        for i in range(len(ordinary) - 1, -1, -1)
        if ordinary[i].get("action") == "toc.configure"
    )
    return (
        ordinary[: last_toc + 1]
        + [item for item in operations if scoped(item)]
        + ordinary[last_toc + 1 :]
    )


def body_start(document: Any) -> Any:
    index = DocumentIndex(document)
    headings = [
        index.elements[oid]
        for oid in index.top_ids
        if index.records[oid].get("semantic_role") == "chapter_heading"
        and index.records[oid].get("heading_level") == 1
    ]
    if not headings:
        raise FormatError("Body start is ambiguous; specify start_section", "needs_clarification")
    return headings[0]


def boundary_preparation(document: Any) -> tuple[Any, Any, Any] | None:
    """Return the exact clone insertion authorized by scope=body, if required."""
    first = body_start(document)
    previous = first.getprevious()
    if previous is None:
        return None
    if previous.tag != qn("w:p"):
        raise FormatError("Body boundary needs a preceding paragraph", "needs_clarification")
    if previous.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
        return None
    ending = first
    while ending is not None:
        section = (
            ending
            if ending.tag == qn("w:sectPr")
            else ending.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        )
        if section is not None:
            return previous, section, copy.deepcopy(section)
        ending = ending.getnext()
    raise FormatError("Containing body section not found")


def ensure_body_boundary(document: Any) -> None:
    preparation = boundary_preparation(document)
    if preparation is None:
        return
    from .ooxml import insert_child

    previous, ending, clone = preparation
    insert_child(element(previous, "pPr"), clone)
    element(ending, "type").set(qn("w:val"), "nextPage")
    # A page break at the same boundary would add an empty page in Word.
    if not "".join(node.text or "" for node in previous.iter(qn("w:t"))).strip():
        for br in list(previous.iter(qn("w:br"))):
            if br.get(qn("w:type")) == "page":
                br.getparent().remove(br)


def body_section_index(document: Any) -> int:
    try:
        first = body_start(document)
    except FormatError:
        return 0
    position = list(document.element.body).index(first)
    return sum(
        1
        for node in list(document.element.body)[:position]
        if node.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")) is not None
    )
