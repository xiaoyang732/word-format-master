#!/usr/bin/env python3
"""Build and apply a deterministic AI citation-placement exchange format."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from docx.document import Document as DocumentType
from docx.text.paragraph import Paragraph


REFERENCE_HEADING_RE = re.compile(
    r"^\s*(?:references|bibliography|works cited|参考文献|参考资料)\s*$", re.I
)
TERMINAL_PUNCTUATION = set(".。!！?？;；:：,，、")


def normalized_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def paragraph_sha256(text: str) -> str:
    return hashlib.sha256(normalized_text(text).encode("utf-8")).hexdigest()


def _is_heading(paragraph: Paragraph) -> bool:
    style = paragraph.style.name.lower() if paragraph.style else ""
    return style.startswith("heading")


def extract_citation_data(document: DocumentType) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Return body paragraph anchors and references using stable top-level indexes."""
    from word_format.inspect import DocumentIndex

    index_data = DocumentIndex(document)
    records = {index_data.elements[oid]: index_data.records[oid] for oid in index_data.top_ids}
    reference_nodes = index_data.reference_nodes()
    paragraphs: list[dict[str, Any]] = []
    references: list[dict[str, Any]] = []
    reference_index = 0
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        record = records.get(paragraph._p, {})
        if paragraph._p in reference_nodes:
            if text:
                reference_index += 1
                label_match = re.match(r"^\s*(\[\s*\d+\s*\]|\d+[.)])", text)
                label = label_match.group(1) if label_match else f"[{reference_index}]"
                references.append({
                    "id": f"ref-{reference_index}",
                    "label": label,
                    "paragraph_index": index,
                    "text": text,
                    "paragraph_sha256": paragraph_sha256(text),
                })
            continue
        if text and record.get('role') == 'body' and record.get('module') in {None, 'chapters'}:
            paragraphs.append({
                "paragraph_index": index,
                "text": text,
                "paragraph_sha256": paragraph_sha256(text),
            })
    return paragraphs, references


def build_citation_task(document: DocumentType, *, filename: str = "document.docx", source_sha256: str | None = None) -> dict[str, Any]:
    paragraphs, references = extract_citation_data(document)
    return {
        "schema_version": "1.0",
        "kind": "word-citation-placement-task",
        "source": {
            "filename": Path(filename).name,
            "sha256": source_sha256,
            "paragraph_index_scope": "top-level document paragraphs",
        },
        "instructions": {
            "goal": "Identify claims in body paragraphs that require a cited reference.",
            "insert_options": ["before-terminal-punctuation", "after-terminal-punctuation", "end", "start"],
            "spacing_options": ["none", "space-before", "space-after", "spaces-around"],
            "return_only": "JSON with citations.placements; do not rewrite paragraph text or references.",
            "fingerprint_rule": "Copy paragraph_index and paragraph_sha256 exactly from candidates.",
        },
        "paragraphs": paragraphs,
        "references": references,
        "citations": {"placements": []},
    }


def _citation_offset(text: str, placement: dict[str, Any]) -> int:
    position = str(placement.get("insert") or "before-terminal-punctuation")
    if position == "start":
        return 0
    if position == "char-offset":
        offset = int(placement.get("char_offset", -1))
        if offset < 0 or offset > len(text):
            raise ValueError("citation char_offset is outside the paragraph")
        return offset
    end = len(text.rstrip())
    if position == "end":
        return len(text)
    if position in {"before-terminal-punctuation", "after-terminal-punctuation"}:
        index = end - 1
        while index >= 0 and text[index] in TERMINAL_PUNCTUATION:
            index -= 1
        punctuation_start = index + 1
        return punctuation_start if position.startswith("before") else end
    raise ValueError(f"unsupported citation insert position: {position}")


def apply_citations(document: DocumentType, citations: dict[str, Any] | None) -> list[str]:
    from word_format.citations import insert_citations

    return insert_citations(document, citations)
