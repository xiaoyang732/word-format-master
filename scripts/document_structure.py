#!/usr/bin/env python3
"""Shared document-module detection and ordering rules.

This module deliberately contains no DOCX writes.  XML analysis, DOCX style
application, and verification all use the same marker vocabulary so a module
cannot be extracted under one name and validated under another.
"""

from __future__ import annotations

import re
from typing import Any, Iterable


MODULE_ORDER = (
    "cover",
    "declaration",
    "abstract",
    "keywords",
    "toc",
    "symbols",
    "chapters",
    "references",
    "acknowledgements",
    "appendix",
)

MODULE_MARKERS = {
    "declaration": re.compile(
        r"^\s*(?:(?:诚信|原创性|独创性|学术诚信|郑重)\s*声明|(?:学位论文)?(?:版权使用|使用)?授权书|声明与授权|授权声明)\s*$",
        re.I,
    ),
    "abstract": re.compile(r"^\s*(?:摘\s*要|中文摘要|英文摘要|abstracts?)(?:\s*[:：].*)?$", re.I),
    "keywords": re.compile(r"^\s*(?:关\s*键\s*词|key\s*words?)(?:\s*[:：].*)?$", re.I),
    "toc": re.compile(r"^\s*(?:目\s*录|table\s+of\s+contents|contents)\s*$", re.I),
    "references": re.compile(r"^\s*(?:参考文献|参考资料|references|bibliography|works\s+cited)\s*$", re.I),
    "acknowledgements": re.compile(r"^\s*(?:致\s*谢|acknowledg(?:e)?ments?)\s*$", re.I),
    "appendix": re.compile(r"^\s*(?:附\s*录|appendix)(?:\s+[A-Z0-9一二三四五六七八九十]+)?\s*$", re.I),
    "symbols": re.compile(r"^\s*(?:符号说明|符号表|术语表|symbols?|glossary)\s*$", re.I),
}

_CHAPTER_RE = re.compile(r"^\s*(?:第\s*[一二三四五六七八九十0-9]+\s*章|\d+[、.\s])")
_LEVEL2_RE = re.compile(r"^\s*\d+\.\d+(?:[、.\s])")
_LEVEL3_RE = re.compile(r"^\s*\d+\.\d+\.\d+(?:[、.\s])")


def normalized_style_name(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def heading_level(text: str, style_name: str = "", outline_level: int | None = None) -> int | None:
    if outline_level is not None:
        try:
            return int(outline_level) + 1
        except (TypeError, ValueError):
            pass
    normalized = normalized_style_name(style_name)
    match = re.search(r"(?:heading|标题)([1-9])$", normalized)
    if match:
        return int(match.group(1))
    if _LEVEL3_RE.match(text):
        return 3
    if _LEVEL2_RE.match(text):
        return 2
    if _CHAPTER_RE.match(text):
        return 1
    return None


def classify_module(text: str, style_name: str = "", outline_level: int | None = None) -> str | None:
    normalized = normalized_style_name(style_name)
    if normalized in {"wfmtocfield", "wfmtocbreak"}:
        return None
    if normalized in {"wfmtocheading"}:
        return "toc"
    for kind, pattern in MODULE_MARKERS.items():
        if pattern.match(text):
            return kind
    if heading_level(text, style_name, outline_level) == 1:
        return "chapters"
    return None


def is_heading_candidate(text: str, style_name: str = "") -> bool:
    normalized = normalized_style_name(style_name)
    if normalized.startswith("heading") or normalized.startswith("标题") or normalized in {"title", "subtitle"}:
        return True
    if heading_level(text, style_name) is not None:
        return True
    return False


def classify_role(text: str, style_name: str = "", *, first: bool = False) -> str:
    normalized = normalized_style_name(style_name)
    if normalized in {"title", "subtitle"}:
        return "title"
    if first:
        return "heading"
    if "keyword" in normalized or "关键词" in normalized or re.search(r"^\s*(?:关\s*键\s*词|key\s*words?)\s*[:：]", text, re.I):
        return "keywords"
    if normalized.startswith("heading") or normalized.startswith("标题") or is_heading_candidate(text, style_name):
        return "heading"
    return "body"


def detect_module_spans(paragraphs: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return ordered module spans for text paragraphs.

    Input items must provide ``index`` and ``text`` and may provide
    ``style_name``/``outline_level``.  Empty paragraphs should be omitted by
    callers because they do not provide reliable semantic anchors.
    """
    items = [item for item in paragraphs if str(item.get("text") or "").strip()]
    if not items:
        return []
    starts: list[tuple[int, str]] = []
    for offset, item in enumerate(items):
        kind = classify_module(
            str(item.get("text") or "").strip(),
            str(item.get("style_name") or ""),
            item.get("outline_level"),
        )
        if kind is None:
            continue
        if kind == "chapters":
            if starts and starts[-1][1] == "chapters":
                continue
        starts.append((offset, kind))
    if not starts:
        starts = [(0, "cover")]
    elif starts[0][0] > 0:
        starts.insert(0, (0, "cover"))
    spans: list[dict[str, Any]] = []
    for pos, (start, kind) in enumerate(starts):
        end = starts[pos + 1][0] - 1 if pos + 1 < len(starts) else len(items) - 1
        first = items[start]
        spans.append({
            "id": kind,
            "kind": kind,
            "repeatable": kind == "chapters",
            "paragraph_start": first.get("index", start),
            "paragraph_end": items[end].get("index", end),
            "heading_text": str(first.get("text") or "").strip(),
            "paragraph_count": end - start + 1,
            "role_counts": {
                role: sum(
                    1
                    for item in items[start : end + 1]
                    if classify_role(
                        str(item.get("text") or "").strip(),
                        str(item.get("style_name") or ""),
                        first=item is first,
                    ) == role
                )
                for role in ("title", "heading", "body", "keywords")
            },
        })
    return spans


MODULE_STAGE_RANK = {
    "cover": 0,
    "declaration": 1,
    "abstract": 2,
    "keywords": 2,
    "toc": 3,
    "symbols": 4,
    "chapters": 5,
    "references": 6,
    "acknowledgements": 7,
    "appendix": 8,
}


def validate_module_order(spans: Iterable[dict[str, Any]], *, strict: bool = True) -> list[str]:
    if not strict:
        return []
    errors: list[str] = []
    previous_stage = -1
    for item in spans:
        kind = str(item.get("kind") or item.get("id") or "")
        current_stage = MODULE_STAGE_RANK.get(kind)
        if current_stage is None:
            errors.append(f"unknown document module: {kind}")
            continue
        if current_stage < previous_stage:
            errors.append(f"document modules are out of order at {kind}")
        previous_stage = max(previous_stage, current_stage)
    return errors
