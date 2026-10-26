#!/usr/bin/env python3
"""Report thesis length and citation-mapping checks without claiming fact verification."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from cli_output import configure_cli_output
from document_structure import classify_module, heading_level
from docx import Document
from word_format.contracts import sha256
from word_format.properties import outline_level, paragraph_style_name

_BODY_START = re.compile(
    r"^\s*(?:第\s*[一二三四五六七八九十百零〇0-9]+\s*章|第一篇|绪论|引言|前言|"
    r"chapter\s+1\b|introduction\b)",
    re.I,
)
_REFERENCE_ID = re.compile(r"^\s*(?:\[(\d+)\]|(\d+)\s*[.、．])")
_CITATION_GROUP = re.compile(r"[\[【]([^\]】]*\d[^\]】]*)[\]】]")
_HAN = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U0002FA1F]")
_NUM_ID = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId"
_ILVL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl"
_ABSTRACT_NUM_ID = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId"
_ABSTRACT_NUM = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum"
_NUM = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num"
_NUM_FMT = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt"
_LVL_TEXT = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText"
_START = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start"
_LVL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl"
_LVL_OVERRIDE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlOverride"
_START_OVERRIDE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}startOverride"
_VAL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
_NUM_ID_ATTR = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId"
_ABSTRACT_NUM_ID_ATTR = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId"


def _text(paragraph) -> str:
    return paragraph.text.strip()


def _module(document, paragraph) -> str | None:
    return classify_module(
        _text(paragraph),
        paragraph_style_name(document, paragraph._p),
        outline_level(document, paragraph._p),
    )


def _counts(text: str) -> dict[str, int]:
    return {
        "chinese_han_characters": len(_HAN.findall(text)),
        "non_whitespace_characters": sum(not char.isspace() for char in text),
        "latin_words": len(re.findall(r"\b[A-Za-z]+(?:['’-][A-Za-z]+)*\b", text)),
    }


def _citation_ids(text: str) -> set[int]:
    found: set[int] = set()
    for group in _CITATION_GROUP.findall(text):
        if not re.fullmatch(r"\s*\d+(?:\s*[,，;；、]\s*\d+|\s*[-–—]\s*\d+)*\s*", group):
            continue
        numbers = [int(value) for value in re.findall(r"\d+", group)]
        if len(numbers) == 1 and 1500 <= numbers[0] <= 2100:
            # A lone bracketed four-digit year is ambiguous and commonly
            # appears in author-date citations; do not treat it as a reference ID.
            continue
        if len(numbers) == 2 and re.search(r"\d\s*[-–—]\s*\d", group):
            low, high = numbers
            if low <= high and high - low <= 100:
                found.update(range(low, high + 1))
        else:
            found.update(numbers)
    return found


def _numbering_reference_ids(document, paragraphs) -> dict[int, int]:
    """Resolve decimal Word numbering on reference paragraphs to their display IDs."""
    try:
        numbering = document.part.numbering_part.element
    except (KeyError, NotImplementedError):
        return {}
    num_definitions = {node.get(_NUM_ID_ATTR): node for node in numbering.findall(_NUM)}
    abstract_definitions = {
        node.get(_ABSTRACT_NUM_ID_ATTR): node
        for node in numbering.findall(_ABSTRACT_NUM)
    }
    counters: dict[tuple[str, int], int] = {}
    output: dict[int, int] = {}
    for paragraph_index, paragraph in enumerate(paragraphs):
        ppr = paragraph._p.pPr
        num_pr = ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr") if ppr is not None else None
        if num_pr is None:
            style_ppr = paragraph.style._element.pPr
            num_pr = style_ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr") if style_ppr is not None else None
        if num_pr is None:
            continue
        num_node = num_pr.find(_NUM_ID)
        num_id = num_node.get(_VAL) if num_node is not None else None
        if not num_id:
            continue
        ilvl_node = num_pr.find(_ILVL)
        ilvl = int(ilvl_node.get(_VAL, "0")) if ilvl_node is not None else 0
        num = num_definitions.get(num_id)
        abstract_ref = num.find(_ABSTRACT_NUM_ID) if num is not None else None
        abstract = abstract_definitions.get(abstract_ref.get(_VAL)) if abstract_ref is not None else None
        override = next(
            (node for node in num.findall(_LVL_OVERRIDE) if int(node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl", "-1")) == ilvl),
            None,
        ) if num is not None else None
        level = override.find(_LVL) if override is not None else None
        if level is None and abstract is not None:
            level = next(
                (node for node in abstract.findall(_LVL) if int(node.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl", "-1")) == ilvl),
                None,
            )
        fmt = level.find(_NUM_FMT) if level is not None else None
        text = level.find(_LVL_TEXT) if level is not None else None
        if fmt is None or fmt.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") != "decimal" or text is None:
            continue
        marker = text.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "")
        if re.search(r"%[2-9]", marker):
            continue
        key = (num_id, ilvl)
        if key not in counters:
            start = level.find(_START) if level is not None else None
            start_override = override.find(_START_OVERRIDE) if override is not None else None
            raw_start = (
                start_override.get(_VAL)
                if start_override is not None
                else start.get(_VAL)
                if start is not None
                else "1"
            )
            counters[key] = int(raw_start or "1") - 1
        counters[key] += 1
        output[paragraph_index] = counters[key]
    return output


def analyze_thesis(
    path: str | Path,
    *,
    minimum_chinese_characters: int | None = None,
    maximum_chinese_characters: int | None = None,
) -> dict[str, Any]:
    source = Path(path).expanduser().resolve(strict=True)
    document = Document(source)
    paragraphs = document.paragraphs
    modules = [_module(document, paragraph) for paragraph in paragraphs]

    body_start = next(
        (
            i
            for i, paragraph in enumerate(paragraphs)
            if _BODY_START.match(_text(paragraph))
            and heading_level(
                _text(paragraph),
                paragraph_style_name(document, paragraph._p),
                outline_level(document, paragraph._p),
            ) == 1
        ),
        None,
    )
    references_start = next(
        (i for i, module in enumerate(modules) if module == "references"), None
    )

    body_end = references_start
    if body_end is None:
        body_end = next(
            (
                i
                for i, module in enumerate(modules)
                if body_start is not None
                and i > body_start
                and module in {"acknowledgements", "appendix"}
            ),
            len(paragraphs),
        )
    elif body_start is not None:
        body_end = min(
            (
                i
                for i, module in enumerate(modules)
                if i > body_start and module in {"references", "acknowledgements", "appendix"}
            ),
            default=len(paragraphs),
        )
    body_paragraphs = paragraphs[body_start:body_end] if body_start is not None else []
    body_text = "\n".join(_text(paragraph) for paragraph in body_paragraphs if _text(paragraph))
    body_counts = _counts(body_text)

    abstract_reports: list[dict[str, Any]] = []
    for start, module in enumerate(modules):
        if module != "abstract":
            continue
        end = next(
            (
                i
                for i in range(start + 1, len(paragraphs))
                if modules[i] in {"abstract", "keywords", "toc", "chapters", "references"}
            ),
            len(paragraphs),
        )
        content = "\n".join(
            _text(paragraph) for paragraph in paragraphs[start + 1 : end] if _text(paragraph)
        )
        abstract_reports.append(
            {"heading": _text(paragraphs[start]), **_counts(content)}
        )

    reference_entries: list[str] = []
    reference_ids: set[int] = set()
    if references_start is not None:
        numbering_ids = _numbering_reference_ids(
            document,
            [paragraphs[i] for i in range(references_start + 1, len(paragraphs))],
        )
        for i in range(references_start + 1, len(paragraphs)):
            text = _text(paragraphs[i])
            if modules[i] and modules[i] != "references":
                break
            if not text:
                continue
            reference_entries.append(text)
            match = _REFERENCE_ID.match(text)
            if match:
                reference_ids.add(int(match.group(1) or match.group(2)))
            elif i - references_start - 1 in numbering_ids:
                reference_ids.add(numbering_ids[i - references_start - 1])

    citation_text = "\n".join(_text(paragraph) for paragraph in body_paragraphs if _text(paragraph))
    citations = _citation_ids(citation_text)
    missing_reference_ids = sorted(citations - reference_ids)
    unused_reference_ids = sorted(reference_ids - citations)

    if minimum_chinese_characters is None and maximum_chinese_characters is None:
        length_status = "not_requested"
    elif body_start is None:
        length_status = "not_assessable"
    else:
        in_range = (
            (minimum_chinese_characters is None or body_counts["chinese_han_characters"] >= minimum_chinese_characters)
            and (maximum_chinese_characters is None or body_counts["chinese_han_characters"] <= maximum_chinese_characters)
        )
        length_status = "passed" if in_range else "failed"

    if citations and reference_ids:
        citation_status = "failed" if missing_reference_ids else "passed"
    elif citations and not reference_entries:
        citation_status = "failed"
    elif citations or reference_entries:
        citation_status = "not_assessable"
    else:
        citation_status = "not_assessable"

    hard_failures = length_status == "failed" or citation_status == "failed"
    return {
        "schema_version": "1.0",
        "kind": "thesis-content-quality-report",
        "source": {"path": str(source), "sha256": sha256(source)},
        "status": "failed" if hard_failures else "review_required",
        "measurement_scope": {
            "body_start_heading": _text(paragraphs[body_start]) if body_start is not None else None,
            "body_end_heading": _text(paragraphs[references_start]) if references_start is not None else None,
            "body": body_counts,
            "body_count_method": "Chinese Han characters in paragraphs from the first recognized body heading up to the references heading; heading text included.",
            "abstracts": abstract_reports,
        },
        "targets": {
            "minimum_chinese_characters": minimum_chinese_characters,
            "maximum_chinese_characters": maximum_chinese_characters,
        },
        "checks": {
            "body_length": {
                "status": length_status,
                "measured_chinese_han_characters": body_counts["chinese_han_characters"],
            },
            "numeric_citation_mapping": {
                "status": citation_status,
                "cited_reference_ids": sorted(citations),
                "listed_reference_ids": sorted(reference_ids),
                "citations_without_matching_reference": missing_reference_ids,
                "listed_but_uncited_reference_ids": unused_reference_ids,
            },
            "bibliography_source_verification": {
                "status": "manual_review_required" if reference_entries else "no_entries_found",
                "entry_count": len(reference_entries),
                "entries": [
                    {"text": entry, "source_verification": "not_verified"}
                    for entry in reference_entries
                ],
                "note": "This local check cannot establish that author, title, edition, publisher, year, pages, or cited claims match authoritative sources.",
            },
            "academic_claims": {
                "status": "not_assessed",
                "note": "The report does not verify historical, scholarly, experimental, or other factual claims.",
            },
        },
    }


def main(argv: list[str] | None = None) -> int:
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Thesis DOCX to analyze")
    parser.add_argument("--min-chinese-characters", type=int)
    parser.add_argument("--max-chinese-characters", type=int)
    parser.add_argument("--output", type=Path, help="JSON report path (default: stdout)")
    args = parser.parse_args(argv)
    if args.min_chinese_characters is not None and args.min_chinese_characters < 0:
        parser.error("--min-chinese-characters must be non-negative")
    if args.max_chinese_characters is not None and args.max_chinese_characters < 0:
        parser.error("--max-chinese-characters must be non-negative")
    if (
        args.min_chinese_characters is not None
        and args.max_chinese_characters is not None
        and args.min_chinese_characters > args.max_chinese_characters
    ):
        parser.error("minimum Chinese character target exceeds maximum")
    try:
        report = analyze_thesis(
            args.input,
            minimum_chinese_characters=args.min_chinese_characters,
            maximum_chinese_characters=args.max_chinese_characters,
        )
    except (OSError, ValueError, KeyError) as exc:
        print(str(exc), file=sys.stderr)
        return 2
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        target = args.output.expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    length_requested = (
        args.min_chinese_characters is not None or args.max_chinese_characters is not None
    )
    return (
        1
        if report["status"] == "failed"
        or (length_requested and report["checks"]["body_length"]["status"] != "passed")
        else 0
    )


if __name__ == "__main__":
    raise SystemExit(main())
