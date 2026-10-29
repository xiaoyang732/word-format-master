"""Native numbering driven by explicit template levels, without appearance defaults."""

from __future__ import annotations

import re

from docx.oxml import OxmlElement

from .contracts import FormatError, json_hash, number, require_keys, text_value
from .namespaces import qn
from .properties import element, ppr_sources, strip_text_prefix


def validate_scheme(params):
    require_keys(params, {"levels"}, {"levels"}, "numbering scheme")
    if not isinstance(params["levels"], list) or not 1 <= len(params["levels"]) <= 9:
        raise FormatError("Numbering requires 1..9 explicit levels")
    levels = []
    for i, level in enumerate(params["levels"]):
        require_keys(
            level,
            {"format", "text", "start", "suffix", "restart", "legal"},
            {"format", "text", "start", "suffix"},
            "numbering level",
        )
        text_value(level["format"], "number format")
        if level["format"] not in {
            "decimal",
            "decimalZero",
            "upperRoman",
            "lowerRoman",
            "upperLetter",
            "lowerLetter",
            "ordinal",
            "cardinalText",
            "ordinalText",
            "chineseCounting",
            "chineseCountingThousand",
            "chineseLegalSimplified",
            "ideographDigital",
            "ideographTraditional",
            "taiwaneseCounting",
            "taiwaneseCountingThousand",
            "none",
        }:
            raise FormatError("Unsupported OOXML chapter number format", "unsupported")
        text_value(level["text"], "number text")
        number(level["start"], 0, 32767, "number start", True)
        if level["suffix"] not in {"space", "tab", "nothing"}:
            raise FormatError("Number suffix must be space, tab or nothing")
        if any(int(n) > i + 1 or int(n) == 0 for n in re.findall(r"%(\d+)", level["text"])):
            raise FormatError("Number text references a level not available here")
        if "restart" in level:
            number(level["restart"], 0, i, "number restart", True)
        if "legal" in level and not isinstance(level["legal"], bool):
            raise FormatError("legal must be boolean")
        levels.append(dict(level))
    return {"levels": levels}


def effective_numbering(document, paragraph):
    num_id = None
    ilvl = 0
    for pr in ppr_sources(document, paragraph):
        np = pr.find(qn("w:numPr")) if pr is not None else None
        if np is None:
            continue
        ni, il = np.find(qn("w:numId")), np.find(qn("w:ilvl"))
        if il is not None:
            ilvl = int(il.get(qn("w:val")))
        if ni is not None:
            num_id = ni.get(qn("w:val"))
            break
    if num_id in {None, "0"}:
        return None
    numbering = document.part.numbering_part.element
    num = next((n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == num_id), None)
    aid = num.find(qn("w:abstractNumId")) if num is not None else None
    abstract = next(
        (
            n
            for n in numbering.findall(qn("w:abstractNum"))
            if aid is not None and n.get(qn("w:abstractNumId")) == aid.get(qn("w:val"))
        ),
        None,
    )
    return num_id, ilvl, abstract, num


def level_spec(level):
    result = {}
    for tag, key in (
        ("numFmt", "format"),
        ("lvlText", "text"),
        ("start", "start"),
        ("suff", "suffix"),
        ("lvlRestart", "restart"),
    ):
        node = level.find(qn("w:" + tag))
        if node is not None:
            raw = node.get(qn("w:val"))
            result[key] = int(raw) if key in {"start", "restart"} else raw
    # OOXML defaults, not a thesis preset.
    result.setdefault("start", 1)
    result.setdefault("suffix", "tab")
    if level.find(qn("w:isLgl")) is not None:
        result["legal"] = True
    return result


def extract_heading_scheme(index):
    schemes = {}
    for oid in index.top_ids:
        r = index.records[oid]
        if r.get("semantic_role") != "chapter_heading":
            continue
        found = effective_numbering(index.document, index.elements[oid])
        if found is None:
            continue
        _, _, abstract, num = found
        if abstract is None or num.findall(qn("w:lvlOverride")):
            raise FormatError(
                "Heading numbering overrides require explicit interpretation", "needs_clarification"
            )
        used = max(
            (
                x.get("heading_level") or 1
                for x in index.records.values()
                if x.get("semantic_role") == "chapter_heading"
            ),
            default=1,
        )
        levels = sorted(abstract.findall(qn("w:lvl")), key=lambda n: int(n.get(qn("w:ilvl"))))
        if len(levels) < used:
            raise FormatError(
                "Sample heading numbering does not define all observed levels",
                "needs_clarification",
            )
        scheme = validate_scheme({"levels": [level_spec(n) for n in levels[:used]]})
        schemes[json_hash(scheme)] = scheme
    if len(schemes) > 1:
        raise FormatError("Sample chapters use different numbering schemes", "needs_clarification")
    return next(iter(schemes.values()), None)


def _definition(document, params):
    root = document.part.numbering_part.element
    name = "WFM:" + json_hash(params)
    for abstract in root.findall(qn("w:abstractNum")):
        label = abstract.find(qn("w:name"))
        if label is not None and label.get(qn("w:val")) == name:
            aid = abstract.get(qn("w:abstractNumId"))
            for num in root.findall(qn("w:num")):
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and ref.get(qn("w:val")) == aid:
                    return int(num.get(qn("w:numId")))
    aid = (
        max(
            (int(n.get(qn("w:abstractNumId"))) for n in root.findall(qn("w:abstractNum"))),
            default=-1,
        )
        + 1
    )
    nid = max((int(n.get(qn("w:numId"))) for n in root.findall(qn("w:num"))), default=0) + 1
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(aid))
    element(abstract, "multiLevelType").set(qn("w:val"), "multilevel")
    element(abstract, "name").set(qn("w:val"), name)
    for i, config in enumerate(params["levels"]):
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), str(i))
        for tag, key in (
            ("start", "start"),
            ("numFmt", "format"),
            ("suff", "suffix"),
            ("lvlText", "text"),
            ("lvlRestart", "restart"),
        ):
            if key in config:
                element(level, tag).set(qn("w:val"), str(config[key]))
        if config.get("legal"):
            element(level, "isLgl")
        abstract.append(level)
    root.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(nid))
    element(num, "abstractNumId").set(qn("w:val"), str(aid))
    root.append(num)
    return nid


def apply_heading_scheme(index, ref, params):
    from .structures import _set_paragraph_numbering

    records = [
        index.records[oid]
        for oid in index.top_ids
        if index.records[oid].get("semantic_role") == "chapter_heading"
    ]
    if not records:
        return []
    nid = _definition(index.document, params)
    for record in records:
        level = record["heading_level"]
        if not level or level > len(params["levels"]):
            raise FormatError("Template has no numbering definition for this heading level")
        paragraph = index.paragraph(record["id"])
        paragraph.style = index.document.styles[f"Heading {level}"]
        # Strip only an established heading prefix. Bare digits are content.
        prefix = re.match(
            r"^\s*(?:第[一二三四五六七八九十百千零〇0-9]+章\s*|\d+(?:\.\d+){0,8}[、.．]?\s+)",
            paragraph.text,
        )
        if prefix:
            strip_text_prefix(paragraph._p, prefix.end())
        _set_paragraph_numbering(paragraph, nid, level - 1)
    return [f"Applied template numbering to {len(records)} chapter headings"]


def read_heading_scheme(index, ref, cap):
    result = []
    for oid in index.top_ids:
        r = index.records[oid]
        if r.get("semantic_role") != "chapter_heading":
            continue
        found = effective_numbering(index.document, index.elements[oid])
        config = None
        if found and found[2] is not None:
            node = next(
                (n for n in found[2].findall(qn("w:lvl")) if int(n.get(qn("w:ilvl"))) == found[1]),
                None,
            )
            config = level_spec(node) if node is not None else None
        result.append(
            {
                "level": r["heading_level"],
                "config": config,
                "ilvl": found[1] if found else None,
                "outline": r.get("outline_level"),
            }
        )
    return result


def verify_heading_scheme(actual, params, cap):
    return all(
        r["level"]
        and r["level"] <= len(params["levels"])
        and r["ilvl"] == r["level"] - 1
        and r["outline"] == r["level"] - 1
        and r["config"] == params["levels"][r["level"] - 1]
        for r in actual
    )
