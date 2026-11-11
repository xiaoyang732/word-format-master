"""Keyword separator detection, requirement parsing, and run-preserving normalization."""

from __future__ import annotations

import re

from document_structure import keyword_label_ranges

from .contracts import FormatError, require_keys
from .namespaces import qn
from .properties import text_runs

KEYWORD_SEPARATORS = (
    "，",
    "；",
    "、",
    "; ",
    ";",
    ", ",
    ",",
    " ",
    "  ",
    "\u3000",
)

_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
_ASCII_WORD_RE = re.compile(r"[A-Za-z0-9_+-]+")
_TRAILING_TRIM_CHARS = " \t\xa0\u3000。.；;，,、"


def _normalize_nbsp(raw: str) -> str:
    if raw in {";\xa0", ",\xa0"}:
        return raw[0] + " "
    if set(raw) <= {" ", "\t", "\xa0"} and len(raw) >= 2:
        return " " * len(raw)
    if raw == "\xa0":
        return " "
    return raw


def validate_keyword_separator(params):
    require_keys(params, {"value"}, {"value"})
    value = params.get("value")
    if not isinstance(value, str) or value not in KEYWORD_SEPARATORS:
        raise FormatError(
            f"keyword separator must be one of {list(KEYWORD_SEPARATORS)!r}"
        )
    return {"value": value}


def keyword_separator_spans(text: str) -> list[tuple[int, int, str]]:
    """Return (start, end, separator) spans inside the value segment of a keyword line."""
    ranges = keyword_label_ranges(text)
    if not ranges:
        return []
    val_start, val_end = ranges["value"]
    content_end = val_end
    while content_end > val_start and text[content_end - 1] in _TRAILING_TRIM_CHARS:
        content_end -= 1
    if content_end <= val_start:
        return []
    sub = text[val_start:content_end]

    punct_matches = [
        m
        for m in re.finditer(r"[ \t\xa0\u3000]*[，；、;,][ \t\xa0\u3000]*", sub)
        if 0 < m.start() and m.end() < len(sub)
    ]
    if punct_matches:
        return [
            (val_start + m.start(), val_start + m.end(), _normalize_nbsp(m.group(0)))
            for m in punct_matches
        ]

    wide_matches = [
        m
        for m in re.finditer(r"(?:\u3000+|[ \t\xa0]{2,})", sub)
        if 0 < m.start() and m.end() < len(sub)
    ]
    if wide_matches:
        return [
            (val_start + m.start(), val_start + m.end(), _normalize_nbsp(m.group(0)))
            for m in wide_matches
        ]

    space_matches = [
        m
        for m in re.finditer(r"[ \t\xa0]+", sub)
        if 0 < m.start() and m.end() < len(sub)
    ]
    if space_matches:
        parts = re.split(r"[ \t\xa0]+", sub)
        has_cjk = any(_CJK_RE.search(part) for part in parts)
        adjacent_ascii = any(
            _ASCII_WORD_RE.fullmatch(parts[i]) and _ASCII_WORD_RE.fullmatch(parts[i + 1])
            for i in range(len(parts) - 1)
        )
        if has_cjk and not adjacent_ascii:
            return [
                (val_start + m.start(), val_start + m.end(), _normalize_nbsp(m.group(0)))
                for m in space_matches
            ]

    return []


def keyword_separators(text: str) -> list[str]:
    """Return sorted distinct separators observed in a keyword paragraph."""
    return sorted({sep for _, _, sep in keyword_separator_spans(text)})


def read_keyword_separator(index, ref, cap=None):
    p = index.elements[ref["id"]]
    text = "".join(t.text or "" for r in text_runs(p) for t in r.findall(qn("w:t")))
    spans = keyword_separator_spans(text)
    if ref["kind"] == "text_range":
        lo = ref.get("start", 0)
        hi = ref.get("end", 10**12)
        spans = [(s, e, sep) for s, e, sep in spans if s >= lo and e <= hi]
    return sorted({sep for _, _, sep in spans})


def set_keyword_separator(index, ref, params):
    p = index.elements[ref["id"]]
    runs = list(text_runs(p))
    text = "".join(t.text or "" for r in runs for t in r.findall(qn("w:t")))
    spans = keyword_separator_spans(text)
    if not spans:
        return []
    if any(child.tag not in {qn("w:rPr"), qn("w:t")} for r in runs for child in r):
        raise FormatError(
            "Keyword paragraphs containing fields, drawings, tabs or breaks require a supported explicit operation",
            "unsupported",
        )
    new_sep = params["value"]
    pieces = []
    cursor = 0
    for start, end, _ in spans:
        pieces.append(text[cursor:start])
        pieces.append(new_sep)
        cursor = end
    pieces.append(text[cursor:])
    new_text = "".join(pieces)
    new_spans = keyword_separator_spans(new_text)
    if len(new_spans) != len(spans) or any(raw != new_sep for _, _, raw in new_spans):
        raise FormatError(
            f"Keyword separator {new_sep!r} is ambiguous for the keyword terms in {ref['id']}",
            "needs_clarification",
        )
    start_map = {start: new_sep for start, _, _ in spans}
    skip_offsets = {offset for start, end, _ in spans for offset in range(start + 1, end)}
    offset = 0
    for run in runs:
        for t in run.findall(qn("w:t")):
            original = t.text or ""
            if not original:
                continue
            out = []
            for ch in original:
                if offset in start_map:
                    out.append(start_map[offset])
                elif offset not in skip_offsets:
                    out.append(ch)
                offset += 1
            updated = "".join(out)
            if updated != original:
                t.text = updated
                if updated[:1].isspace() or updated[-1:].isspace() or "  " in updated:
                    t.set(qn("xml:space"), "preserve")
    if ref["id"] in index.records:
        index.records[ref["id"]]["text"] = new_text
        ranges = keyword_label_ranges(new_text)
        if ranges:
            index.records[ref["id"]]["segments"] = {
                k: list(v) for k, v in ranges.items()
            }
    return [f"Normalized keyword separators in {ref['id']} to {new_sep!r}"]


_EXPLICIT_KEYWORD_BETWEEN_RE = (
    r"(?:(?:每|各)?(?:关键)?词(?:与(?:关键)?词)?(?:之)?间|词间|between\s+keywords?)"
)
_GENERAL_BETWEEN_RE = (
    r"(?:(?:每|各)?(?:关键)?词(?:与(?:关键)?词)?(?:之)?间|词间|(?:之)?间|between\s+(?:each\s+)?(?:keywords?|terms?|words?))?"
)
_SEP_VERB_RE = r"(?:均\s*|都\s*|一律\s*|要\s*|应\s*|需\s*|须\s*)*(?:用|以|采用|使用|为|with|use|using|separated\s+by)"
_SEP_TOKEN_RE = (
    r"(?P<cn_semi>(?:中文|全角)\s*分号(?:\s*[（(“\"']\s*[；;]\s*[）)”\"'])?)"
    r"|(?P<en_semi>(?:英文|西文|半角)\s*分号(?:\s*[（(“\"']\s*[；;]\s*[）)”\"'])?)"
    r"|(?P<semi>分号(?:\s*[（(“\"']\s*(?P<semi_char>[；;])\s*[）)”\"'])?|semicolons?)"
    r"|(?P<cn_comma>(?:中文|全角)\s*逗号(?:\s*[（(“\"']\s*[，,]\s*[）)”\"'])?)"
    r"|(?P<en_comma>(?:英文|西文|半角)\s*逗号(?:\s*[（(“\"']\s*[，,]\s*[）)”\"'])?)"
    r"|(?P<comma>逗号(?:\s*[（(“\"']\s*(?P<comma_char>[，,])\s*[）)”\"'])?|commas?)"
    r"|(?P<dunhao>顿号(?:\s*[（(“\"']\s*、\s*[）)”\"'])?)"
    r"|(?P<fw_space>全角空格)"
    r"|(?P<two_space>(?:两个|两|双)\s*空格)"
    r"|(?P<one_space>(?:一个|一|单)?\s*空格)"
    r"|(?P<quoted>[“\"']\s*(?P<quoted_val>[，；、]|;\s*|,\s*)\s*[”\"'])"
)
_SEP_SUFFIX_RE = r"(?:\s*(?:隔开|分隔|分开|分割|相隔|间隔))?"

_SCOPED_KEYWORD_SEP_RE = re.compile(
    rf"{_GENERAL_BETWEEN_RE}\s*{_SEP_VERB_RE}\s*(?:{_SEP_TOKEN_RE})\s*{_SEP_SUFFIX_RE}"
    rf"|{_GENERAL_BETWEEN_RE}\s*空\s*(?P<space_count>两|二|2|一|1)\s*(?:格|个?字符|个?空格)\s*{_SEP_SUFFIX_RE}",
    re.I,
)
_UNSCOPED_KEYWORD_SEP_RE = re.compile(
    rf"{_EXPLICIT_KEYWORD_BETWEEN_RE}\s*{_SEP_VERB_RE}\s*(?:{_SEP_TOKEN_RE})\s*{_SEP_SUFFIX_RE}"
    rf"|{_EXPLICIT_KEYWORD_BETWEEN_RE}\s*空\s*(?P<space_count>两|二|2|一|1)\s*(?:格|个?字符|个?空格)\s*{_SEP_SUFFIX_RE}",
    re.I,
)


def resolve_separator_match(match: re.Match[str], language: str | None = None) -> str:
    """Map a regex match from requirement wording to a canonical keyword separator."""
    groups = match.groupdict()
    if groups.get("cn_semi"):
        return "；"
    if groups.get("en_semi"):
        return "; "
    if groups.get("semi"):
        ch = groups.get("semi_char")
        if ch == ";":
            return "; "
        if ch == "；":
            return "；"
        return "; " if language == "en" else "；"
    if groups.get("cn_comma"):
        return "，"
    if groups.get("en_comma"):
        return ", "
    if groups.get("comma"):
        ch = groups.get("comma_char")
        if ch == ",":
            return ", "
        if ch == "，":
            return "，"
        return ", " if language == "en" else "，"
    if groups.get("dunhao"):
        return "、"
    if groups.get("fw_space"):
        return "\u3000"
    if groups.get("two_space"):
        return "  "
    if groups.get("one_space"):
        return " "
    if groups.get("quoted"):
        raw_val = groups["quoted_val"]
        val = raw_val.strip()
        if val == ";":
            return "; " if (len(raw_val) > 1 or language == "en") else ";"
        if val == ",":
            return ", " if (len(raw_val) > 1 or language == "en") else ","
        return val
    if groups.get("space_count"):
        return "  " if groups["space_count"] in {"两", "二", "2"} else " "
    raise FormatError("Unrecognized keyword separator phrase")


def match_keyword_separator_clauses(text: str, *, scoped: bool, language: str | None = None):
    """Yield (match, separator_value) from a requirement clause."""
    pattern = _SCOPED_KEYWORD_SEP_RE if scoped else _UNSCOPED_KEYWORD_SEP_RE
    for match in pattern.finditer(text):
        yield match, resolve_separator_match(match, language=language)
