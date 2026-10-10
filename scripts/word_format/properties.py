"""Single owners for atomic Word properties and their independent XML readers."""

from __future__ import annotations

import copy
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .contracts import FormatError

FONT_TAGS = {
    "font_east_asia": ("rFonts",), "font_latin": ("rFonts",),
    "font_size_pt": ("sz", "szCs"), "bold": ("b", "bCs"),
    "italic": ("i", "iCs"), "color": ("color",),
}
PARAGRAPH_TAGS = {
    "alignment": ("jc",), "space_before_pt": ("spacing",),
    "space_after_pt": ("spacing",), "line_spacing": ("spacing",),
    "first_line_indent": ("ind",), "left_indent": ("ind",),
    "right_indent": ("ind",), "hanging_indent": ("ind",),
    "keep_with_next": ("keepNext",), "keep_together": ("keepLines",),
    "page_break_before": ("pageBreakBefore",), "widow_control": ("widowControl",),
}
ATTRIBUTES = {
    "font_east_asia": {"rFonts": ("eastAsia", "eastAsiaTheme")},
    "font_latin": {"rFonts": ("ascii", "hAnsi", "asciiTheme", "hAnsiTheme")},
    "space_before_pt": {"spacing": ("before", "beforeLines", "beforeAutospacing")},
    "space_after_pt": {"spacing": ("after", "afterLines", "afterAutospacing")},
    "line_spacing": {"spacing": ("line", "lineRule")},
    "first_line_indent": {"ind": ("firstLine", "firstLineChars", "hanging", "hangingChars")},
    "hanging_indent": {"ind": ("firstLine", "firstLineChars", "hanging", "hangingChars")},
    "left_indent": {"ind": ("left", "leftChars", "start", "startChars")},
    "right_indent": {"ind": ("right", "rightChars", "end", "endChars")},
}


def xml_value(node) -> str:
    """Semantic representation independent of namespace prefix/attribute ordering."""
    if node is None:
        return ""
    def value(n):
        return (n.tag,tuple(sorted(n.attrib.items())),n.text or "",tuple(value(c) for c in n))
    return repr(value(node))


def element(parent, name: str):
    child = parent.find(qn("w:" + name))
    if child is None:
        child = OxmlElement("w:" + name)
        parent.append(child)
    return child


def sanitize_properties(pr, property_name: str):
    clone = copy.deepcopy(pr)
    if clone is None:
        return ""
    tags = FONT_TAGS.get(property_name, PARAGRAPH_TAGS.get(property_name, ()))
    for tag in tags:
        node = clone.find(qn("w:" + tag))
        if node is None:
            continue
        if property_name in ATTRIBUTES:
            for attr in ATTRIBUTES[property_name].get(tag, ()):
                node.attrib.pop(qn("w:" + attr), None)
            if not node.attrib and not len(node):
                clone.remove(node)
        else:
            clone.remove(node)
    return xml_value(clone) if len(clone) or clone.attrib else ""


def rpr_sources(document, paragraph, run=None):
    if run is not None:
        pr = run.find(qn("w:rPr"))
        if pr is not None:
            yield pr
            char_style = pr.find(qn("w:rStyle"))
            if char_style is not None:
                style_id = char_style.get(qn("w:val"))
                style = next((s for s in document.styles if s.style_id == style_id), None)
                seen = set()
                while style is not None and style.style_id not in seen:
                    seen.add(style.style_id)
                    yield style.element.find(qn("w:rPr"))
                    style = style.base_style
    style_id = None
    ppr = paragraph.find(qn("w:pPr"))
    if ppr is not None:
        if run is None:
            yield ppr.find(qn("w:rPr"))
        pstyle = ppr.find(qn("w:pStyle"))
        style_id = pstyle.get(qn("w:val")) if pstyle is not None else None
    style = next((s for s in document.styles if s.style_id == (style_id or "Normal")), None)
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style.element.find(qn("w:rPr"))
        style = style.base_style
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        yield defaults.find("./" + qn("w:rPrDefault") + "/" + qn("w:rPr"))


def ppr_sources(document, paragraph):
    pr = paragraph.find(qn("w:pPr"))
    yield pr
    pstyle = pr.find(qn("w:pStyle")) if pr is not None else None
    style_id = pstyle.get(qn("w:val")) if pstyle is not None else "Normal"
    style = next((s for s in document.styles if s.style_id == style_id), None)
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style.element.find(qn("w:pPr"))
        style = style.base_style
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        yield defaults.find("./" + qn("w:pPrDefault") + "/" + qn("w:pPr"))


def _attr(sources, tag, attr="val"):
    for pr in sources:
        node = pr.find(qn("w:" + tag)) if pr is not None else None
        if node is not None:
            value = node.get(qn("w:" + attr))
            if value is not None:
                return value
    return None


def read_font(document, paragraph, run, name):
    sources = list(rpr_sources(document, paragraph, run))
    if name in {"font_east_asia", "font_latin"}:
        attr = "eastAsia" if name == "font_east_asia" else "ascii"
        # Theme attributes take precedence within each layer, not across layers.
        for pr in sources:
            node = pr.find(qn("w:rFonts")) if pr is not None else None
            if node is None:
                continue
            theme = node.get(qn("w:" + attr + "Theme"))
            if theme:
                from analyze_docx import parse_theme_fonts
                # Reuse the already loaded theme part without a filesystem write.
                class Parts:
                    def read(self, path):
                        return next(p.blob for p in document.part.package.parts if str(p.partname).lstrip("/") == path)
                try:
                    fonts = parse_theme_fonts(Parts())
                    return fonts.get(theme, theme)
                except (StopIteration, KeyError):
                    return theme
            value = node.get(qn("w:" + attr))
            if value:
                return value
        return None
    if name == "font_size_pt":
        raw = _attr(sources, "sz")
        return float(raw) / 2 if raw is not None else None
    if name in {"bold", "italic"}:
        # OOXML style b/i are toggle properties; direct rPr values are absolute.
        tag = "b" if name == "bold" else "i"
        direct = run.find(qn("w:rPr")) if run is not None else None
        node = direct.find(qn("w:" + tag)) if direct is not None else None
        if node is not None:
            return node.get(qn("w:val"), "1") not in {"0", "false", "off"}
        state = False
        for pr in reversed(sources):
            node = pr.find(qn("w:" + tag)) if pr is not None else None
            if node is not None and node.get(qn("w:val"), "1") not in {"0", "false", "off"}:
                state = not state
        return state
    return _attr(sources, "color") or "auto"


def write_font(pr, name, value):
    if name in {"font_east_asia", "font_latin"}:
        fonts = element(pr, "rFonts")
        attrs = ("eastAsia",) if name == "font_east_asia" else ("ascii", "hAnsi")
        for attr in attrs:
            fonts.set(qn("w:" + attr), value)
            fonts.attrib.pop(qn("w:" + attr + "Theme"), None)
    elif name == "font_size_pt":
        for tag in ("sz", "szCs"):
            element(pr, tag).set(qn("w:val"), str(int(round(value * 2))))
    elif name in {"bold", "italic"}:
        for tag in (("b", "bCs") if name == "bold" else ("i", "iCs")):
            element(pr, tag).set(qn("w:val"), "1" if value else "0")
    elif name == "color":
        node = element(pr, "color")
        node.set(qn("w:val"), value.lstrip("#").upper() if value != "auto" else value)
        for attr in ("themeColor", "themeTint", "themeShade"):
            node.attrib.pop(qn("w:" + attr), None)


def read_paragraph(document, paragraph, name):
    sources = list(ppr_sources(document, paragraph))
    if name == "alignment":
        value = _attr(sources, "jc") or "left"
        return "justify" if value == "both" else value
    if name in {"space_before_pt", "space_after_pt"}:
        raw = _attr(sources, "spacing", "before" if name == "space_before_pt" else "after")
        return float(raw or 0) / 20
    if name == "line_spacing":
        rule = _attr(sources, "spacing", "lineRule") or "auto"
        raw = float(_attr(sources, "spacing", "line") or 240)
        return {"kind": {"auto": "multiple", "exact": "exact", "atLeast": "at_least"}.get(rule, rule),
                "value": raw / (240 if rule == "auto" else 20), "unit": "multiple" if rule == "auto" else "pt"}
    if name.endswith("indent"):
        base = {"first_line_indent": "firstLine", "left_indent": "left", "right_indent": "right", "hanging_indent": "hanging"}[name]
        for pr in sources:
            ind=pr.find(qn("w:ind")) if pr is not None else None
            if ind is None: continue
            # A direct first-line/hanging declaration replaces the mutually
            # exclusive property inherited from a style. Hanging wins when a
            # malformed layer declares both (including hanging="0").
            if base in {"firstLine", "hanging"}:
                opposite="hanging" if base=="firstLine" else "firstLine"
                has_opposite=any(ind.get(qn("w:"+opposite+suffix)) is not None for suffix in ("", "Chars"))
                has_own=any(ind.get(qn("w:"+base+suffix)) is not None for suffix in ("", "Chars"))
                if has_opposite and base=="firstLine":
                    chars=ind.get(qn("w:hangingChars"))
                    raw=ind.get(qn("w:hanging"))
                    # A hanging indent is a negative first-line offset. Reading
                    # it as zero would incorrectly skip a request to clear it.
                    if chars is not None: return {"value":-float(chars)/100,"unit":"chars"}
                    return {"value":-float(raw or 0)*25.4/1440,"unit":"mm"}
                if has_opposite and not has_own:
                    chars=ind.get(qn("w:firstLineChars"))
                    raw=ind.get(qn("w:firstLine"))
                    if chars is not None: return {"value":-float(chars)/100,"unit":"chars"}
                    return {"value":-float(raw or 0)*25.4/1440,"unit":"mm"}
            chars=ind.get(qn("w:"+base+"Chars"))
            raw=ind.get(qn("w:"+base))
            if chars is not None: return {"value":float(chars)/100,"unit":"chars"}
            if raw is not None: return {"value":float(raw)*25.4/1440,"unit":"mm"}
        return {"value":0,"unit":"mm"}
    tag = PARAGRAPH_TAGS[name][0]
    for pr in sources:
        node = pr.find(qn("w:" + tag)) if pr is not None else None
        if node is not None:
            return node.get(qn("w:val"), "1") not in {"0", "false", "off"}
    return name == "widow_control"


def write_paragraph(pr, name, value):
    if name == "alignment":
        element(pr, "jc").set(qn("w:val"), "both" if value == "justify" else value)
    elif name in {"space_before_pt", "space_after_pt"}:
        attr = "before" if name == "space_before_pt" else "after"
        spacing = element(pr, "spacing")
        for key in (attr + "Lines", attr + "Autospacing"):
            spacing.attrib.pop(qn("w:" + key), None)
        spacing.set(qn("w:" + attr), str(int(round(value * 20))))
    elif name == "line_spacing":
        node = element(pr, "spacing")
        factor = 240 if value["kind"] == "multiple" else 20
        node.set(qn("w:line"), str(int(round(value["value"] * factor))))
        node.set(qn("w:lineRule"), {"multiple": "auto", "exact": "exact", "at_least": "atLeast"}[value["kind"]])
    elif name.endswith("indent"):
        node = element(pr, "ind")
        base = {"first_line_indent": "firstLine", "left_indent": "left", "right_indent": "right", "hanging_indent": "hanging"}[name]
        for attr in ATTRIBUTES[name]["ind"]:
            node.attrib.pop(qn("w:" + attr), None)
        # ATTRIBUTES removes the opposite indent. Do not write its zero value:
        # hanging="0" still takes precedence over a nonzero firstLineChars.
        # A direct declaration already replaces the opposite style declaration.
        if value["unit"] == "chars":
            node.set(qn("w:" + base + "Chars"), str(int(round(value["value"] * 100))))
            # Keep character units authoritative; a fixed 12pt conversion is
            # wrong for other font sizes and unnecessary in Word/WPS.
            if not value["value"]:
                node.set(qn("w:" + base), "0")
        else:
            node.set(qn("w:" + base + "Chars"), "0")
            # A zero character value can override a nonzero twips value in Word.
            # Remove it for positive physical indents, using 0 only when clearing.
            if value["value"]:
                node.attrib.pop(qn("w:" + base + "Chars"), None)
            node.set(qn("w:" + base), str(int(round(value["value"] * 1440 / 25.4))))
    else:
        element(pr, PARAGRAPH_TAGS[name][0]).set(qn("w:val"), "1" if value else "0")


def text_runs(paragraph):
    """Visible text runs, including hyperlink runs, excluding fields/revisions."""
    for run in paragraph.iter(qn("w:r")):
        if any(a.tag in {qn("w:del"), qn("w:ins"), qn("w:txbxContent")} for a in run.iterancestors() if a is not paragraph):
            continue
        yield run


def simple_range(paragraph, start, end, split=False):
    runs = list(text_runs(paragraph))
    if any(child.tag not in {qn("w:rPr"), qn("w:t")} for r in runs for child in r):
        raise FormatError("Text ranges containing fields, drawings, tabs or breaks require a supported explicit operation", "unsupported")
    result = []
    offset = 0
    for run in runs:
        text = "".join(t.text or "" for t in run.findall(qn("w:t")))
        length = len(text)
        lo, hi = max(0, start - offset), min(length, end - offset)
        if lo < hi:
            if split and (lo or hi < length):
                parent = run.getparent()
                pos = parent.index(run)
                segments = [(text[:lo], False), (text[lo:hi], True), (text[hi:], False)]
                for segment, selected in segments:
                    if not segment:
                        continue
                    clone = copy.deepcopy(run)
                    for child in list(clone):
                        if child.tag != qn("w:rPr"):
                            clone.remove(child)
                    t = OxmlElement("w:t")
                    t.set(qn("xml:space"), "preserve")
                    t.text = segment
                    clone.append(t)
                    parent.insert(pos, clone)
                    pos += 1
                    if selected:
                        result.append(clone)
                parent.remove(run)
            else:
                result.append(run)
        offset += length
    return result


def strip_text_prefix(paragraph, count):
    """Remove only a plain text prefix, preserving bookmarks, hyperlinks and fields."""
    remaining = count
    for run in text_runs(paragraph):
        for child in run:
            if child.tag == qn("w:rPr"):
                continue
            if not remaining:
                return
            if child.tag != qn("w:t"):
                raise FormatError("Numbering prefix crosses a field or drawing", "unsupported")
            text = child.text or ""
            removed = min(remaining, len(text))
            child.text = text[removed:]
            remaining -= removed
    if remaining:
        raise FormatError("Numbering prefix could not be mapped to text nodes", "unsupported")
