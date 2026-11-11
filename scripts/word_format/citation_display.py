"""Read and format numeric bibliography citations without changing their text or fields."""

from __future__ import annotations

import copy
import re
from collections import defaultdict
from dataclasses import dataclass
from functools import cached_property

from docx.oxml import OxmlElement

from .inspect import paragraph_text
from .namespaces import qn
from .numbering import effective_numbering, sequence_number
from .properties import element, rpr_sources


DISPLAY_VALUES = {"inline": "baseline", "superscript": "superscript"}
LABEL = re.compile(r"\[\s*([1-9][0-9]*)\s*\]")
REFERENCE_PREFIX = re.compile(r"^\s*(?:[\[［]\s*([1-9][0-9]*)\s*[\]］]|([1-9][0-9]*)[.)、])")
NATIVE_LABELS = {"[%1]", "［%1］", "%1", "%1.", "%1)", "%1、"}
REF_CODE = re.compile(r'^\s*REF\s+(?:"([^"]+)"|(\S+))(?=\s|$)', re.I)


def run_display(document, paragraph, run):
    for pr in rpr_sources(document, paragraph, run):
        node = pr.find(qn("w:vertAlign")) if pr is not None else None
        if node is not None:
            value = node.get(qn("w:val"))
            return "inline" if value == "baseline" else value or "unknown"
    return "inline"


def write_display(pr, display):
    element(pr, "vertAlign").set(qn("w:val"), DISPLAY_VALUES[display])


@dataclass(frozen=True)
class CitationMark:
    start: int
    end: int
    text: str
    display: str
    slices: tuple
    controls: tuple

    def evidence(self):
        return {"start": self.start, "end": self.end, "text": self.text, "display": self.display}


def is_citation_body(record):
    return (record.get("top_level") and not record.get("unsupported_container")
            and record.get("role") == "body" and record.get("module") in {None, "chapters"})


def citation_evidence(paragraph, marks, text):
    if not marks:
        return {}
    source = "".join(node.text or "" for node in paragraph.iter(qn("w:t")))
    result = {"citations": [mark.evidence() for mark in marks]}
    if source != text:
        result["citation_source_text"] = source
    return result


class CitationReader:
    def __init__(self, index):
        self.document = index.document
        self.index = index

    @cached_property
    def references(self):
        return self.index.reference_nodes()

    @cached_property
    def bookmarks(self):
        return {node.get(qn("w:name")) for paragraph in self.references
                for node in paragraph.iter(qn("w:bookmarkStart"))}

    @cached_property
    def labels(self):
        labels = set()
        sequences = set()
        for paragraph in self.references:
            text = paragraph_text(paragraph)
            prefix = REFERENCE_PREFIX.match(text)
            if prefix:
                labels.add(prefix[1] or prefix[2])
            elif text.strip() and (number := effective_numbering(self.document, paragraph)):
                sequences.add(number[0])
        if sequences:
            counters = {}
            for paragraph in self.document.element.body.iter(qn("w:p")):
                if any(node.tag == qn("w:txbxContent") for node in paragraph.iterancestors()):
                    continue
                number = effective_numbering(self.document, paragraph)
                if number is None or number[0] not in sequences:
                    continue
                state = sequence_number(number, counters)
                if (paragraph in self.references and state["level"] == 0 and state["format"] == "decimal"
                        and state["marker"] in NATIVE_LABELS):
                    labels.add(str(state["value"]))
        return labels

    def marks(self, oid):
        if not is_citation_body(self.index.records[oid]):
            return []
        paragraph = self.index.elements[oid]
        if not any("[" in (node.text or "") for node in paragraph.iter(qn("w:t"))) or not self.references:
            return []
        marks = []
        offset = 0
        ordinary = []
        fields = []

        def flush():
            if not ordinary:
                return
            text = "".join(part[2] for part in ordinary)
            for match in LABEL.finditer(text):
                start = ordinary[0][1] + match.start()
                end = ordinary[0][1] + match.end()
                parts = [part for part in ordinary if part[1] < end and part[1] + len(part[2]) > start]
                if not any(part[3] is not None for part in parts) and match[1] not in self.labels:
                    continue
                slices = tuple((run, max(start - pos, 0), min(end - pos, len(value)))
                               for run, pos, value, _ in parts)
                controls = {run for part in parts for run in part[3] or ()}
                displays = {run_display(self.document, paragraph, run) for run, _, _ in slices}
                display = displays.pop() if len(displays) == 1 else "mixed"
                marks.append(CitationMark(start, end, match[0], display, slices, tuple(controls)))
            ordinary.clear()

        def collect_field(parts, code, controls=()):
            target = REF_CODE.match(code)
            text = "".join(part[2] for part in parts)
            if (target and (target[1] or target[2]) in self.bookmarks
                    and (LABEL.fullmatch(text) or re.fullmatch(r"[1-9][0-9]*", text))):
                ordinary.extend((*part, controls) for part in parts)
            else:
                flush()

        def run_text(run):
            return "".join(node.text or "" for node in run.iter(qn("w:t")))

        def is_text_run(run):
            return run.tag == qn("w:r") and all(
                node.tag in {qn("w:rPr"), qn("w:t")} for node in run
            )

        def read_run(run):
            nonlocal offset
            text = run_text(run)
            if text and is_text_run(run):
                if not fields:
                    ordinary.append((run, offset, text, None))
                elif fields[-1]["result"]:
                    fields[-1]["parts"].append((run, offset, text))
            elif any(node.tag != qn("w:rPr") for node in run):
                if not fields and not any(node.tag == qn("w:fldChar") for node in run):
                    flush()
                for node in run:
                    if node.tag == qn("w:fldChar"):
                        kind = node.get(qn("w:fldCharType"))
                        if kind == "begin":
                            if fields:
                                fields[-1]["supported"] = False
                            fields.append({"code": "", "result": False, "parts": [],
                                           "controls": {run}, "supported": not text})
                        elif kind == "separate" and fields:
                            fields[-1]["result"] = True
                            fields[-1]["controls"].add(run)
                        elif kind == "end" and fields:
                            field = fields.pop()
                            if not fields:
                                if field["supported"] and not text:
                                    collect_field(field["parts"], field["code"], field["controls"] | {run})
                                else:
                                    flush()
                    elif node.tag == qn("w:instrText") and fields:
                        fields[-1]["code"] += node.text or ""
                        fields[-1]["controls"].add(run)
                    elif fields and node.tag != qn("w:rPr"):
                        fields[-1]["supported"] = False
            offset += len(text)

        def visit(container):
            nonlocal offset
            for child in container:
                if child.tag == qn("w:r"):
                    read_run(child)
                elif (child.tag == qn("w:hyperlink") and child.get(qn("r:id")) is None
                      and child.get(qn("w:anchor")) in self.bookmarks):
                    visit(child)
                elif child.tag == qn("w:fldSimple"):
                    if fields:
                        fields[-1]["supported"] = False
                    parts = []
                    for run in child:
                        text = run_text(run)
                        if text:
                            parts.append((run, offset, text))
                        offset += len(text)
                    if not fields:
                        if all(is_text_run(run) for run in child):
                            collect_field(parts, child.get(qn("w:instr"), ""))
                        else:
                            flush()
                elif child.tag not in {qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")}:
                    flush()
                    if fields:
                        fields[-1]["supported"] = False
                    offset += sum(len(node.text or "") for node in child.iter(qn("w:t")))

        visit(paragraph)
        flush()
        return marks


def read_citation_display(index, ref, cap):
    paragraph = index.elements[ref["id"]]
    displays = set()
    for mark in CitationReader(index).marks(ref["id"]):
        displays.add(mark.display)
        displays.update(run_display(index.document, paragraph, run) for run in mark.controls)
    return sorted(displays)


def set_citation_display(index, ref, params):
    by_run = defaultdict(list)
    for mark in CitationReader(index).marks(ref["id"]):
        for control in mark.controls:
            write_display(control.get_or_add_rPr(), params["value"])
        for run, start, end in mark.slices:
            by_run[run].append((start, end))
    for run, ranges in by_run.items():
        text = "".join(node.text or "" for node in run.findall(qn("w:t")))
        if ranges == [(0, len(text))]:
            write_display(run.get_or_add_rPr(), params["value"])
            continue
        cursor = 0
        for start, end in [*ranges, (len(text), len(text))]:
            for lo, hi, selected in ((cursor, start, False), (start, end, True)):
                if lo >= hi:
                    continue
                clone = copy.deepcopy(run)
                for node in clone.findall(qn("w:t")):
                    clone.remove(node)
                node = OxmlElement("w:t")
                node.set(qn("xml:space"), "preserve")
                node.text = text[lo:hi]
                clone.append(node)
                if selected:
                    write_display(clone.get_or_add_rPr(), params["value"])
                run.addprevious(clone)
            cursor = end
        run.getparent().remove(run)
