"""Explicit repair of confirmed numbering sequence defects in existing documents."""

import copy

from .contracts import FormatError, require_keys
from .namespaces import qn
from .ooxml import SEQUENCES, sequence_warnings
from .properties import xml_value

OWNED_CONTAINERS = {"numbering", "abstractNum", "lvl"}


def validate(params):
    require_keys(params, set(), label="numbering order repair")
    return {}


def normalize(root):
    for parent in root.iter():
        name = parent.tag.rsplit("}", 1)[-1]
        if name not in OWNED_CONTAINERS:
            continue
        positions = {qn("w:" + tag): i for i, tag in enumerate(SEQUENCES[name].split())}
        slots = [i for i, node in enumerate(parent) if node.tag in positions]
        ordered = sorted((parent[i] for i in slots), key=lambda node: positions[node.tag])
        children = list(parent)
        for position, child in zip(slots, ordered):
            children[position] = child
        parent[:] = children


def integrity(root):
    errors = []
    ids = {}
    for tag, attribute in (("abstractNum", "abstractNumId"), ("num", "numId")):
        values = [node.get(qn("w:" + attribute)) for node in root.findall(qn("w:" + tag))]
        if len(set(values)) != len(values) or any(value is None for value in values):
            errors.append("Duplicate or missing " + attribute)
        ids[tag] = set(values)
    for node in root.findall(qn("w:num")):
        ref = node.find(qn("w:abstractNumId"))
        if ref is None or ref.get(qn("w:val")) not in ids["abstractNum"]:
            errors.append("Dangling abstract numbering reference")
    for abstract in root.findall(qn("w:abstractNum")):
        levels = [node.get(qn("w:ilvl")) for node in abstract.findall(qn("w:lvl"))]
        if len(levels) != len(set(levels)):
            errors.append("Duplicate numbering level")
    return errors


def repair(index, ref, params):
    root = index.document.part.numbering_part.element
    errors = integrity(root)
    if errors:
        raise FormatError(
            "Numbering order repair cannot resolve identity/link defects: " + "; ".join(errors)
        )
    normalize(root)
    return [
        "Repaired known numbering container sequences; IDs, links and all other content retained"
    ]


def read(index, ref, cap):
    root = index.document.part.numbering_part.element
    return {"sequence_warnings": sequence_warnings(root), "integrity_errors": integrity(root)}


def verify(actual, params, cap):
    return not actual["sequence_warnings"] and not actual["integrity_errors"]


def projection(index):
    root = copy.deepcopy(index.document.part.numbering_part.element)
    normalize(root)
    return xml_value(root), {
        str(part.partname): part.blob
        for part in index.document.part.package.parts
        if str(part.partname) != "/word/numbering.xml"
    }
