"""Targeted Transitional WordprocessingML sequence rules used by our writers.

These are container rules, not a general schema validator. In particular CT_RPr
and CT_TrPr permit repeated choices and must never be sorted by a fixed table.
Unknown extension children are retained in place.
"""

from .namespaces import qn

SEQUENCES = {
    "numbering": "numPicBullet abstractNum num numIdMacAtCleanup",
    "abstractNum": "nsid multiLevelType tmpl name styleLink numStyleLink lvl",
    "lvl": "start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr",
    "num": "abstractNumId lvlOverride",
    "pPr": (
        "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
        "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap "
        "overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
        "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap "
        "jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle "
        "rPr sectPr pPrChange"
    ),
    "sectPr": (
        "headerReference footerReference footnotePr endnotePr type pgSz pgMar "
        "paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote "
        "titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange"
    ),
    "pBdr": "top left bottom right between bar",
}


def insert_child(parent, child):
    """Insert one owned child before its known successors without reordering peers."""
    order = SEQUENCES.get(parent.tag.rsplit("}", 1)[-1], "").split()
    positions = {qn("w:" + name): position for position, name in enumerate(order)}
    rank = positions.get(child.tag)
    if rank is not None:
        for position, sibling in enumerate(parent):
            if positions.get(sibling.tag, -1) > rank:
                parent.insert(position, child)
                return child
    parent.append(child)
    return child


def sequence_warnings(root):
    """Return local rule warnings, without claiming full ECMA schema validation."""
    warnings = []
    for parent in root.iter():
        order = SEQUENCES.get(parent.tag.rsplit("}", 1)[-1])
        if not order:
            continue
        positions = {qn("w:" + name): i for i, name in enumerate(order.split())}
        ranks = [positions[child.tag] for child in parent if child.tag in positions]
        if ranks != sorted(ranks):
            warnings.append(parent.tag.rsplit("}", 1)[-1] + ": child sequence")
    return warnings
