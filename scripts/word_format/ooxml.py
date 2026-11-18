"""Targeted Transitional WordprocessingML sequence rules used by our writers.

These are container rules, not a general schema validator. In particular CT_RPr
and CT_TrPr permit repeated choices and must never be sorted by a fixed table.
Unknown extension children are retained in place.
"""

from __future__ import annotations

from typing import Any

from .namespaces import qn

MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MC_IGNORABLE_ATTR = f"{{{MC_NAMESPACE}}}Ignorable"

SEQUENCES: dict[str, str] = {
    "numbering": "numPicBullet abstractNum num numIdMacAtCleanup",
    "abstractNum": "nsid multiLevelType tmpl name styleLink numStyleLink lvl",
    "lvl": "start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr",
    "num": "abstractNumId lvlOverride",
    "numPr": "ilvl numId numberingChange ins",
    "styles": "docDefaults latentStyles style",
    "docDefaults": "rPrDefault pPrDefault",
    "style": (
        "name aliases basedOn next link autoRedefine hidden uiPriority "
        "semiHidden unhideWhenUsed qFormat locked personal personalCompose "
        "personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr"
    ),
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
    "tblPr": (
        "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize "
        "tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd "
        "tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange"
    ),
    "tblBorders": "top start left bottom end right insideH insideV",
    "tcPr": (
        "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar "
        "textDirection tcFitText vAlign hideMark headers cellIns cellDel "
        "cellMerge tcPrChange"
    ),
    "tcBorders": "top start left bottom end right insideH insideV tl2br tr2bl",
}


def insert_child(parent: Any, child: Any) -> Any:
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


def undeclared_mc_ignorable(root: Any) -> list[str]:
    """Return ordered unique prefixes in mc:Ignorable missing from root.nsmap."""
    if root is None:
        return []
    raw = root.get(MC_IGNORABLE_ATTR)
    if not raw:
        return []
    declared = {prefix for prefix in (getattr(root, "nsmap", None) or {}) if prefix}
    return [prefix for prefix in dict.fromkeys(raw.split()) if prefix not in declared]


def sanitize_mc_ignorable(root: Any) -> list[str]:
    """Drop undeclared prefixes from mc:Ignorable and return the removed prefixes."""
    if root is None:
        return []
    raw = root.get(MC_IGNORABLE_ATTR)
    if raw is None:
        return []
    declared = {prefix for prefix in (getattr(root, "nsmap", None) or {}) if prefix}
    tokens = list(dict.fromkeys(raw.split()))
    kept = [prefix for prefix in tokens if prefix in declared]
    removed = [prefix for prefix in tokens if prefix not in declared]
    if kept:
        root.set(MC_IGNORABLE_ATTR, " ".join(kept))
    else:
        root.attrib.pop(MC_IGNORABLE_ATTR, None)
    return removed


def package_mc_ignorable_errors(document: Any) -> list[str]:
    """Validate that every XML part declares all prefixes named in its mc:Ignorable attribute."""
    errors: list[str] = []
    for part in document.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            continue
        missing = undeclared_mc_ignorable(root)
        if missing:
            errors.append(
                f"{part.partname}: undeclared mc:Ignorable prefixes ({', '.join(missing)})"
            )
    return errors


def sequence_warnings(root: Any) -> list[str]:
    """Return local rule warnings, without claiming full ECMA schema validation."""
    warnings: list[str] = []
    for parent in root.iter():
        order = SEQUENCES.get(parent.tag.rsplit("}", 1)[-1])
        if not order:
            continue
        positions = {qn("w:" + name): i for i, name in enumerate(order.split())}
        ranks = [positions[child.tag] for child in parent if child.tag in positions]
        if ranks != sorted(ranks):
            warnings.append(parent.tag.rsplit("}", 1)[-1] + ": child sequence")
    return warnings
