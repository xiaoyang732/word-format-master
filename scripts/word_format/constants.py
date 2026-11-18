"""Centralized typography constants, style catalog, unit conversions, and pattern builders."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

# Physical and OOXML measurement constants
TWIPS_PER_INCH = 1440.0
POINTS_PER_INCH = 72.0
MM_PER_INCH = 25.4
EMU_PER_TWIP = 635.0
TWIPS_PER_PT = 20.0
LINE_UNITS_PER_MULTIPLE = 240.0
CHAR_UNITS_SCALE = 100.0
HALF_POINTS_PER_PT = 2.0
WORD_POINTS_PER_LINE = 12.0

# Comparison tolerances
NUMERIC_TOLERANCE = 0.025
FONT_SIZE_TOLERANCE_PT = 0.05
INDENT_TOLERANCE = 0.25
LINE_SPACING_TOLERANCE_PT = 0.5

# Default structural geometry
DEFAULT_TOC_RIGHT_TAB_TWIPS = 8306
DEFAULT_TOC_INDENT_STEP_CHARS = 2.0
DEFAULT_LIST_LEFT_INDENT_TWIPS = 720
DEFAULT_LIST_HANGING_INDENT_TWIPS = 360
THREE_LINE_OUTER_BORDER_EIGHTH_PT = 12
THREE_LINE_HEADER_BORDER_EIGHTH_PT = 6

# Word Format Master managed style catalog
STYLE_WFM_TOC_HEADING = "WFM TOC Heading"
STYLE_WFM_TOC_FIELD = "WFM TOC Field"
STYLE_WFM_TOC_BREAK = "WFM TOC Break"
TOC_MANAGED_STYLES = frozenset(
    {STYLE_WFM_TOC_HEADING, STYLE_WFM_TOC_FIELD, STYLE_WFM_TOC_BREAK}
)

STYLE_WFM_PAGE_NUMBER = "WFM Page Number"
STYLE_ID_WFM_PAGE_NUMBER = "WFMPageNumber"
STYLE_WFM_ABSTRACT_TITLE = "WFM Abstract Title"
STYLE_WFM_MODULE_HEADING = "WFM Module Heading"
STYLE_WFM_KEYWORDS = "WFM Keywords"

STYLE_FIGURE_CAPTION = "Figure Caption"
STYLE_TABLE_CAPTION = "Table Caption"
STYLE_BIBLIOGRAPHY = "Bibliography"

BIBLIOGRAPHY_STYLE_ALIASES = frozenset(
    {"bibliography", "书目", "書目", "references", "reference", "参考文献"}
)

BUILTIN_PARAGRAPH_STYLES = frozenset(
    {
        "normal",
        "title",
        "subtitle",
        "header",
        "footer",
        "caption",
        "bibliography",
        "list paragraph",
        "list number",
        "list bullet",
        "toc heading",
        *(f"heading {i}" for i in range(1, 10)),
        *(f"toc {i}" for i in range(1, 10)),
    }
)

TOC_ENTRY_STYLE_RE = re.compile(r"^(?:toc|目录)\s*([1-9])$", re.I)

# Caption recognition defaults and regex builder
DEFAULT_CAPTION_LABELS: dict[str, tuple[str, ...]] = {
    "figure": ("图", "Figure", "Fig."),
    "table": ("表", "Table"),
}
CAPTION_NUMBER_CORE_PATTERN = (
    r"(?:\d|[一二三四五六七八九十]+[\s.．\-－—_]*\d|[A-Za-z]\s*[.．\-－—_]\s*\d)"
)


def mm_to_twips(mm: float) -> int:
    return int(round(float(mm) * TWIPS_PER_INCH / MM_PER_INCH))


def twips_to_mm(twips: float) -> float:
    return float(twips) * MM_PER_INCH / TWIPS_PER_INCH


def mm_to_pt(mm: float) -> float:
    return float(mm) * POINTS_PER_INCH / MM_PER_INCH


def pt_to_twips(pt: float) -> int:
    return int(round(float(pt) * TWIPS_PER_PT))


def pt_to_half_points(pt: float) -> int:
    return int(round(float(pt) * HALF_POINTS_PER_PT))


def section_content_width_twips(section, fallback: int = DEFAULT_TOC_RIGHT_TAB_TWIPS) -> int:
    """Compute usable line width in twips from a python-docx Section when geometry is present."""
    if (
        section is None
        or section.page_width is None
        or section.left_margin is None
        or section.right_margin is None
    ):
        return fallback
    usable_emu = (
        int(section.page_width)
        - int(section.left_margin)
        - int(section.right_margin)
        - int(section.gutter or 0)
    )
    if usable_emu <= 0:
        return fallback
    return int(round(usable_emu / EMU_PER_TWIP))


def is_builtin_paragraph_style(name: str | None) -> bool:
    return (name or "").strip().casefold() in BUILTIN_PARAGRAPH_STYLES


def parse_toc_style_level(style_name: str | None) -> int | None:
    """Return 1..9 when ``style_name`` denotes a native TOC entry style (e.g. 'toc 1', '目录 2')."""
    match = TOC_ENTRY_STYLE_RE.fullmatch((style_name or "").strip())
    return int(match.group(1)) if match else None


def is_toc_entry_style(style_name: str | None) -> bool:
    return parse_toc_style_level(style_name) is not None


def wfm_module_heading_style(module: str | None) -> str:
    return f"WFM {module} Heading" if module else STYLE_WFM_MODULE_HEADING


def wfm_header_footer_style(kind: str) -> str:
    return f"WFM {kind.title()}"


def wfm_header_footer_style_id(kind: str) -> str:
    return f"WFM{kind.title()}"


def build_caption_pattern(labels: Iterable[str]) -> re.Pattern[str]:
    unique_labels = [label for label in dict.fromkeys(labels) if label]
    escaped = "|".join(re.escape(label) for label in unique_labels)
    return re.compile(rf"^\s*(?:{escaped})\s*{CAPTION_NUMBER_CORE_PATTERN}", re.I)


ALL_DEFAULT_CAPTION_RE = build_caption_pattern(
    DEFAULT_CAPTION_LABELS["figure"] + DEFAULT_CAPTION_LABELS["table"]
)


@dataclass(frozen=True)
class UniversalChineseBaseConfig:
    """Configurable typography defaults for ``normalize_universal_chinese_base``."""

    east_asia_body_font: str = "宋体"
    east_asia_heading_font: str = "黑体"
    latin_font: str = "Times New Roman"
    text_color: str = "000000"
    body_font_size_pt: float = 12.0
    reference_font_size_pt: float = 10.5
    title_font_size_pt: float = 16.0
    header_footer_font_size_pt: float = 9.0
    line_spacing_multiple: float = 1.25
    toc_indent_step_chars: float = DEFAULT_TOC_INDENT_STEP_CHARS
    h1_font_size_pt: float = 15.0
    h2_font_size_pt: float = 14.0
    h3_plus_font_size_pt: float = 12.0


DEFAULT_CHINESE_BASE_CONFIG = UniversalChineseBaseConfig()
