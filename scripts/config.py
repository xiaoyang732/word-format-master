from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

# Paths
SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_RUNTIME_ROOT = SKILL_DIR / ".runtime"
ASSETS_DIR = SKILL_DIR / "assets"
PRESETS_PATH = ASSETS_DIR / "presets.json"

# File size limits
MAX_DOCUMENT_BYTES = 30 * 1024 * 1024  # 30 MiB for DOCX/DOTX files

# Archive limits
MAX_PARTS = 5000  # Maximum number of parts in a DOCX zip archive
MAX_UNCOMPRESSED = 250 * 1024 * 1024  # 250 MiB uncompressed size limit

# Rendering
RENDER_METHODS = {"auto", "word", "libreoffice"}
RENDER_ORDER = ["word", "libreoffice"]
VISUAL_MODEL_SELECTIONS = {"auto", "current", "custom"}
SUPPORTED_SPEC_SCHEMA_VERSIONS = {"1.0"}

# Numbered list styles: mapping to (w:numFmt, w:lvlText)
LIST_NUMBERING_STYLES = {
    "decimal-period": ("decimal", "%1."),
    "decimal-parenthesis": ("decimal", "%1)"),
    "decimal-parentheses": ("decimal", "(%1)"),
    "decimal-fullwidth-parentheses": ("decimal", "（%1）"),
    "decimal-bracket": ("decimal", "[%1]"),
    "chinese-period": ("chineseCounting", "%1、"),
    "chinese-parentheses": ("chineseCounting", "（%1）"),
    "upper-alpha-period": ("upperLetter", "%1."),
    "lower-alpha-parenthesis": ("lowerLetter", "%1)"),
    "upper-roman-period": ("upperRoman", "%1."),
}


def utc_now() -> str:
    """Return current ISO formatted UTC timestamp."""
    return datetime.now(timezone.utc).isoformat()
