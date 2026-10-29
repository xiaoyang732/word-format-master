"""Compatibility imports; the cover writer lives in the format core."""

import sys
from pathlib import Path

_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from word_format.cover import add_standard_thesis_cover, apply_east_asia_font, render_cover  # noqa: E402

__all__ = ["add_standard_thesis_cover", "apply_east_asia_font", "render_cover"]
