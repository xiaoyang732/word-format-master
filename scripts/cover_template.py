# -*- coding: utf-8 -*-
"""
高校学术学位论文/毕业设计通用精美封面模板模块
通用规范、排版美观合理，仅展示必要学术信息（标题、作者、学号、学院、专业、指导教师、日期等）。
无绑定特定高校硬编码，排版疏密得体，适用于演示与初版通用规范。
"""

from __future__ import annotations

import sys
from pathlib import Path

# 确保脚本脱离包安装路径或从不同工作目录直接运行时也能正确定位 word_format
_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from word_format.cover import add_standard_thesis_cover, apply_east_asia_font

__all__ = [
    "add_standard_thesis_cover",
    "apply_east_asia_font",
]
