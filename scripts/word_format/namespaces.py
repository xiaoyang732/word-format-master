"""Bounded cache for immutable OOXML qualified names."""

from functools import lru_cache

from docx.oxml.ns import qn as _qn

qn = lru_cache(maxsize=256)(_qn)
