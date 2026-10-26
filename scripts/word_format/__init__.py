"""Deterministic Word formatting for direct AI tools."""

from .api import apply_plan, audit_document, build_plan, capabilities, inspect_document, verify_plan
from .cover import add_standard_thesis_cover, apply_east_asia_font

__all__ = [
    "apply_plan",
    "audit_document",
    "build_plan",
    "capabilities",
    "inspect_document",
    "verify_plan",
    "add_standard_thesis_cover",
    "apply_east_asia_font",
]

