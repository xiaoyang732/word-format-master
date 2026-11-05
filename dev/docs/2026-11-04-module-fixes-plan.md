# Modular formatting audit fixes

Goal: resolve F1–F5 in the 2026-11-03 audit using the existing registered plan/apply/verify engine.

Design: scope legacy body requests to actual selected paragraphs rather than inherited Normal defaults. Represent numeric citations with validated reference identities, native bibliography numbering, bookmarks and REF fields. Share bibliography boundaries with document module detection. Extend paragraph semantics and atomic operations for contextual spacing, grid snapping, mirrored indentation and full tab-stop attributes; continue reporting other unsupported XML.

Constraints: preserve source files and unrelated content/objects; do not force black color for local requests that omit color. Keep native DOCX structure. Reject ambiguous/complex citation anchors instead of flattening objects. No publishing or commits requested.

- [x] F1: add failing effective-format scope tests, remove shared-style leakage, run targeted scope tests.
- [x] F2/F3/F4: test dynamic fields/bookmarks, mismatched display rejection, bibliography end boundaries, repeat insertion and malformed anchors; implement citation writer plus independent structure checks.
- [x] F5: test semantic read/profile/write/verify for extended paragraph attributes and tab leaders; implement registry-backed operations and validators, document remaining limits.
- [x] Run all regression tests, validation and smoke test; rerun original acceptance probes with updated supported-property assertions and dynamic-citation expectations.
- [x] Render representative black-heading thesis-format sample and update fields using installed WPS; inspect output and verify references still target bibliography bookmarks.
- [x] Review diffs and document results and remaining limits.
