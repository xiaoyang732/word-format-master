---
name: docx-regression-testing
description: Use when designing, running, or investigating regression and integrity tests for the Word Format Master formatting engine or DOCX output.
---

# Word Format Master regression testing

Input: a change, suspected regression, or requested coverage assessment. Output: relevant tests and an evidence-based report of results and gaps.

1. Inspect `references/architecture.md`, the affected registry entries and existing `dev/tests/` before choosing tests. The suite currently uses `unittest`; do not presume pytest or migrate it without an explicit reason. Check registry test metadata against actual files and test behavior rather than interpreting a filename as proof of coverage.
2. Select meaningful cases across applicable layers: parameter validation and operation/readback; source preservation, content and unrelated XML/relationship integrity; `inspect`, profile, plan, apply and verify integration; and visual behavior. Include failure, no-op and boundary cases when relevant. Keep fixtures minimal and representative.
3. Run targeted tests first, then the appropriate broader suite. Existing commands include `python dev/scripts/validate_project.py`, `python dev/scripts/smoke_test.py` and `python -m unittest discover -s dev/tests -v`. Use the project's configured environment; report command, outcome, failures and skipped checks without inventing results.
4. For page-layout regressions, retain render manifests and inspect each page. Compare using a controlled renderer version and fonts where possible; renderer differences alone do not establish a code regression. Mock visual evidence and structural checks are not actual page reviews.
5. Report confirmed coverage and concrete gaps separately; avoid claiming all operations are covered without checking each one. Preserve input documents and avoid modifying unrelated fixtures.

Acceptance: tests exercise user-visible behavior and protection boundaries, results are reproducible to the extent available, and incomplete visual or environment checks are clearly identified.
