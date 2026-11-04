# Module formatting audit repair results

Baseline: `094f0811366cdd16edea70c86753134d44c6d26d`.

Resolved all five findings from the 2026-11-03 modular formatting audit:

- Legacy body requests preserve effective formatting outside chapter/body scope.
- Numeric citations use native bibliography numbering, stable bookmarks and independently verified REF fields.
- Citation display IDs must match reference targets; invalid requests do not publish output.
- Reference extraction shares module boundaries and stops at acknowledgements/appendices.
- Contextual spacing, grid snapping, mirrored indents and tab stops/leaders round-trip through inspection, profiles, registered paragraph/header operations and verification.

Validation on 2026-11-04:

- Full unittest discovery: **260 passed**, no failures/errors/skips (244.685 seconds).
- Original acceptance probes: **34 passed** (previously 29 passed, 5 failed).
- CLI inspect/plan/apply/verify/audit and invalid-input rejection: **6 checks passed**.
- Project validation: **75 registered operations**. Smoke test and diff whitespace checks passed.
- Installed WPS rendered the final black-heading sample correctly. Swapping bibliography entries and updating fields changed citations from `[1][2]` to `[2][1]`, preserving reference identity.

Portable dynamic citations use complete labels such as `[1][2]` or `[1],[2]`. Grouped/range syntax and author-date citations are explicitly unsupported. Complex anchors, bibliography numbering shared with outside lists and mixed numbering-marker typography require explicit normalization rather than risking unrelated changes. Other unsupported OOXML remains evidence/issues, not assumed migration support.

Reproduce from the repository root:

```powershell
python dev/scripts/validate_project.py
python dev/scripts/smoke_test.py
python -m unittest discover -s dev/tests -v
```

New focused tests: `test_body_scope_fixes.py`, `test_dynamic_citations.py`, `test_paragraph_extended.py`. Existing core/thesis/module tests were updated to cover new capabilities and effective inherited values instead of assuming redundant direct formatting.

Full local evidence and the user-facing Chinese report are retained at:
`C:/Users/yang/Documents/Codex/word format master/fix-report-2026-11-04/`.

Tests and these repair documents are intentionally tracked despite the existing `/dev/` ignore rule. Unrelated local development artifacts, document outputs and caches remain outside this change.
