---
name: wfm-feature-development
description: Use when implementing or changing Word Format Master formatting operations, CLI behavior, or format-engine features.
---

# Word Format Master feature development

Input: a feature or behavior change in the Word Format Master repository. Output: the requested change, relevant verification results, and any necessary contract documentation updates.

1. Read the repository's `references/architecture.md` and relevant sections of `references/direct-format-api.md`, `references/format-spec.md`, or `references/template-engine.md`. Inspect the existing implementation and capabilities before designing a change; do not assume an operation is absent.
2. Trace the affected path through `scripts/format_cli.py`, `scripts/word_format/api.py`, `registry.py`, and the relevant property, operation, structure, or template modules. Keep all format operations on the central registry's existing write path. Do not add a parallel DOCX writer.
3. For a new or changed operation, check target scope, parameter validation, handler, independent reader, verifier, effects and conflict/protection behavior. Preserve document content, unrelated structures and relationships, and the input document unless the user explicitly authorizes otherwise.
4. Design meaningful regression coverage for the changed behavior and failure cases in the existing `dev/tests/` suite. Run relevant tests and `dev/scripts/validate_project.py`; expand checks when the change warrants it. Record failures or unavailable dependencies honestly.
5. If contracts or user-visible capabilities change, update the relevant references, `README.md`, root `SKILL.md`, and `agents/openai.yaml` as applicable. Check `capabilities` metadata against actual tests. Do not modify documentation merely to satisfy a checklist.
6. For visual changes, render representative outputs and inspect every page before claiming visual verification. If rendering or inspection is unavailable, report it as pending or skipped.

Acceptance: the requested behavior is supported through the existing architecture, relevant tests and independent verification have been run or their limitations reported, and affected contracts remain consistent.
