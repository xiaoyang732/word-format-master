# Review repair implementation plan

> **For agentic workers:** Use superpowers:executing-plans to implement this plan task by task.

**Goal:** Resolve report E1–E6 in the existing engine and record verified changes in Git.

**Architecture:** Retain the central registry and its independent readers, verifiers and protection projections. Template profiles provide evidence and require explicit header/footer decisions. Preserve target paragraphs before expanding profile rules.

**Tech Stack:** Python, python-docx, lxml, unittest, Word/WPS rendering.

**Spec:** `C:/Users/yang/Documents/Codex/word format master/detailed_repair_report.md`

## Global constraints

- Work in `D:/project/word-format-master` on `fix/review-report-2026-11-02`; preserve existing history and do not push without a request.
- Keep `scope: all` as the compatible page-number default; support `scope: body` and an explicit one-based `start_section`.
- Do not infer institution text, alignment or page-number policy from template majority values.
- Tests construct their own documents and do not depend on temporary template caches.
- D1–D4 describe a separate thesis delivery workflow. This task modifies the project; do not overwrite those external documents.

## Review focus

- Explicit zero indents on ordinary paragraphs must override style inheritance.
- Numbering disabled with numId 0 must override style numbering; partial numPr and level overrides must resolve correctly.
- Whitespace-only text ranges must retain their actual font properties.
- First/even headers and linked footer parts must remain isolated; unrelated content and relationships survive.
- Declaration extraction ends at its semantic section boundary; preserve includes empty paragraphs and prevents subsequent formatting.

### Task 1: Effective paragraph and range properties (E4, E5, E6)

**Files:** `scripts/word_format/properties.py`, `scripts/word_format/inspect.py`, `_dev/tests/test_declaration_ref_hf.py`.

**Interfaces:** `ppr_sources(document, paragraph)` remains an iterator; add an internal origin-aware iterator. `DocumentIndex.font_values(..., start, end)` retains explicit-range runs.

- [ ] Replace cache fixtures with constructed numbering and declaration fixtures; add ordinary zero, style numPr, override, disabled numbering and whitespace range tests.
- [ ] Run the tests and confirm the reported failures.
- [ ] Resolve numbering origins and only skip a direct zero firstLineChars cancellation when an effective numbering hanging indent exists. Limit whitespace filtering to full-paragraph aggregation.
- [ ] Run targeted regression tests and commit.

### Task 2: Header evidence and border operation (E1)

**Files:** `scripts/word_format/{profiles,headers,registry}.py`, `scripts/analyze_docx.py`, `_dev/tests/test_headers.py`, header references.

**Interfaces:** `header.border_bottom.set` accepts `{value: boolean, variant: default|first|even}`. Effective border evidence includes style inheritance.

- [ ] Add failing tests for explicit header evidence, inherited Header size/border, border cancellation and variant isolation through plan/apply/verify.
- [ ] Remove guessed header/page-number rules; retain the explicit mapping issue and add paragraph-level effective evidence to header/footer parts.
- [ ] Register one border writer, independent reader and projection; preserve existing Header style when replacing content.
- [ ] Run header/profile tests, update contracts, and commit.

### Task 3: Scoped page numbering (E2)

**Files:** `scripts/word_format/{structures,spec,verification,registry}.py`, `_dev/tests/test_declaration_ref_hf.py`, page-number references.

**Interfaces:** `page_number.configure` adds `scope` and `start_section`, validates types/ranges and independently verifies the selected start section and all variants.

- [ ] Add failing multi-section tests for body scope, explicit starts, linked footers, first/even variants, later section restart removal, and wrong writer rejection.
- [ ] Detach before mutations, remove standalone front-matter PAGE paragraphs, reset decimal numbering at the actual start and continue later sections.
- [ ] Run page-number and full regression tests, update contracts, and commit.

### Task 4: Declaration boundaries and preservation (E3, E6)

**Files:** `scripts/word_format/{cover,profiles}.py`, `_dev/tests/test_declaration_ref_hf.py`, template workflow references.

**Interfaces:** declaration detection uses `classify_module`, not the count of encountered section breaks; `mapping.preserve` excludes preserved paragraphs from generated operations.

- [ ] Add failing extraction/replacement tests for page-break and section-break templates, plus explicit preserve tests covering blank lines.
- [ ] Fix boundary state transitions, keep existing spacing/indent OOXML while materializing missing effective formatting, and exclude preserved targets before rule compilation.
- [ ] Run cover/template tests and commit.

### Task 5: Integration and delivery

- [ ] Run the complete unittest suite and `_dev/scripts/validate_project.py`; check `git diff --check`.
- [ ] Exercise representative CLI plan/apply/verify and render every page of a constructed multi-section fixture, if a renderer is available.
- [ ] Review the final diff for contract consistency, record results and limitations, and commit documentation.
- [ ] Confirm a clean working tree and report the branch, commits and checks.
