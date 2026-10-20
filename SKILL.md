---
name: word-format-master
description: Inspect, audit and edit Word formatting through deterministic local code. Use direct AI instructions, requirement documents or formatted DOCX/DOTX samples to build evidence-bearing profiles, resolve mappings, plan, apply and verify. Preserve source documents; convert legacy DOC before use.
---

# Word Format Master

Use the current AI conversation and local CLI. AI interprets intent and evidence;
registered code reads, validates, writes and verifies DOCX. Never create independent
document-editing scripts. Editing project code is separate from formatting documents.

## Resolve paths and scope

Set `SKILL_DIR` to this directory. Resolve `PYTHON_BIN` to an available Python
executable after checking `import docx`. Use absolute paths for scripts, source
files, output, profiles, plans and reports. Keep intermediate files in a task folder.
Never overwrite the target manuscript or formatting source.

Identify target, output, scope and authority from the user's message and files.
Continue when clear; ask only about missing information that changes the result,
conflicting evidence or ambiguous targets. Do not ask for routine permission to
carry out a clearly requested format edit.

Choose the appropriate direct route:

1. **Requirement material**: extract from TXT/MD, DOCX/DOTX or text PDF.
2. **Formatted sample**: extract effective supported formatting from DOCX/DOTX.
3. **Local change**: inspect and translate the instruction into registered actions.
4. **Audit only**: check expected values without applying or saving a DOCX.

Do not combine formatting with prose rewriting unless requested. For a new thesis
draft, read [references/thesis-generation.md](references/thesis-generation.md).
When drafting from a sample, the sample defines the complete structure: include
every part it has, in its order, and add no absent parts. Include a cover only
when present; never automatically add an English abstract, English keywords,
integrity pledge, declaration, TOC, acknowledgements or appendix. Preserve the
sample's part counts, abstract languages and chapter hierarchy. Replace personal
metadata and research text with the new thesis's content, not the sample's.

## Extract evidence

Read [references/format-profile.md](references/format-profile.md) for profiles,
mappings and resolutions; [references/direct-format-api.md](references/direct-format-api.md)
for operations and selectors.

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" capabilities --output capabilities.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" profile requirements REQUIREMENTS.docx --output profile.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" profile template SAMPLE.docx --output profile.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" inspect INPUT.docx --output inspect.json
```

Run one extraction command matching the source. Save pasted requirements verbatim
as UTF-8 text before extraction. The parser recognizes common clauses; it is not
a complete natural-language interpreter. Read complete source blocks and every
issue. Translate unparsed semantics into explicit replacement rules with evidence
and reasons. Do not ignore unmatched wording when positive tokens were recognized.

Sample profiles record effective role properties, direct XML, styles, all page
sections, fields and feature inventory. Mixed values stay unresolved. Do not use
majority values or presets to fill missing properties. Review role detection
against text and inspect IDs; bold text alone does not establish a heading.
Tables, numbering, headers, footers and special structures may require explicit
operations. Never copy sample-specific names, titles, body text or headers.

Legacy DOC requires conversion to a separate DOCX. Scanned PDF needs OCR with
page evidence. Strict OOXML and macro formats are unsupported for writing.

## Resolve authority and mappings

Priority: explicit user choice > supplied institution/publisher template > written
requirements > named standard > user-selected preset > existing manuscript values.
See [references/standards-registry.md](references/standards-registry.md). Preserve
conflicting evidence and record the choice. Do not add presets to partial requests.

Record `choose`, `replace` or `ignore` decisions with reasons in resolutions JSON.
Use available evidence and prior user instructions; ask when evidence cannot decide.
An ignore reason must explain preservation or non-applicability, not hide unsupported
work. Retain ignored items in the delivery's coverage summary.

For multi-section samples, explicitly map every sample section in `mapping.sections`;
chapter numbers are not section numbers. Map rule IDs to current target selectors
in `mapping.rules` for unmatched or ambiguous roles. Use `mapping.preserve` with
reasons for target paragraphs outside sample coverage. Review all preserved rules
and paragraphs. For formatting an existing manuscript, preserve its content and
do not invent, reorder or delete modules to make the sample fit. For a new draft,
author exactly the sample's modules first; preservation mappings cannot excuse
missing or extra parts. Review `snapshot.structure_contract` against the complete
sample, including table/image covers and rendered pages. Resolve ambiguous
structures before drafting; do not fill them from a generic thesis outline.

## Plan, apply and verify

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" plan INPUT.docx OUTPUT.docx --profile profile.json --mapping mapping.json --resolutions resolutions.json --output plan.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" apply --plan plan.json --report report.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" verify --plan plan.json --report report.json --output verification.json
```

Omit optional mapping/resolution files when unnecessary. A result with
`kind: word-format-profile-review` and `status: needs_clarification` is a review,
not an executable plan. Resolve and plan again. No DOCX is written until apply.
Source changes require fresh extraction/planning. Never patch plans or reports
to force acceptance.

For every sample-based thesis draft, run `check-structure DRAFT.docx --profile
profile.json --output structure.json` before formatting and add `--match-structure`
to the `plan` command above. These commands compare part kinds, languages,
occurrences, order and chapter hierarchy against the actual sample. Fix the draft
when they differ; do not bypass the check with `ignore` or `mapping.preserve`.
The bound structure is checked again during apply and verify. It checks detected
structure, so page review must also confirm cover layout and any unrecognized
content. Only a user's explicit structural override authorizes a different outline.

For local changes, create a schema 2.0 request and use `--request request.json`
instead of `--profile`. Use inspect IDs, exact chapter titles or unique quotes.
Do not use a full-document profile for a local edit. Shared styles require an
explicit style target; unrequested properties and text stay unchanged.

For audit-only tasks:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" audit INPUT.docx --request request.json --output audit.json
```

For Format Spec 1.0, read [references/format-spec.md](references/format-spec.md).
It expands through the same registry. Official publisher packages remain the
authority where numeric rules cannot represent their layout or submission format.

Read [references/header-requirements.md](references/header-requirements.md) for
section/variant headers and STYLEREF. Header content, numbering, citations, TOC
and section insertion are explicit structural operations. Insert boundaries in
a separate plan, then inspect the new DOCX. Use real fields for supported TOCs
and page numbers. Do not claim unregistered SEQ/REF creation is supported.

## Render and deliver

Structural verification is mandatory. Use `render_docx.py INPUT.docx --output-dir
RENDER_DIR --method auto` for visual evidence; choose `word` or `libreoffice` when
specified. Explicit choices must not fall back silently. Inspect every page for
clipping, missing glyphs, broken tables, page breaks, headers and blank pages.
Repair the owning rules/request and rerender if needed.

The engine cannot evaluate Word fields; the renderer must refresh them. If
rendering is unavailable, report visual `skipped` with a reason. If pages await
inspection, use `pending`. Neither is a visual pass; never fabricate a report.

After page inspection, create the bound visual JSON described in
[references/direct-format-api.md](references/direct-format-api.md), then run:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" verify --plan plan.json --report report.json --visual-report visual.json --render-manifest RENDER_DIR/render-manifest.json --output verification.json
```

Hash the actual output and render manifest; never reuse an old review after
rerendering. If the renderer was explicitly selected, retain that selection in
the plan or pass `--render-method`. The checker verifies every page file hash.

Deliver the new DOCX and a concise summary of authority, applied/verified
properties, preserved/ignored items, unsupported work and visual status. Claim
full compliance only when every required property is represented and verified,
including visual checks. Do not promise arbitrary Word formats all migrate.
Keep internal JSON and page images out of delivery unless useful or requested.

For batches, inspect and plan each manuscript separately. Reusing a profile does
not authorize reusing source-bound IDs.

Example requests: “按这份学校要求整理论文并另存新文件”；“读取排版样例，应用支持的
格式并列出未覆盖项”；“只改第三章正文”；“Apply this sample's formatting to each
manuscript; preserve their content.”

When extending implementation, read [references/architecture.md](references/architecture.md).
