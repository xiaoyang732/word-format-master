# Thesis draft handoff and regression checks

Read this when a user explicitly requests a thesis draft, or when diagnosing
formatting failures from a generated thesis. Existing-document formatting still
uses the normal direct or confirmed Dashboard workflow.

## Authoring boundary

Clarify consequential missing requirements: academic level, topic, length,
institution template, and available research evidence. A generic preset supplies
formatting parameters; it does not establish a dissertation's scientific merit.

Use supplied research data and verified publications. If a demonstration needs
illustrative numbers, identify them as simulated in the draft and table captions.
Do not present invented experiments, hardware runs, identities, supervisors or
unverified bibliography entries as real. When evidence is missing, use explicit
placeholders or a proposed experimental design rather than claims of completion.

Build a new draft with semantic styles: Title for the cover title, Heading 1–3
for actual chapters/sections, Normal for prose, and separate caption/reference
styles. Use paragraph/page/section break properties deliberately. Blank cover
paragraphs and page-break runs are legitimate content; do not remove them merely
to make preservation checks pass.

The `thesis-standard` preset explicitly makes Title/Subtitle and Heading 1–3
text black. Do not rely on the Office template defaults: Heading styles often
inherit accent blue, while Title can use a dark theme color. A black cached RGB
does not override an active theme color. Validate effective color after formatting,
including direct run overrides. Preserve a school template's explicit color when
that template is the chosen authority instead of imposing the generic preset.

Chapter introductions such as “第二章：相关理论。介绍技术基础。” are body prose.
Do not rewrite them to evade a classifier. The shared heading detector honors
explicit heading styles and conservatively infers short numbered titles from
unstyled text. Colon-led descriptions, complete sentences, decimal measurements,
list styles and cached TOC lines do not establish a new chapter. If an intended
title cannot be inferred, assign its Heading style explicitly before formatting.

## Formatting and verification

Apply the selected template/preset through `format_cli.py plan/apply/verify`.
Do not repair draft formatting with an independent editing script or patch an
installed skill mid-run. Fix the repository implementation, validate it, then
sync the changed skill files without replacing unrelated local modifications.

Large specifications group identical paragraph action/parameter pairs into
multi-target operations after module/global token merging. Every target still
receives its own check; property projection and source/output hashes remain
required. Fewer operations alone do not prove a correct plan: verify coverage,
requested values and the resulting document.

First-line and hanging indents are mutually exclusive. Write only the requested
side; an opposite `hanging="0"` can suppress first-line indentation in Word/WPS.
Character indents stay in character units, not a fixed conversion based on 12pt.
Inspect the effective value and confirm it in the target renderer, including
documents with inherited or stale direct indentation overrides.

Before delivery, compare the actual chapter outline with the intended outline,
check that each chapter has its intended content, refresh fields and inspect all
rendered pages for extra chapters, indent loss, stale TOC entries, blank-page
drift and broken tables. A structural pass is neither a visual pass nor a
validation of research claims. When reviewing a user-supplied test sample without
authorization to edit it, keep it read-only and reproduce defects with synthetic
fixtures.
