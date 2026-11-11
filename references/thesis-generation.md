# Thesis drafting and regression checks

Read this when a user explicitly requests a thesis draft, or when diagnosing
formatting failures from a generated thesis. Existing-document formatting still
uses the direct profile/request, plan, apply and verify workflow.

New reports are also routed here for content and citation checks. For “帮我写一份报告”
or a course/subject report without a supplied template or explicit format basis,
first follow [the default report workflow](report-generation.md):
`format_cli.py report prepare TASK_DIR` selects the approved source from
`templates/reports.json`, clones its package and extracts a source-bound profile.
Author the draft and use `--profile --match-structure`, never the thesis preset as
a fallback. User-supplied templates and explicit requirements take priority;
existing-document formatting and thesis requests retain their existing routes.

## Authoring boundary

Clarify consequential missing requirements: academic level, topic, length,
institution template, and available research evidence. A generic preset supplies
formatting parameters; it does not establish a dissertation's scientific merit.

Use supplied research data and verified publications. If a demonstration needs
illustrative numbers, identify them as simulated in the draft and table captions.
Do not present invented experiments, hardware runs, identities, supervisors or
unverified bibliography entries as real. When evidence is missing, use explicit
placeholders or a proposed experimental design rather than claims of completion.
Do not repair citations by editing fixed line numbers in a generated script. Use
`scripts/analyze_citations.py` to export paragraph fingerprints and stable reference
IDs, review placements, then let the citation workflow reject stale anchors or
unknown reference IDs before inserting markers.

Set `display` on every citation placement using the explicit requirements or the
sample's `citation.display.set` rule for the corresponding body scope. Use
`superscript` for raised citations and `inline` for citations on the text baseline.
If the sample uses both in the same scope, check the examples and confirm the
placement convention before choosing a display for new citations. Keep punctuation
placement in `insert`. Omitting `display` defaults to `inline`.

## Sample-based structure: nothing added or omitted

When the user supplies a thesis sample, its actual contents define the complete
draft structure. Extract a template profile before writing and review
`snapshot.structure_contract.ordered_modules` against the entire sample. Record
each part, language, occurrence, sequence and chapter hierarchy. The sample wins
over a preset's conventional outline, including declarations placed at the end.

- A sample with a cover requires a cover using its layout and metadata fields.
  When drafting or migrating from a DOCX/DOTX/DOC/DOT/WPS sample cover that carries
  school emblems, banners, metadata tables (`w:tbl` with borders/`vMerge`), VML
  textboxes, underlines or fixed typography, clone the cover via
  `python scripts/format_cli.py cover TEMPLATE OUTPUT.docx [--into DRAFT.docx] [--replace-existing-cover] --fields fields.json`
  (or `word_format.cover.clone_template_cover`) rather than replacing it with plain text,
  and fill in the new thesis's title, author and other supplied details; missing
  personal information stays a placeholder. Do not copy the sample's identity.
- A sample without a cover must not acquire one. A title inside an article is
  not permission to invent a separate cover page; inspect the page layout.
  The contract distinguishes an inline `title_block` from a `cover` and retains
  detected page/section boundaries in `separate_page`. Automatically flowed
  pages and unusual artwork still need visual confirmation.
- Chinese and English abstracts and keywords are separate parts. Generate only
  the languages and occurrences actually present. Do not translate an abstract
  to add a language the sample lacks.
- Include integrity pledges, originality declarations and authorization pages
  only when present, preserving their distinct types and counts. Do not fabricate
  a signature or claim the author has signed a declaration.
- Apply the same rule to TOC, symbols, references, acknowledgements, appendices,
  chapter count and subsection hierarchy. Never add a conventional “complete
  thesis” module merely because a generic preset includes it.

Use the sample's chapter/section structure with titles and content appropriate
to the new topic. An unknown heading, a table/image cover, a text box or other
ambiguous part requires reading its actual content and rendered page. Do not
treat a detector's missed marker as proof of absence or bypass a mismatch with
an ignore/preserve decision. If the evidence cannot establish the structure,
resolve that specific ambiguity before generating the affected content.

The formatting engine does not author a missing cover or research text. AI must
write the requested draft with exactly the required parts before formatting;
formatting must not silently delete existing manuscript content to make it fit.

```powershell
python scripts/format_cli.py profile template "样例.docx" --output profile.json
# Author the new draft from the reviewed sample structure and supplied research.
python scripts/format_cli.py check-structure "草稿.docx" --profile profile.json --output structure.json
python scripts/format_cli.py plan "草稿.docx" "论文.docx" --profile profile.json --match-structure --mapping mapping.json --resolutions resolutions.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json --output verification.json
```

Omit mapping/resolution files when unnecessary. `check-structure` is read-only;
it returns exit code 2 with exact differences for missing, extra, reordered,
duplicated or wrong-language parts and changed chapter hierarchy. The generation
plan must use `--match-structure`; a mismatch yields a review rather than an
executable plan. Apply and verify repeat the bound check. This mode follows the
sample's order, not a universal module ordering. Cover layout and prose quality
still require page/content review. Only explicit user instructions can change
the sample's structure; retain those instructions and the resulting outline.

## Semantic authoring

### Mandatory dynamic numbering

Never type structural numbers into generated content. Write only heading titles
and bind chapter/section styles to native multilevel Word numbering. Ordered lists
and bibliography entries require native numbering definitions; captions and
equation labels require SEQ fields (with STYLEREF for chapter prefixes), page
numbers require PAGE/NUMPAGES, and citations/cross-references require bookmarks
and REF/PAGEREF fields. Preserve visible formats through numbering definitions
and field formatting, never literal prefixes or duplicate typed labels. Ordinary
data, years, measurements and identifiers remain content.

This requirement also applies when the sample uses typed numbers: infer its
display scheme and implement it dynamically. Resolve an unknown scheme before
delivery rather than falling back to typed numbers. Do not treat the caption
`numbering_mode: seq` configuration alone as proof that a SEQ field was inserted;
the current caption operation only controls position/formatting. Author the
actual fields in the draft and inspect them after formatting.

Run `python scripts/check_numbering.py OUTPUT.docx --output numbering.json`
after plan/apply/verify; a failure blocks delivery. The shared pipeline runs this
guard automatically. It checks common typed labels in body/table/header/footer
paragraphs and excludes real field results. Also review equations, custom
enumerations, appendix labels and cross-references explicitly: the conservative
detector cannot establish all numbering intent from arbitrary prose. Confirm
that inserting/removing an item and updating fields renumbers dependent labels
in Word/WPS. Report structural and visual validation separately.

Build a new draft with semantic styles: Title for a cover title when present, Heading 1–3
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
An authoring script's `doc.save()` creates a draft, never an accepted final
document. Inspecting or rendering that draft does not replace the plan, apply
report and independent verification bound to the delivered output's hash.
Before applying, record effective font channels, size, bold and paragraph
properties separately for the cover title, each metadata field/date, each
abstract title/label/body and keyword label/value in each language. Read style
inheritance; do not guess a heading font from its name or translate “加黑” to
the 黑体 font family. Empty placeholders require interpretation against the
written requirements, and separate-document specifications need a scope decision.
See [format-profile.md](format-profile.md) for semantic selectors and coverage.
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

## Single-Pass Authoring & Acceptance Playbook

- **Multi-authority extraction & conflict check**: When a task folder contains both a written requirements document (e.g. `考核要求.docx`) and a formatting sample (e.g. `格式范文.doc`), you MUST extract BOTH profiles (`format_cli.py profile requirements` and `format_cli.py profile template`) and compile or compare them (`format_cli.py template compile --profile req.json --profile tpl.json`). Never ignore the written requirements and never silently mix conflicting fonts, alignments, headers, or section page numbers.
- **No custom formatter bypass**: Authoring scripts (`python-docx`) may ONLY create the raw semantic draft (`draft.docx`). Never import `word_format.*` internal modules in a scratch script to format or deliver a final document without running `run_format_pipeline.py` (or `format_cli.py plan → apply → verify`).
- **Cover cloning & spacer compaction**: When a long thesis title wraps to 2 lines on a cloned sample cover, pass `--compact-spacers <N>` to `python scripts/format_cli.py cover` so pure-empty spacer paragraphs are compacted safely without ever deleting empty-text paragraphs that contain `<w:drawing>` / `<w:pict>` school logos or calligraphy banners. `format_cli.py cover` also strips red format-hint runs (such as `（小二号居中`) and resets compensating `firstLine` indentation on centered cover dates. Ensure the cover title and the abstract page title match character-for-character.
- **Resilient citation placement & 100% reference coverage**: `citation.insert` placements accept `"text_contains"` (in addition to or instead of `"paragraph_index"` and `"paragraph_sha256"`) and automatically re-resolve shifted paragraph indices when earlier chapters are expanded. Every listed bibliography entry MUST be cited in the body (`--require-all-references-cited`), and all Chinese/English references must be real, verifiable publications (never fabricated).
- **Built-in field refresh & per-level TOC sync**: Run `python scripts/format_cli.py refresh-fields <draft.docx> --in-place` before final formatting/rendering to update TOC, `REF`, `SEQ`, and `STYLEREF` fields via Word COM, automatically synchronize `toc 1..9` style fonts onto all TOC runs (including `<w:tab/>` dot-leader runs with `<w:rPr>` enforced at child index `0`), and strip cached `<w:lastRenderedPageBreak/>` markers.
- **Flexible chapter counts & multi-table rules**: When the user asks to expand a sample-based draft to more chapters than the sample's chapter count while keeping front/back matter identical, pass `--flexible-chapters` to `format_cli.py check-structure`, `format_cli.py plan`, or `run_format_pipeline.py`. Sample body table rules (`table.alignment.set`, `table.header_repeat.set`, `table.three_line.apply`) automatically target all body tables (`all: True`).
- **Pipeline content quality & visual sign-off gate**: Pass `--check-quality` and the task's quantitative bounds (`--min-chinese-characters`, `--min-abstract-chinese-characters`, `--max-abstract-chinese-characters`, `--min-references`, `--min-recent-references`, `--recent-start-year`, `--require-all-references-cited`) to `run_format_pipeline.py`. After `--render`, inspect every `page-1.png..page-N.png` via your image-reading tool, verify all 6 `checklist` items in `visual-report.template.json`, change `"status"` from `"pending_visual_review"` to `"passed"`, write a concrete `"summary"`, and run `run_format_pipeline.py ... --visual-report ".../visual-report.template.json" --resume`.
