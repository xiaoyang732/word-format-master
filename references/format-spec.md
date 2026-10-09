# Format Specification

## Contract

Use JSON as the single source of truth shared by template analysis, requirement parsing, the dashboard, and DOCX application. Store measurements as millimeters or points at the boundary and convert to Word units only inside the application engine.

Every recognized field may carry provenance:

```json
{
  "value": 25.4,
  "source": "APA 7 preset",
  "confidence": 1.0,
  "locked": false,
  "evidence": "Margins are 1 inch on all sides"
}
```

The MVP dashboard edits plain values and keeps provenance in its analysis panel. Future versions may store provenance inline for every token.

## Top-Level Shape

```json
{
  "schema_version": "1.0",
  "id": "custom-format",
  "name": "Custom format",
  "mode": "parameterized",
  "template_required": false,
  "page": {},
  "body": {},
  "lists": {},
  "headings": [],
  "captions": {},
  "references": {},
  "document_structure": {},
  "citations": {},
  "tables": {},
  "headers_footers": {},
  "page_numbers": {},
  "table_of_contents": {},
  "authority": {},
  "notes": [],
  "unsupported": []
}
```

## Official Template Workflows

Store publisher-controlled workflows under `template_workflows` in `assets/presets.json`, never as numeric presets. A workflow contains `mode: "official-template"`, `template_required: true`, an HTTPS `official_url`, publisher authority, and ordered workflow steps. It must not contain synthetic `page`, `body`, or `headings` values.

An analyzed publisher template may produce an inferred format specification for inspection and comparison. Keep the workflow marker and reject direct application: inferred tokens cannot reproduce the template package, semantic styles, numbering, section/column behavior, or downstream publishing system.

An official workflow may expose local copies supplied by the user under `local_templates`. Record the exact SHA-256 and treat them as local personal assets; the current publisher instructions still decide which variant is valid. Prefer an explicit path relative to `assets/templates`; filename-only entries are accepted only when exactly one file matches.

## Supported MVP Tokens

### Document Structure

`document_structure` is the template's ordered module contract. It is extracted
from the body paragraph sequence and is kept empty when a module is absent; no
default caption, reference, or numbering rule is invented during distillation.

```json
{
  "schema_version": "1.0",
  "strict_order": true,
  "ordered_modules": [
    {
      "id": "cover",
      "kind": "cover",
      "repeatable": false,
      "paragraph_start": 0,
      "paragraph_end": 4,
      "style_roles": {
        "title": {"tokens": {}},
        "body": {"tokens": {}}
      }
    },
    {"id": "abstract", "kind": "abstract", "repeatable": false},
    {"id": "toc", "kind": "toc", "repeatable": false},
    {"id": "chapters", "kind": "chapters", "repeatable": true},
    {"id": "references", "kind": "references", "repeatable": false},
    {"id": "appendix", "kind": "appendix", "repeatable": false}
  ],
  "order_errors": []
}
```

The shared module detector recognizes cover, declaration, abstract, keywords, TOC, symbols,
chapters, references, acknowledgements, and appendix. During application, the
module's `style_roles` are assigned through named Word styles. Structural
verification fails when a required module is missing or appears out of order.

### Tables

- `three_line_table`: boolean (applies standard academic three-line table: 1.5pt top/bottom borders, 0.75pt header bottom border, no vertical borders)
- `repeat_header_row`: boolean (sets `w:tblHeader` and `w:cantSplit` on row 0)
- `alignment`: `left`, `center`, or `right`
- `font_size_pt`, `font_east_asia`, `font_latin`: table cell typography

### Table of Contents

- `size`: `A4`, `Letter`, or `Custom`
- `width_mm`, `height_mm`
- `orientation`: `portrait` or `landscape`
- `margin_top_mm`, `margin_right_mm`, `margin_bottom_mm`, `margin_left_mm`
- `gutter_mm`, `header_distance_mm`, `footer_distance_mm`

### Body

- `font_latin`, `font_east_asia`
- `font_size_pt`, `bold`, `italic`, `color`
- `alignment`: `left`, `center`, `right`, or `justify`
- `line_spacing`: `{ "kind": "multiple", "value": 1.5 }` or `{ "kind": "exact", "value_pt": 20 }`
- `space_before_pt`, `space_after_pt`
- `first_line_indent_mm`, `left_indent_mm`, `right_indent_mm`
- `keep_together`, `widow_control`

### Headings

Each heading contains `level`, typography tokens, paragraph tokens, `keep_with_next`, and optional `page_break_before`. Use Word Heading styles rather than direct formatting.

### Numbered Lists & Heading Numbering

Store heading numbering format and small-point numbering under `lists`:

- `lists.headings.level1`: `chinese` (第一章、第二章...), `arabic` (1、2...), or `keep` (preserve original text)
- `lists.headings.level2`: `arabic` (1.1, 1.2...) or `keep`
- `lists.headings.level3`: `arabic` (1.1.1, 1.1.2...) or `keep`
- `lists.numbered.style`: `decimal-period`, `decimal-parenthesis`, `decimal-parentheses`, `decimal-fullwidth-parentheses`, `chinese-period`, `chinese-parentheses`, `upper-alpha-period`, `lower-alpha-parenthesis`, or `upper-roman-period`
- `lists.numbered.start`: integer greater than or equal to 1
- `lists.numbered.left_indent_mm`, `hanging_indent_mm`

Create a real `w:abstractNum`/`w:num` definition and attach `w:numPr` to existing numeric list paragraphs. Restart each contiguous list group. Do not convert bullet lists.

### References

- `citation_system`
- `font_*`, `font_size_pt`, `line_spacing`
- `hanging_indent_mm`, `space_after_pt`
- `citation_mode`: `numeric`, `author-date`, or `author-page`
- `numbering_mode`: `word-numbering` or `none`

Treat `citation_system`, `citation_mode`, and `numbering_mode` as provenance and workflow metadata unless a dedicated citation converter is present. The dashboard displays the citation system read-only; it does not claim to rewrite bibliographic content. The application engine updates the `Bibliography` style and assigns it to paragraphs inside an identifiable References/参考文献 section.

For citation placement, export body paragraphs and existing bibliography entries with `scripts/analyze_citations.py`. Let AI return only `citations.placements`, including `paragraph_index`, `paragraph_sha256`, `reference_ids`, `citation_text`, `insert`, `spacing`, `confidence`, and `reason`. `spacing` is one of `none`, `space-before`, `space-after`, or `spaces-around`. Accept only existing reference IDs. Recompute every paragraph fingerprint immediately before insertion and reject the whole batch on mismatch. AI chooses placement; deterministic code performs the DOCX edit.

### Figure and Table Captions

Store separate `captions.figure` and `captions.table` objects:

- `style_name`, `source_style_id`
- `label`, `position`: `above` or `below`
- typography and paragraph tokens from Body
- `numbering_mode`: `seq` for generated captions or an analyzed template's existing numbering definition

Applying a specification updates named caption styles and reassigns existing top-level captions whose text starts with the configured matching label. For an adjacent top-level table or drawing paragraph, it can move the caption above or below the object by repositioning the existing OOXML paragraph. It does not guess where separated or missing captions belong. Insert new captions with real `SEQ Figure` or `SEQ Table` fields and update fields in Word.

### Headers and Footers

See [header-requirements.md](header-requirements.md) for the optional article-backed rules and exact action/handler/readback mapping. Header spec fields expand to independent `header.*` operations. `section_number` restricts the header target; `sections` carries per-section overrides. `even_header` and `first_header` carry separate variants. `mode: chapter_title`, `style_name` and optional `prefix` generate a STYLEREF field; `mode: text` is literal text.

- `preserve_existing`
- `different_first_page`, `different_odd_even`
- `header` and `footer`: `enabled`, `text`, `alignment`, and typography tokens

Managed footer text uses `WFM Footer`; header content has an explicit managed marker and legacy `WFM Header` remains recognized. When `preserve_existing` is `true`, disabling global managed text clears only tool-managed paragraphs. Explicit content replacement or selected-section clearing replaces that variant. `preserve_existing: false` explicitly clears existing headers/footers before rebuilding template content; it is forbidden with a selected header section. Template extraction preserves separate sections and variants where present.

### Page Numbers

- `enabled`, `location`: `header` or `footer`
- `alignment`, `start`, `show_on_first_page`
- `format`: `number`, `page-number`, or `page-x-of-y`

Generate actual `PAGE` and optional `NUMPAGES` fields, keep cached display text, and set `w:updateFields=true`. Never write a fixed page number as ordinary text.
When the target header/footer already contains a standalone PAGE-field paragraph, reuse and rebuild that paragraph rather than appending a second page number. Structural verification requires exactly one PAGE field in each selected target part.
Disabling page numbers clears managed `WFM Page Number` paragraphs. For a distilled uploaded template, `headers_footers.preserve_existing: false` also removes pre-existing PAGE fields before any explicitly enabled template header, footer, or page number is rebuilt.

### Table of Contents

- `enabled`
- `title`
- `max_heading_level`: integer from 1 through 9
- `page_break_after`

Generate a real `TOC \\o "1-N" \\h \\z \\u` Word field rather than a static list. The engine creates a managed title, field, and optional trailing page-break paragraph, sets `w:updateFields=true`, and removes only those managed paragraphs when disabled. The field carries a cached heading-label snapshot for renderers that do not recalculate Word fields; Word refreshes the authoritative page numbers on open. A TOC includes only paragraphs using actual Word Heading styles; bold, enlarged body text is not a heading.

Treat an enabled TOC with zero non-empty Heading paragraphs in the selected level range as a structural verification failure. Assign the intended sections to real Heading styles and apply the specification again; do not accept an empty directory page.

## Dashboard Application Contract

Every editable Dashboard setting maps through the legacy adapter to executable operations in `scripts/word_format/registry.py`. `SUPPORTED_APPLICATION_PATHS` is generated by that adapter. `/api/capabilities` exposes legacy settings and `format_core.operations`, including handlers, validators and independent readers/verifiers. Keep citation-system identifiers and other non-applicable metadata read-only.

Direct AI requests use schema 2.0 and need no Dashboard or Handoff; see [direct-format-api.md](direct-format-api.md). The legacy spec remains 1.0 and expands into the same planner/executor. Supply one authoritative indent unit (chars or mm); inconsistent duplicates fail. Explicit zero and false remain meaningful. Font cleanup affects requested properties only. Numbering, citations and TOC changes require explicit structural operations.

## Merge Rules

- Replace a lower-authority value only when a higher-authority value is explicit.
- Keep unspecified fields inherited; do not replace them with zero or Word defaults.
- Exception: an uploaded DOCX/DOTX selected as the template is distilled as a complete snapshot, not a patch. Its extracted specification replaces the selected preset wholesale, and absent supported features are represented explicitly as empty or disabled instead of inheriting preset values.
- Record conflicts when two sources of equal authority disagree.
- Reject invalid units, negative margins, zero page dimensions, unsupported orientation, and fonts represented only by theme placeholders.

## Application Limits

The application engine applies page geometry, named paragraph styles, managed header/footer text, PAGE/NUMPAGES fields, and managed TOC fields. It does not rebuild publisher multilevel numbering, infer missing figure/table anchors, generate bibliography entries, evaluate fields, or rewrite equations, controls, comments, tracked changes, text boxes, and drawings. Preserve these parts and list them under `unsupported` when their behavior could change.

## Verification

After applying a specification:

1. Re-analyze the output and compare page/style tokens with the requested values.
2. Update Word fields using Microsoft Word when available.
3. Render every page.
4. Inspect every page at 100% zoom.
5. Repair the specification or owning style, not isolated visible symptoms.

Structural verification is mandatory after every application and must reopen the generated package and compare requested geometry, numbering, fields, and inserted citation text. Visual verification is user-selectable. When enabled, render every page and use only a model whose runtime declares image-input support. If the selected model does not support image input, record visual status as `skipped` with the reason and notify the user; never report a visual pass.
