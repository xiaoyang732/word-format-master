# Word/DOCX Project Research

## Purpose

Use this matrix when extending the implementation. It records which ideas are adopted without copying third-party code. Recheck the upstream license and current behavior before vendoring any dependency or template.

## Research Method

- Prefer official documentation, official repositories, published specifications, and maintained package documentation.
- Treat popularity as a discovery signal, not proof of fitness.
- Record architecture patterns separately from reusable code and template licensing.
- Keep this project dependency-light: the dashboard and analyzers run with Python's standard library; `python-docx` and `lxml` are used only for DOCX application.

## Adopted Project Matrix

| Project | Category | Useful idea | Adopted here | Boundary |
|---|---|---|---|---|
| [OpenAI Build skills](https://learn.chatgpt.com/docs/build-skills) | Agent workflow | Progressive disclosure, concise trigger metadata, scripts and references | Lean `SKILL.md`; detailed research and schemas in `references/`; deterministic operations in `scripts/` | Do not turn the skill into a general application framework |
| [openai/skills](https://github.com/openai/skills) document tooling | Agent workflow | Preset tokens, template distillation, structural audits, render-and-inspect gate | Evidence-backed format spec, explicit numeric values, non-destructive edits, render gate | Do not claim render success without page inspection |
| [ppt-master](https://github.com/hugohe3/ppt-master) | Routed skill | Thin entry point, one active workflow, user gates, stable paths | Exactly one route per request; template-first route; local interactive dashboard | Do not copy presentation-specific design or scripts |
| [python-docx](https://github.com/python-openxml/python-docx) | DOCX API | High-level paragraphs, runs, styles, sections, headers, footers, tables | Primary safe application layer in `apply_spec.py` | Use direct OOXML for unsupported fields; do not expect field evaluation |
| [python-docx-template](https://github.com/elapouya/python-docx-template) | Template filling | Preserve a designer-authored DOCX while replacing controlled slots | Future slot-map route; template remains the visual authority | Not a general normalization engine |
| [docxtemplater](https://github.com/open-xml-templating/docxtemplater) | Template filling | Structured tags, modules, browser/Node workflows | Adopt the separation between content data and document template | Core and commercial modules have different capabilities/licenses |
| [docx](https://github.com/dolanmiu/docx) | DOCX generation | Typed, declarative document model in JavaScript | Format specification uses declarative JSON tokens | Do not add a second generation engine to the MVP |
| [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) | OOXML model | Schema-aware validation and package-level editing | Adopt package inventory, relationship checks, and explicit unsupported-feature reporting | .NET validator is optional, not bundled |
| [docx4j](https://github.com/plutext/docx4j) | OOXML model | JAXB object model and deep numbering/field support | Use as a reference for advanced OOXML coverage | Java runtime is outside MVP |
| [Pandoc](https://github.com/jgm/pandoc) | Conversion | Reference DOCX, semantic content conversion, citation filters | Future import/export route; keep content conversion separate from formatting normalization | Round-tripping loses some Word-specific constructs |
| [Mammoth](https://github.com/mwilliamson/python-mammoth) | DOCX to HTML | Semantic rather than visual conversion | Useful for text/structure extraction and accessibility preview | Not suitable for layout-faithful preview |
| [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) | Browser preview | Fast client-side DOCX approximation | Architecture supports adding an advisory preview layer | Never use as the final pagination authority |
| [LibreOffice](https://git.libreoffice.org/core) | Rendering | Headless conversion and PDF output on the supported Windows runtime | Selectable renderer and automatic fallback; Dashboard can provision an official Windows runtime into the ignored local `.runtime/` directory | Layout can differ from Microsoft Word; an explicit Word choice must not fall back silently |
| [unoserver](https://github.com/unoconv/unoserver) | Conversion service | Reuse one LibreOffice process for repeated conversion | Candidate for batch/server deployments | Unnecessary for the single-user MVP |
| [ONLYOFFICE Document Server](https://github.com/ONLYOFFICE/DocumentServer) | Editing/rendering | Full web editor and collaborative rendering | Reference for a future hosted edition | Too heavy and license-sensitive for a local skill MVP |

## Additional Projects Audited

| Project | Category | Useful idea | Decision in this project |
|---|---|---|---|
| [docxcompose](https://github.com/4teamwork/docxcompose) | DOCX composition | Merge packages while respecting styles, relationships, numbering, headers, and footers | Keep composition separate from normalization; never simulate a merge by copying paragraph text |
| [docx2python](https://github.com/ShayHill/docx2python) | Structured extraction | Traverse body, tables, headers, footers, footnotes, and nested content instead of relying only on top-level paragraphs | `analyze_docx.py` inventories the package and scans content-bearing parts; `apply_spec.py` walks nested table and header/footer paragraphs |
| [MarkItDown](https://github.com/microsoft/markitdown) | Semantic conversion | Produce LLM-friendly semantic text from Office files | Candidate input route for content understanding; excluded from visual-format inference because semantic conversion discards layout evidence |
| [Docling](https://github.com/docling-project/docling) | Document understanding | Rich structural extraction and normalized document models | Candidate for complex audit reports; not required by the dependency-light local MVP |
| [Quarto](https://github.com/quarto-dev/quarto-cli) | Publishing | Use a reference DOCX as a repeatable style authority | Reinforces template-first precedence and the explicit split between semantic content and Word styling |
| [docx-templates](https://github.com/guigrpa/docx-templates) | Template automation | Validate command/data boundaries and preserve a designer-owned template | Adopt strict separation between user content, format tokens, and template package; do not execute expressions from uploaded files |
| [Carbone](https://github.com/carboneio/carbone) | Report templating | Treat templates as immutable inputs and render to new outputs | Source files are never overwritten; dashboard applications always produce a new DOCX |
| [docx2pdf](https://github.com/AlJohri/docx2pdf) | Native rendering | Use Microsoft Word automation where installed for highest-fidelity PDF conversion | Word rendering is the preferred final authority; LibreOffice is the documented fallback |
| [OpenTBS](https://github.com/Skrol29/opentbs) | ZIP/XML templating | Make narrow package-level edits without rebuilding the whole document | Use direct OOXML only for evidence or unsupported constructs; keep high-level edits in `python-docx` |
| [UniOffice](https://github.com/unidoc/unioffice) | Typed OOXML API | Strongly typed Word generation and manipulation | Reference for future service implementations; avoid adding a second runtime to the MVP |

## Thesis Skill Comparison (2026-10-10)

Reviewed the named upstream repositories and two related projects at their then-current `main` snapshots. This matrix records ideas, not copied implementations; details and review findings are in [review-2026-10-10.md](review-2026-10-10.md).

| Project | Useful evidence | Adopted now | Boundary |
|---|---|---|---|
| [the-shy123456/thesis-docx](https://github.com/the-shy123456/thesis-docx) | `failure-patterns-and-quality-gates.md` checks effective styles, hidden indentation and field drift; Word PDF helper refreshes story fields | Clarify quality gates; refresh fields across Word story ranges before PDF export | A rendered WPS page is not proof of Microsoft Word or university-template parity |
| [WEN-JY/academic-research-skills](https://github.com/WEN-JY/academic-research-skills) | `docx-thesis-format` separates explicit institutional JSON rules from check/fix scripts, including allowed margin profiles | Add read-only audit to the existing planner/reader rather than copy one school's numeric profile | Zhejiang University values are not universal thesis requirements |
| [Jonnys-Li/software-thesis-docx-skill](https://github.com/Jonnys-Li/software-thesis-docx-skill) | Manifest owns source order, figures and formatting mode; template extractor captures semantic roles | Keep template precedence, isolate dashboard workflow from the skill entrypoint | Repo-to-thesis content generation is a different route from formatting an existing DOCX |
| [zxyasfas/paper_format_agent](https://github.com/zxyasfas/paper_format_agent) | Content fingerprint and synthetic text-survival benchmark report explicit coverage gaps | Strengthen report completeness and render-source fingerprinting; retain per-operation package-preservation checks | A body-text fingerprint alone cannot prove preservation of fields, footnotes or drawings |
| [zouchenzhen/docx-template-translator-skill](https://github.com/zouchenzhen/docx-template-translator-skill) | Separates Pandoc content conversion, template adaptation and Word finalization | Document a possible future conversion route distinct from deterministic formatting | Template-specific generated Python should not bypass this project's registered operation boundary |

## Word Field and Reference Notes

The user supplied two practical Word articles as secondary operational references:

- [论文写作中怎样正确插入参考文献，引用文献如何标注？](https://zhuanlan.zhihu.com/p/309606954): use automatic list numbering and cross-references instead of hand-typed citation numbers.
- [参考文献著录格式（国标 GB/T 7714-2015）](https://zhuanlan.zhihu.com/p/355312827): use hanging indents, repeat citations through cross-references, and refresh fields after edits.

Adopt only the Word mechanics. Do not use either article to override publisher, venue, institution, or current GB/T rules. Publisher-specific citation and reference formatting comes from the authoritative template; GB/T bibliographic formatting remains a separate Chinese-paper mode.

## Resulting Architecture

The project combines five proven patterns:

1. Route one task at a time, following `ppt-master` and OpenAI skill design.
2. Represent format as an explicit, inspectable JSON specification.
3. Use `python-docx` for common edits and direct OOXML inspection for evidence.
4. Use a browser only for configuration; execute the selected Word or LibreOffice renderer through `scripts/render_docx.py` and treat its page PNGs as the visual evidence.
5. Preserve the original and report unsupported constructs instead of silently flattening them.

## Adoption Traceability

| Borrowed pattern | Concrete implementation |
|---|---|
| Thin routed skill and progressive disclosure | `SKILL.md`, with details kept in `references/` and deterministic work in `scripts/` |
| Declarative document model | `assets/presets.json` and `references/format-spec.md` |
| Reference-DOCX/template-first precedence | Authority rules in `SKILL.md` and `references/standards-registry.md` |
| Package-aware inspection | `scripts/analyze_docx.py`, including ZIP limits, styles, inheritance, sections, nested content parts, and complex-object warnings |
| Conservative style application | `scripts/apply_spec.py`, named styles, explicit page geometry, opt-in direct-format cleanup, and no in-place writes |
| Semantic requirement extraction | `scripts/parse_requirements.py`, with field-level evidence, confidence, and unresolved text |
| Local editable control surface | `scripts/serve_dashboard.py` and `assets/ui/` |
| Native-render quality gate | `scripts/render_docx.py` executes Word COM export or LibreOffice headless export, then rasterizes every PDF page; browser preview is explicitly advisory |

## Explicit Non-Goals

- Do not rebuild Microsoft Word in the browser.
- Do not treat HTML preview as layout proof.
- Do not copy or redistribute official publisher templates without permission.
- Do not flatten fields, comments, tracked changes, equations, controls, or macros merely to simplify implementation.
- Do not rank projects by an unverified live star count.
