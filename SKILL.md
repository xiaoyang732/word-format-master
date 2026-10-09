---
name: word-format-master
description: Inspect and edit Word formatting through deterministic local code, for whole documents, chapters, paragraphs or text ranges. Use direct AI requests or user-configured dashboard handoffs through the same format core. Also analyze DOCX/DOTX templates and audit formatting. Legacy DOC must first be converted to DOCX.
---

# Word Format Master

Use a template-first, evidence-backed workflow. Keep source documents unchanged, represent formatting as explicit JSON tokens, and verify both package structure and rendered pages before delivery.

## Resolve Paths

Set `SKILL_DIR` to the directory containing this file. Use absolute paths for every input, output, temporary directory, script, and asset.

Before running any command shown below, resolve `PYTHON_BIN`. In AI Agent Desktop environments, resolve and set it to the available Python executable. Do not assume the environment variable already exists. Outside that runtime, use an absolute `python` executable only after `& "C:\absolute\path\python.exe" -c "import docx"` succeeds.

## Choose the Entrance

For a direct AI request, inspect the document and call `format_cli.py` without launching a browser. When the source, scope and parameters are clear, execute directly into a new DOCX. Ask only about ambiguous requirements or selectors. Do not add unspecified defaults or apply a full preset to a local edit.

For webpage configuration, launch the Dashboard and let the human edit and submit it. AI must not click confirmation, fill controls or fabricate a handoff on the human's behalf. The confirmed handoff uses the same planner, registry and executor as direct requests.

Existing handoffs remain bound to their exact source and output paths and source SHA-256. Direct requests do not require a handoff. Editing code in this project is separate from modifying user documents.

## Direct AI Workflow

Read [references/direct-format-api.md](references/direct-format-api.md) for request contracts and selectors. Use absolute paths:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" capabilities
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" inspect INPUT.docx --output inspect.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" plan INPUT.docx OUTPUT.docx --request request.json --output plan.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" apply --plan plan.json --report report.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" verify --plan plan.json --report report.json
```

Map the user's intent to registered operations, never generate ad hoc document-editing scripts. Use inspect IDs, exact chapter titles, filtered paragraph ordinals or unique quotes. A plan resolves selectors and detects conflicts before execution; source changes require a new plan. Atomic format edits preserve text, objects and unrelated properties. Heading/reference numbering conversion, citation insertion and TOC generation are separate explicit operations.

## Select One Route

Choose exactly one route:

1. **Distill template**: inspect a DOCX/DOTX and produce a reusable format specification.
2. **Interpret requirements**: convert pasted formatting instructions into supported tokens with evidence and confidence.
3. **Execute confirmed handoff**: apply only an existing, user-confirmed `HANDOFF.json` for its matching source DOCX.
4. **Audit only**: report structural and formatting drift without modifying the source.
5. **Configure interactively**: when the user chooses webpage configuration, preprocess the source document, launch the dashboard, let the human inspect or adjust tokens, then execute the handoff.
6. **Direct formatting**: translate clear instructions to registered operations and execute the direct workflow above.

Do not combine formatting with content rewriting unless the user explicitly requests both.

## Load References

- Read [references/format-spec.md](references/format-spec.md) before creating, merging, or applying a specification.
- Read [references/standards-registry.md](references/standards-registry.md) when selecting an academic preset or interpreting an authority hierarchy.
- Read [references/project-research.md](references/project-research.md) only when extending the implementation or choosing a new engine.
- Read [references/architecture.md](references/architecture.md) before changing API routes, session lifecycle, module boundaries, or data-contract versions.

## Analyze Inputs

For DOCX/DOTX input, run:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/analyze_docx.py" INPUT.docx --output analysis.json
```

Inspect Strict and Transitional OOXML, styles, inheritance, direct formatting, numbering, sections, headers, footers, tables, fields, and package parts. Do not infer the whole template from the first page or from `Document.paragraphs` alone.

Treat an uploaded DOCX/DOTX used as a formatting template as a complete formatting snapshot. Populate every supported Dashboard setting from that template's final effective values, resolving direct formatting, style inheritance, document defaults, and theme fonts. Use explicit `false`/empty values only for genuinely absent features such as headers, footers, page numbers, and table of contents. Never deep-merge a distilled template over a built-in preset or use the preset to fill fields that the uploaded template does not define.

For thesis-style documents, also extract the ordered module contract from the
body sequence: cover, declaration, abstract, keywords, table of contents, symbols, chapters,
references, acknowledgements, and appendix. Each present module carries its
paragraph anchors and role-specific style tokens; absent modules remain absent.
The same detector is used during application and structural verification, so a
missing or out-of-order module fails acceptance instead of being silently
reclassified.

For pasted requirements, run:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/parse_requirements.py" requirements.txt --output detected.json
```

Keep unmatched requirements unresolved. Never invent missing margins, fonts, numbering, or citation rules.

Legacy `.doc` files are not accepted by the Dashboard. If institution material is supplied as `.doc`, convert a read-only copy to `.docx` with Microsoft Word for analysis, record the conversion as evidence, and use a newer user-supplied `.docx` mother template when one is available.

## Resolve Authority

Merge values in this order:

1. User-locked field
2. Uploaded institution, publisher, or organization template
3. Explicit institution or publisher instructions
4. Named citation or academic standard
5. Selected built-in preset
6. Existing document value

Preserve each field's source, confidence, and override reason. Only apply a preset when the user requests whole-document normalization against that preset. Existing values remain the authority for unrequested properties in local edits.

When the user selects **标准学位论文/毕业论文规范模板** (`thesis-standard`) for whole-document normalization, its rules are:
- **页面边距**：标准 A4 纵向，页边距上 25.4mm、下 25.4mm、左 25.0mm、右 25.0mm，装订线 0mm，页眉/页脚距边界 15.0mm。
- **正文排版**：中文小四号宋体（12pt），西文小四号 Times New Roman，1.5 倍行距，段前段后 0 磅，首行缩进 2 字符（7.4mm），两端对齐。
- **标题层级**：一级标题（章）三号（16pt）黑体居中加粗，1.5 倍行距，段前 18 磅、段后 12 磅；二级标题（节）四号（14pt）黑体左对齐加粗，1.5 倍行距，段前 12 磅、段后 6 磅；三级标题（小节）小四号（12pt）黑体左对齐加粗，1.5 倍行距，段前 6 磅、段后 6 磅；各级标题支持独立控制首行缩进。
- **标题编号**：必须采用 Word 原生多级列表自动编号（`w:numPr` 绑定到 `abstractNum`），一级标题为“第X章”，二级标题为“X.X”，三级标题为“X.X.X”；标题段落纯文本中自动剥离手打序号；摘要、目录、参考文献、致谢、附录等独立章节保持无编号。
- **页眉页脚**：页眉居中宋体小五号“毕业设计（论文）”，页脚居中插入阿拉伯数字页码。
- **图表规范**：图名在图下，表名在表上（标准三线表），文字五号宋体/Times New Roman（10.5pt），居中对齐，段前段后 6 磅。
- **目录规范**：标题“目　　录”（黑体三号居中），按三级标题生成，正文小四号宋体，1.5 倍行距。
- **参考文献**：采用 GB/T 7714-2015 顺序编码制，五号宋体/Times New Roman（10.5pt），单倍行距，缩进-文本之前 0 字符，悬挂缩进 2 字符（7.4mm）。

## Launch Dashboard

When AI already has a DOCX or DOTX, preprocess it before opening the dashboard:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/serve_dashboard.py" `
  --input INPUT.docx `
  --output OUTPUT.docx
```

The server binds to `127.0.0.1`, chooses an available port, opens the local page, and preprocesses the source into a session containing document evidence, body paragraph anchors, reference entries, and the source fingerprint. Use `--no-open` for automated checks. The printed URL is the Dashboard URL.

For the webpage route, report the Dashboard URL and wait for the human's in-page submission. A direct request can instead use the direct workflow. Do not automate the webpage to simulate submission.

At startup, the Dashboard detects Microsoft Word and LibreOffice as local renderers and reports each result in the **验收** panel. The user can select automatic mode, Microsoft Word, or LibreOffice. Automatic mode prefers Word and falls back to LibreOffice; an explicit selection must never be replaced silently. If LibreOffice is unavailable, the panel offers its official download page and an Agent-managed background installation. The managed package is stored under the local `.runtime/` directory, verified by re-detecting `soffice`, and is ignored by Git. Structural formatting and verification remain available even when no renderer is installed.

In the very same turn after launching the Dashboard server, AI MUST launch the waiting listener as a background task so the system reactively wakes up AI the moment the user clicks **确认设置并交给 AI** in the browser:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/wait_for_dashboard.py" `
  --url DASHBOARD_URL `
  --output HANDOFF.json
```

Do not end the agent turn without launching this background task; otherwise, the user's browser submission cannot automatically trigger AI execution until the user manually types another chat message.

While the Dashboard is open, also wait for imported template-analysis requests from the page:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/local_ai_analysis.py" wait `
  --url DASHBOARD_URL `
  --output ANALYSIS_REQUEST.json
```

When that command returns, read the complete request. Analyze `document_text`, `baseline_analysis.inferred_spec`, the effective styles, section geometry, headers/footers, fields, and numbering. Write only an evidence-backed supported spec patch to `ANALYSIS_SPEC.json`, including a short `ai_analysis_summary`, then return it to the same local Dashboard:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/local_ai_analysis.py" submit `
  --url DASHBOARD_URL `
  --request ANALYSIS_REQUEST.json `
  --spec ANALYSIS_SPEC.json
```

The page polls the local task and automatically fills the returned result into the template controls. Do not ask for an API key and do not call `llm_client.py`; imported-template semantic analysis belongs to the active conversational AI. Start another `wait` command after each completed request if the user may import another template.

Use the dashboard to inspect the preprocessed source, select a preset, and adjust supported tokens. In an AI session, the source file is locked to the document AI supplied so the evidence, anchors, and SHA-256 fingerprint cannot silently drift. The primary action is **确认设置并交给 AI**. Do not ask the user to export or import citation JSON or a visual-review JSON.

After the user confirms, read `HANDOFF.json`. It contains the confirmed `spec`, every applicable application method registered by `apply_spec.py`, exact local tool paths and arguments, citation context, output path, and verification choices. Treat it as task-local control data, not as a user deliverable.

When visual verification is enabled, `handoff.tools.verify_visual` contains the exact local `scripts/render_docx.py` path, renderer arguments, detection snapshot, report contract, and a local `report_endpoint`. After writing the output DOCX, run that tool to render every page, establish the selected model's image-input capability, and post one final report to that endpoint. Include the output DOCX SHA-256, actual renderer, and exact rendered page count. A `passed` or `failed` report must list every rendered page; when the selected renderer or image input is unavailable, post `skipped` with the reason. Do not tell the user that visual verification passed until the endpoint accepts the final status.

For numbered small-point lists, use the Dashboard **列表** setting. Treat `1.`, `1)`, `(1)`, `（1）`, `一、`, `（一）`, `A.`, `a)`, and `I.` as list-marker choices, not heading levels. Apply them through real Word numbering definitions and leave bullet lists unchanged.

For AI-assisted citation placement, use the Dashboard **请求 AI 分析引用位置** button. It sets `handoff.citation_request=true` while `handoff.citation_context.task` supplies the preprocessed paragraph anchors and references. Only when requested may AI decide which claims require support and add validated `citations.placements` to the confirmed specification. Require each placement to reuse the supplied paragraph index, SHA-256 fingerprint, and existing reference ID. Reject stale fingerprints, unknown references, duplicate paragraph placements, or invented citation text; do not let AI directly rewrite the DOCX.

Expose a value as an editable Dashboard setting only when `scripts/apply_spec.py` registers and implements a deterministic DOCX/OOXML write method for it. Keep citation-system names and other workflow metadata read-only. Confirm `/api/capabilities` reports full coverage before presenting the dashboard.

## Apply Formatting

For the webpage route, `apply_spec.py` verifies the confirmed handoff's source and output paths before passing the spec to the shared core. For direct requests use `format_cli.py`; it needs no Dashboard session.

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/apply_spec.py" INPUT.docx OUTPUT.docx `
  --handoff HANDOFF.json `
  --spec task-local-confirmed-spec.json
```

Omit `--spec` when no citation placement must be added. When supplied, it may differ from `handoff.spec` only in `citations`; all formatting values remain locked to the user-confirmed handoff. Do not invoke the script with `--spec` alone, `--preset-id`, or a manually reconstructed handoff.

Preserve the source file. The legacy direct-font cleanup switch now overrides only requested properties on selected objects; it does not erase unrelated emphasis or table formatting. Local edits write the requested property without changing shared styles unless a style target is explicit.

Never pass an `official-template` workflow to `apply_spec.py`. Download or obtain the current publisher or institution template, analyze it, and use that DOCX/DOTX as the authoring authority. Treat an extracted numeric specification as an audit aid, not a replacement for the template package.

For parameterized workflows, apply figure/table caption styles separately, preserve unmanaged publisher header/footer content, insert page numbers as PAGE/NUMPAGES fields, and format reference paragraphs through a named Bibliography/reference style. Reuse and rebuild an existing standalone PAGE-field paragraph in the selected header/footer instead of adding a duplicate; structural acceptance requires exactly one PAGE field in each target part. When `table_of_contents.enabled` is true, write a real `TOC` Word field from Heading styles, not a static manual list; `title`, `max_heading_level`, and `page_break_after` are deterministic write tokens. The field also carries a cached heading-label snapshot so LibreOffice can show non-empty visual evidence; Word remains the authority for refreshed page numbers. Do not hand-type caption, page, directory, or citation numbers. Add missing captions with `SEQ Figure`/`SEQ Table` fields and cross-references with `REF`; update all fields in Word before delivery.

Before running the apply script, merge AI's validated citation placements into a task-local copy of `handoff.spec`. Keep the original handoff unchanged. For an `official-template` handoff, copy the matching official template to the requested output path and use the publisher package's own styles and OOXML; do not call `apply_spec.py`.

## Validate

Run structural checks before delivery:

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/analyze_docx.py" OUTPUT.docx --output final-analysis.json
```

Then run `scripts/render_docx.py` with the confirmed `verification.render_method`. Automatic mode renders with Microsoft Word when available and LibreOffice as a fallback; an explicit Word or LibreOffice choice has no fallback. Inspect every generated page PNG at 100% zoom for clipping, overlap, broken tables, missing glyphs, page-break drift, header/footer placement, and unexpected blank pages. Reapply or repair the owning specification and rerender until clean.

The Dashboard renderer status is advisory for the current machine: it detects installed executables and any managed LibreOffice runtime. Download progress covers the real package transfer and MSI extraction; completion is reported only after `soffice` is found and executable.

Always run the structural acceptance checks included in the application report. Honor the Dashboard visual-review switch separately. If enabled, render every page and ask the selected model to inspect the images only when its runtime declares image-input support. The AI owns the structured visual report; the user does not import a report file. If the selected model lacks image input or capability cannot be established, set visual status to `skipped`, state the reason, and continue without claiming a visual pass. `pending` means the dashboard is still waiting for AI work and is never a pass; the local download-only button has no connected image model and must report `skipped` rather than `pending`.

Disclose these limitations when they apply:

- Browser preview is advisory and is not Word's pagination engine.
- `python-docx` can write and detect TOC, REF, PAGEREF, bibliography, and other fields, but cannot evaluate them; update fields in Word before final delivery. A TOC only uses real Heading styles, never bold body text.
- Structural acceptance must fail when a TOC is enabled but no non-empty Heading paragraph exists within its selected level range. Ask AI or the user to assign the intended paragraphs to real Word Heading styles before rerunning the formatter; never accept an empty directory page as a visual pass.
- `official-template` workflows are not numeric presets. Use the matching current publisher or institution package, preserve its semantic styles and package structure, and follow its downstream submission requirements.
- Macro-enabled DOCM/DOTM files are unsupported by the dashboard MVP.

## Deliver

Return the formatted DOCX and a concise summary of the chosen authority, applied preset or template, unresolved rules, and render status. Do not return temporary uploads, page images, or internal analysis JSON unless the user asks for them.
