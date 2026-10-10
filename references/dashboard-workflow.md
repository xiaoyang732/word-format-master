# Dashboard Workflow

Read this only for webpage configuration. Resolve `SKILL_DIR` and `PYTHON_BIN` using the entrypoint before running these commands.

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
