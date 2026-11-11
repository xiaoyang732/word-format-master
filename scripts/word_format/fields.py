"""Shared field classification for header content and page-number ownership."""

import re

from .namespaces import qn


def field_codes(root):
    if root is None:
        return []
    return [" ".join((node.text or "").split()) for node in root.iter(qn("w:instrText"))] + [
        " ".join(node.get(qn("w:instr"), "").split()) for node in root.iter(qn("w:fldSimple"))
    ]


def page_number_paragraphs(root):
    """Identify complete page-only paragraphs, including imported simple fields."""
    if root is None:
        return []
    result = []
    for paragraph in root.iter(qn("w:p")):
        codes = field_codes(paragraph)
        if not any(re.match(r"^PAGE(?:\s|$)", code, re.I) for code in codes):
            continue
        if any(not re.match(r"^(?:PAGE|NUMPAGES)(?:\s|$)", code, re.I) for code in codes):
            continue
        style = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:pStyle"))
        managed = style is not None and style.get(qn("w:val")) == "WFMPageNumber"
        visible = "".join(node.text or "" for node in paragraph.iter(qn("w:t"))).strip()
        if (
            managed
            or not visible
            or re.fullmatch(
                r"(?:Page\s*)?(?:第\s*)?\d+(?:\s*(?:of|/|页，共)\s*\d+\s*页?)?", visible, re.I
            )
        ):
            result.append(paragraph)
    return result


def _sync_toc_entry_runs_from_styles(document) -> int:
    """After Word COM updates TOC fields, ensure every run (<w:t> and <w:tab/>) in
    native toc 1..9 paragraphs carries the effective font properties of its TOC style
    in leading <w:rPr> and <w:pPr>/<w:rPr>.
    """
    from .properties import FONT_TAGS, element, paragraph_style, read_font, write_font

    synced = 0
    for p in document.element.body.iter(qn("w:p")):
        style = paragraph_style(document, p)
        if style is None:
            continue
        norm = (style.name or "").lower().replace(" ", "")
        if not re.fullmatch(r"(?:toc|目录)[1-9]", norm):
            continue
        style_p = element(p, "pPr")
        # Read effective style font values
        from .inspect import _style_paragraph

        probe_p = _style_paragraph(style.name, document)
        style_fonts = {name: read_font(document, probe_p, None, name) for name in FONT_TAGS}
        ppr_rpr = element(style_p, "rPr")
        for name, val in style_fonts.items():
            if val is not None and not isinstance(val, dict):
                write_font(ppr_rpr, name, val)
        for run in p.iter(qn("w:r")):
            if any(
                a.tag in {qn("w:del"), qn("w:ins"), qn("w:txbxContent")}
                for a in run.iterancestors()
                if a is not p
            ):
                continue
            has_text = any(t.text for t in run.findall(qn("w:t")))
            has_tab = run.find(qn("w:tab")) is not None
            if has_text or has_tab:
                rpr = element(run, "rPr")
                for name, val in style_fonts.items():
                    if val is not None and not isinstance(val, dict):
                        write_font(rpr, name, val)
                synced += 1
    return synced


def strip_last_rendered_page_breaks_in_doc(document) -> int:
    """Remove Word-cached <w:lastRenderedPageBreak/> nodes so structure checks stay stable."""
    removed = 0
    for br in list(document.element.body.iter(qn("w:lastRenderedPageBreak"))):
        parent = br.getparent()
        if parent is not None:
            parent.remove(br)
            removed += 1
    return removed


def refresh_document_fields(
    input_path,
    destination=None,
    *,
    in_place: bool = False,
    overwrite: bool = False,
    strip_last_rendered_page_breaks: bool = True,
    timeout: int = 120,
) -> dict:
    """Update TOC, REF, SEQ, and STYLEREF fields via Word COM and strip cached page-break markers."""
    import os
    from pathlib import Path
    import shutil
    import subprocess
    import uuid
    from docx import Document
    from .contracts import FormatError, SCHEMA_VERSION, output_path, sha256, source_path

    src_path = source_path(input_path)
    src_hash = sha256(src_path)
    if in_place:
        if destination is not None and Path(destination).resolve() != src_path:
            raise FormatError("Do not combine --in-place with a different destination path")
        out_path = src_path
    else:
        if destination is None:
            raise FormatError("Specify destination path or pass --in-place")
        out_path = output_path(src_path, destination)
        if out_path.exists() and not overwrite:
            raise FormatError(
                "Output already exists; choose a new filename or explicitly request overwrite"
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = out_path.with_name(f".{out_path.stem}.{uuid.uuid4().hex}.docx")
    shutil.copy2(src_path, temp_path)
    method_used = "xml_cleanup_only"
    try:
        ps = shutil.which("powershell") or shutil.which("pwsh")
        if ps:
            script = (
                "$ErrorActionPreference = 'Stop'; "
                "$word = New-Object -ComObject Word.Application; "
                "$word.Visible = $false; "
                "$word.DisplayAlerts = 0; "
                "try { "
                f"  $doc = $word.Documents.Open({repr(str(temp_path))}, $false, $false); "
                "  try { "
                "    foreach ($toc in $doc.TablesOfContents) { $toc.Update() } "
                "    $doc.Fields.Update() | Out-Null; "
                "    foreach ($sec in $doc.Sections) { "
                "      foreach ($hdr in $sec.Headers) { $hdr.Range.Fields.Update() | Out-Null } "
                "      foreach ($ftr in $sec.Footers) { $ftr.Range.Fields.Update() | Out-Null } "
                "    } "
                "    $doc.Save(); "
                "  } finally { $doc.Close(0) } "
                "} finally { $word.Quit() }"
            )
            proc = subprocess.run(
                [ps, "-NoProfile", "-NonInteractive", "-Command", script],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
            )
            if proc.returncode == 0:
                method_used = "word_com"
        removed_breaks = 0
        if strip_last_rendered_page_breaks or method_used == "word_com":
            doc = Document(temp_path)
            if method_used == "word_com":
                _sync_toc_entry_runs_from_styles(doc)
            if strip_last_rendered_page_breaks:
                removed_breaks = strip_last_rendered_page_breaks_in_doc(doc)
            doc.save(temp_path)
        if not in_place and sha256(src_path) != src_hash:
            raise FormatError("Source document changed during field refresh; output not published")
        os.replace(temp_path, out_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()

    return {
        "schema_version": SCHEMA_VERSION,
        "kind": "word-format-refresh-fields-report",
        "status": "passed",
        "input": {"path": str(src_path), "sha256": src_hash},
        "output": str(out_path),
        "output_sha256": sha256(out_path),
        "method": method_used,
        "stripped_last_rendered_page_breaks": removed_breaks,
    }
