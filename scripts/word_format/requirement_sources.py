"""Read complete requirement text with stable source anchors, without modifying files."""

from pathlib import Path
from xml.etree import ElementTree

from lxml import etree

from .contracts import FormatError, sha256
from .namespaces import qn


def extract_requirements(path):
    from config import MAX_DOCUMENT_BYTES

    source = Path(path).resolve()
    if not source.is_file() or source.stat().st_size > MAX_DOCUMENT_BYTES:
        raise FormatError("Requirements file is missing or exceeds the size limit")
    original_hash = sha256(source)
    blocks = []
    issues = []

    def add(anchor, text, **extra):
        if text.strip():
            blocks.append({"anchor": anchor, "text": text.strip(), **extra})

    suffix = source.suffix.lower()
    if suffix in {".txt", ".md"}:
        for i, line in enumerate(source.read_text(encoding="utf-8-sig").splitlines(), 1):
            add(f"line:{i}", line)
    elif suffix in {".docx", ".dotx"}:
        from analyze_docx import package_inventory, read_xml

        package, _ = package_inventory(source)
        try:
            names = ["word/document.xml"] + sorted(
                n
                for n in package.namelist()
                if n.startswith(
                    (
                        "word/header",
                        "word/footer",
                        "word/footnotes",
                        "word/endnotes",
                        "word/comments",
                    )
                )
                and n.endswith(".xml")
            )
            for name in names:
                root = read_xml(package, name)
                if root is None:
                    continue
                root = etree.fromstring(ElementTree.tostring(root))
                # Preserve body/table reading order and avoid duplicating nested
                # text-box paragraphs. A row retains context from its label cell.
                rows = paragraphs = 0
                for node in root.iter(qn("w:tr"), qn("w:p")):
                    if node.tag == qn("w:tr"):
                        rows += 1
                        if any(a.tag == qn("w:tr") for a in node.iterancestors()):
                            continue
                        cells = [
                            {
                                "anchor": f"{name}#row:{rows}/cell:{i}",
                                "text": " ".join(t.text or "" for t in c.iter(qn("w:t"))),
                            }
                            for i, c in enumerate(node.findall(qn("w:tc")), 1)
                        ]
                        add(
                            f"{name}#row:{rows}",
                            " ".join(c["text"] for c in cells),
                            cells=cells,
                            scope=name,
                        )
                        continue
                    paragraphs += 1
                    if any(a.tag == qn("w:tr") for a in node.iterancestors()):
                        continue
                    value = "".join(
                        t.text or ""
                        for t in node.iter(qn("w:t"))
                        if next(t.iterancestors(qn("w:p")), None) is node
                    )
                    add(f"{name}#p:{paragraphs}", value, scope=name)
                if any(n.tag in {qn("w:ins"), qn("w:del")} for n in root.iter()):
                    issues.append(
                        {
                            "anchor": name,
                            "reason": "Requirements contain revisions; resolve authoritative text",
                        }
                    )
                nontext = {
                    tag: len(list(root.iter(qn("w:" + tag))))
                    for tag in ("drawing", "pict", "object", "altChunk")
                }
                nontext = {tag: count for tag, count in nontext.items() if count}
                if nontext:
                    issues.append(
                        {
                            "anchor": name,
                            "features": nontext,
                            "reason": "Non-text requirement content needs visual review or OCR; extracted text alone may be incomplete",
                        }
                    )
                if name.startswith("word/comments"):
                    issues.append(
                        {
                            "anchor": name,
                            "reason": "Requirements include comments; resolve their authority relative to body text",
                        }
                    )
        finally:
            package.close()
    elif suffix == ".pdf":
        try:
            import pypdfium2 as pdfium
        except ImportError as exc:
            raise FormatError("PDF text extraction requires pypdfium2") from exc
        pdf = pdfium.PdfDocument(str(source))
        try:
            for i in range(len(pdf)):
                page = pdf[i]
                textpage = page.get_textpage()
                try:
                    value = textpage.get_text_range()
                    add(
                        f"page:{i + 1}",
                        value,
                        page=i + 1,
                        bounds=[0, 0, page.get_width(), page.get_height()],
                    )
                    if not value.strip():
                        issues.append(
                            {
                                "anchor": f"page:{i + 1}",
                                "reason": "No readable text; OCR is required",
                            }
                        )
                    elif any(page.get_objects(filter=[pdfium.raw.FPDF_PAGEOBJ_IMAGE])):
                        issues.append(
                            {
                                "anchor": f"page:{i + 1}",
                                "reason": "Page contains images alongside text; review for requirements needing OCR",
                            }
                        )
                finally:
                    textpage.close()
                    page.close()
        finally:
            pdf.close()
    else:
        raise FormatError(
            "Requirements support UTF-8 TXT/MD, DOCX/DOTX and text PDF", "unsupported"
        )
    if not blocks and not issues:
        issues.append({"anchor": "document", "reason": "No readable requirements found"})
    if sha256(source) != original_hash:
        raise FormatError("Requirements source changed during extraction; retry")
    return {
        "source": {"kind": "requirements", "path": str(source), "sha256": original_hash},
        "blocks": blocks,
        "issues": issues,
    }
