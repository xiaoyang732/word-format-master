#!/usr/bin/env python3
"""Export body/reference anchors for an AI citation-placement review."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from docx import Document

from citation_workflow import build_citation_task


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Source DOCX")
    parser.add_argument("--output", required=True, help="Task JSON output")
    args = parser.parse_args()
    source = Path(args.input).resolve()
    task = build_citation_task(
        Document(str(source)),
        filename=source.name,
        source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
    )
    destination = Path(args.output).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(task, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Exported {len(task['paragraphs'])} body anchors and {len(task['references'])} references to {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
