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
    from cli_output import configure_cli_output
    from json_output import ensure_protected_paths, write_json
    configure_cli_output()
    ensure_protected_paths(args.output, args.input)
    source = Path(args.input).resolve()
    task = build_citation_task(
        Document(str(source)),
        filename=source.name,
        source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
    )
    destination = Path(args.output).resolve()
    write_json(destination, task, source)
    print(f"Exported {len(task['paragraphs'])} body anchors and {len(task['references'])} references to {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
