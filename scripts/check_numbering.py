"""Reject typed structural numbering in a generated DOCX (read-only)."""

import argparse
import sys

from cli_output import configure_cli_output
from docx import Document
from json_output import ensure_protected_paths, write_json
from word_format.numbering_audit import audit_numbering


def main():
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    try:
        ensure_protected_paths(args.output, args.input)
        result = audit_numbering(Document(args.input))
        write_json(args.output, result, args.input)
    except (OSError, ValueError, KeyError) as exc:
        print(str(exc), file=sys.stderr)
        return 2
    return 0 if result["status"] == "passed" else 2


if __name__ == "__main__":
    raise SystemExit(main())
