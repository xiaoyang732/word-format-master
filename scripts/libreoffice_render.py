#!/usr/bin/env python3
"""Refresh Writer fields through UNO and export the document as PDF."""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

import uno


def property_value(name: str, value):
    item = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def connect(port: int):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_context,
    )
    target = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    last_error = None
    for _ in range(60):
        try:
            return resolver.resolve(target)
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"无法连接 LibreOffice UNO：{last_error}")


def refresh_and_export(source: Path, output: Path, port: int) -> dict:
    context = connect(port)
    service_manager = context.ServiceManager
    desktop = service_manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    document = desktop.loadComponentFromURL(
        source.resolve().as_uri(),
        "_blank",
        0,
        (
            property_value("Hidden", True),
            property_value("ReadOnly", False),
            property_value("UpdateDocMode", 3),
            property_value("MacroExecutionMode", 0),
        ),
    )
    if document is None:
        raise RuntimeError("LibreOffice 无法打开输入文档")
    index_count = 0
    try:
        try:
            document.updateLinks()
        except Exception:
            pass
        try:
            document.getTextFields().refresh()
        except Exception:
            pass
        indexes = document.getDocumentIndexes()
        index_count = indexes.getCount()
        for index in range(index_count):
            indexes.getByIndex(index).update()
        try:
            document.calculateAll()
        except Exception:
            pass
        document.storeToURL(
            output.resolve().as_uri(),
            (
                property_value("FilterName", "writer_pdf_Export"),
                property_value("Overwrite", True),
            ),
        )
    finally:
        document.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    if not output.is_file() or output.stat().st_size == 0:
        raise RuntimeError("LibreOffice UNO 没有生成 PDF")
    return {"status": "exported", "document_indexes_updated": index_count}


def main() -> int:
    from cli_output import configure_cli_output
    configure_cli_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--port", type=int, required=True)
    args = parser.parse_args()
    try:
        result = refresh_and_export(Path(args.input), Path(args.output), args.port)
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
