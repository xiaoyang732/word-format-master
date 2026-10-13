#!/usr/bin/env python3
"""Direct AI formatting tools. No dashboard or browser is required."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from word_format import apply_plan, audit_document, build_plan, capabilities, inspect_document, verify_plan
from word_format.contracts import FormatError
from json_output import ensure_protected_paths, write_json


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def main():
    from cli_output import configure_cli_output
    configure_cli_output()
    parser=argparse.ArgumentParser(description=__doc__)
    commands=parser.add_subparsers(dest="command",required=True)
    p=commands.add_parser("capabilities");p.add_argument("--output")
    p=commands.add_parser("inspect");p.add_argument("input");p.add_argument("--output")
    p=commands.add_parser("audit");p.add_argument("input");p.add_argument("--request",required=True);p.add_argument("--output")
    p=commands.add_parser("plan");p.add_argument("input");p.add_argument("destination");p.add_argument("--request",required=True);p.add_argument("--output")
    p=commands.add_parser("apply");p.add_argument("--plan",required=True);p.add_argument("--report");p.add_argument("--overwrite",action="store_true")
    p=commands.add_parser("verify");p.add_argument("--plan",required=True);p.add_argument("--report",required=True);p.add_argument("--output")
    args=parser.parse_args()
    try:
        destination=args.report if args.command=="apply" else args.output
        protected=[Path(v).resolve() for v in (getattr(args,"input",None),getattr(args,"destination",None),getattr(args,"request",None),getattr(args,"plan",None)) if v]
        if args.command in {"apply","verify"}:
            plan=read_json(args.plan)
            protected.extend([Path(plan["source"]["path"]).resolve(),Path(plan["output_path"]).resolve()])
        if args.command=="verify": protected.append(Path(args.report).resolve())
        ensure_protected_paths(destination, *protected)
        if args.command=="capabilities": result=capabilities()
        elif args.command=="inspect": result=inspect_document(args.input)
        elif args.command=="audit": result=audit_document(args.input,read_json(args.request))
        elif args.command=="plan": result=build_plan(args.input,args.destination,read_json(args.request))
        elif args.command=="apply": result=apply_plan(read_json(args.plan),overwrite=args.overwrite)
        else: result=verify_plan(read_json(args.plan),read_json(args.report))
        rendered=json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False)+"\n"
        if destination:
            write_json(destination, result, *protected)
        else: print(rendered,end="")
        return 0 if result.get("status") not in {"failed","unsupported","needs_clarification"} else 2
    except (FormatError,ValueError,OSError,KeyError,TypeError) as exc:
        print(json.dumps({"status":getattr(exc,"status","failed"),"error":str(exc)},ensure_ascii=False),file=sys.stderr)
        return 2


if __name__=="__main__": raise SystemExit(main())
