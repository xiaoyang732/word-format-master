"""UTF-8 output contract for CLI reports, including redirected Windows pipes."""

import sys


def configure_cli_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")
