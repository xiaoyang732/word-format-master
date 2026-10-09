#!/usr/bin/env python3
"""Wait for a user to submit an AI-backed Word formatting dashboard session."""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", required=True, help="Dashboard base URL")
    parser.add_argument("--output", required=True, help="Write the AI handoff JSON here")
    parser.add_argument("--timeout", type=float, default=3600, help="Maximum seconds to wait")
    parser.add_argument("--poll-interval", type=float, default=0.5, help="Seconds between checks")
    args = parser.parse_args()
    endpoint = urljoin(args.url.rstrip("/") + "/", "api/session/result")
    deadline = time.monotonic() + max(1.0, args.timeout)
    while time.monotonic() < deadline:
        try:
            with urlopen(Request(endpoint, headers={"Accept": "application/json"}), timeout=10) as response:
                payload = json.loads(response.read().decode("utf-8"))
                if getattr(response, "status", 200) == 202 or payload.get("status") == "waiting":
                    time.sleep(max(0.1, args.poll_interval))
                    continue
                destination = Path(args.output).resolve()
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                print(f"Dashboard handoff ready: {destination}")
                return 0
        except HTTPError as exc:
            if exc.code != 202:
                detail = exc.read().decode("utf-8", errors="replace")
                raise SystemExit(f"Dashboard handoff failed ({exc.code}): {detail}") from exc
        except URLError:
            pass
        time.sleep(max(0.1, args.poll_interval))
    raise SystemExit("Timed out waiting for dashboard confirmation")


if __name__ == "__main__":
    raise SystemExit(main())
