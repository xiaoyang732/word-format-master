#!/usr/bin/env python3
"""Exchange template-analysis tasks with the active conversational AI."""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urljoin
from urllib.request import Request, urlopen


def request_json(url: str, *, payload: dict | None = None) -> tuple[int, dict]:
    data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(
        url,
        data=data,
        method="POST" if data is not None else "GET",
        headers={"Accept": "application/json", "Content-Type": "application/json"},
    )
    try:
        with urlopen(request, timeout=15) as response:
            return response.status, json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Dashboard request failed ({exc.code}): {detail}") from exc


def wait_for_task(base_url: str, output: Path, timeout: float, poll_interval: float) -> int:
    endpoint = urljoin(base_url.rstrip("/") + "/", "api/analysis/tasks/next")
    deadline = time.monotonic() + max(1.0, timeout)
    while time.monotonic() < deadline:
        try:
            status, task = request_json(endpoint)
            if status == 200 and task.get("status") == "claimed" and task.get("claim_token"):
                output = output.resolve()
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_text(json.dumps(task, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                print(f"Local AI analysis request ready: {output}")
                return 0
        except (URLError, RuntimeError):
            pass
        time.sleep(max(0.1, poll_interval))
    raise SystemExit("Timed out waiting for a local AI template-analysis request")


def submit_result(base_url: str, request_path: Path, spec_path: Path) -> int:
    task = json.loads(request_path.read_text(encoding="utf-8"))
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    task_id = str(task.get("id") or "")
    if not task_id:
        raise SystemExit("Analysis request does not contain a task ID")
    endpoint = urljoin(
        base_url.rstrip("/") + "/",
        f"api/analysis/tasks/{quote(task_id, safe='')}/result",
    )
    _, result = request_json(endpoint, payload={"spec": spec, "claim_token": task.get("claim_token")})
    if result.get("status") != "completed":
        raise SystemExit(result.get("error") or "Dashboard did not accept the analysis result")
    print(f"Local AI analysis returned to Dashboard: {task_id}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    wait_parser = subparsers.add_parser("wait", help="Wait for the next Dashboard analysis request")
    wait_parser.add_argument("--url", required=True)
    wait_parser.add_argument("--output", required=True)
    wait_parser.add_argument("--timeout", type=float, default=3600)
    wait_parser.add_argument("--poll-interval", type=float, default=0.5)
    submit_parser = subparsers.add_parser("submit", help="Return an AI-produced spec to the Dashboard")
    submit_parser.add_argument("--url", required=True)
    submit_parser.add_argument("--request", required=True)
    submit_parser.add_argument("--spec", required=True)
    args = parser.parse_args()
    if args.command == "wait":
        return wait_for_task(args.url, Path(args.output), args.timeout, args.poll_interval)
    return submit_result(args.url, Path(args.request), Path(args.spec))


if __name__ == "__main__":
    raise SystemExit(main())
