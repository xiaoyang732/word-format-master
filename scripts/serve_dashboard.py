#!/usr/bin/env python3
"""Serve the local Word Format Master dashboard on 127.0.0.1."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import mimetypes
import os
import tempfile
import threading
import copy
import uuid
import webbrowser
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from analyze_docx import analyze_docx, extract_document_text
from apply_spec import SUPPORTED_APPLICATION_PATHS, validate_spec
from citation_workflow import build_citation_task
from config import (
    DEFAULT_RUNTIME_ROOT as RUNTIME_ROOT,
    MAX_DOCUMENT_BYTES,
    MAX_ANALYSIS_TASKS,
    ANALYSIS_TASK_LEASE_SECONDS,
    ANALYSIS_TASK_TTL_SECONDS,
    API_VERSION,
    MAX_REQUEST_BYTES,
    PRESETS_PATH,
    SKILL_DIR,
    UI_DIR,
    utc_now as _utc_now,
)
from dashboard_session import (
    build_ai_handoff,
    prepare_dashboard_session,
    public_dashboard_session,
    public_dashboard_session_status,
    record_visual_verification,
)
from parse_requirements import parse_requirements
from runtime_detection import detect_renderers, download_libreoffice
from verify_output import validate_visual_report


def _analysis_task_public(task: dict, *, include_context: bool = False) -> dict:
    payload = {
        "id": task["id"],
        "status": task["status"],
        "filename": task["filename"],
        "source_sha256": task["source_sha256"],
        "created_at": task["created_at"],
    }
    if task.get("error"):
        payload["error"] = task["error"]
    if task.get("result") is not None:
        payload["result"] = copy.deepcopy(task["result"])
    if include_context and task.get("claim_token"):
        payload["claim_token"] = task["claim_token"]
    if include_context:
        payload["instructions"] = (
            "Analyze the DOCX template evidence and text. Return only a supported format spec patch; "
            "do not invent values absent from the document. Include ai_analysis_summary."
        )
        payload["document_text"] = task["document_text"]
        payload["baseline_analysis"] = copy.deepcopy(task["baseline_analysis"])
    return payload


def _timestamp_age_seconds(value: str) -> float | None:
    try:
        timestamp = datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return None
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    return max(0.0, (datetime.now(timezone.utc) - timestamp).total_seconds())


def _cleanup_analysis_tasks(server, *, reserve_slot: bool = False) -> None:
    with server.dashboard_analysis_lock:
        remove: list[str] = []
        for task_id, task in server.dashboard_analysis_tasks.items():
            age = _timestamp_age_seconds(task.get("created_at"))
            if age is None or age > ANALYSIS_TASK_TTL_SECONDS:
                remove.append(task_id)
                continue
            if task.get("status") in {"claimed", "processing"}:
                lease_age = _timestamp_age_seconds(task.get("claimed_at"))
                if lease_age is None or lease_age > ANALYSIS_TASK_LEASE_SECONDS:
                    task["status"] = "waiting_for_ai"
                    task.pop("claim_token", None)
                    task.pop("claimed_at", None)
        for task_id in remove:
            server.dashboard_analysis_tasks.pop(task_id, None)
        limit = MAX_ANALYSIS_TASKS - 1 if reserve_slot else MAX_ANALYSIS_TASKS
        terminal = sorted(
            (
                (task_id, task)
                for task_id, task in server.dashboard_analysis_tasks.items()
                if task.get("status") in {"completed", "failed"}
            ),
            key=lambda item: item[1].get("completed_at") or item[1].get("created_at", ""),
        )
        while len(server.dashboard_analysis_tasks) > limit and terminal:
            task_id, _ = terminal.pop(0)
            server.dashboard_analysis_tasks.pop(task_id, None)


def _runtime_snapshot(server) -> dict:
    detected = detect_renderers(RUNTIME_ROOT)
    lock = getattr(server, "dashboard_runtime_lock", None)
    if lock:
        with lock:
            job = copy.deepcopy(getattr(server, "dashboard_runtime_job", {"status": "idle", "percent": 0}))
    else:
        job = copy.deepcopy(getattr(server, "dashboard_runtime_job", {"status": "idle", "percent": 0}))
    detected["download"] = job
    return detected


def _run_runtime_download(server) -> None:
    def update(payload: dict) -> None:
        with server.dashboard_runtime_lock:
            server.dashboard_runtime_job.update(payload)

    try:
        result = download_libreoffice(RUNTIME_ROOT, update)
        with server.dashboard_runtime_lock:
            server.dashboard_runtime_job.update(result, percent=100, phase="LibreOffice 已安装并可用")
    except Exception as exc:
        with server.dashboard_runtime_lock:
            server.dashboard_runtime_job.update({
                "status": "error",
                "phase": "下载或安装失败",
                "error": str(exc),
            })


def _start_runtime_download(server) -> dict:
    with server.dashboard_runtime_lock:
        current = server.dashboard_runtime_job
        if current.get("status") in {"queued", "downloading", "installing"}:
            return copy.deepcopy(current)
        if detect_renderers(RUNTIME_ROOT).get("libreoffice", {}).get("available"):
            server.dashboard_runtime_job = {
                "status": "ready",
                "phase": "已检测到 LibreOffice",
                "percent": 100,
            }
            return copy.deepcopy(server.dashboard_runtime_job)
        server.dashboard_runtime_job = {
            "status": "queued",
            "phase": "准备从 LibreOffice 官网下载",
            "percent": 0,
            "downloaded": 0,
            "total": 0,
        }
    threading.Thread(target=_run_runtime_download, args=(server,), daemon=True).start()
    return copy.deepcopy(server.dashboard_runtime_job)


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "WordFormatMaster/0.1"
    trusted_hosts = {"127.0.0.1", "localhost"}

    def log_message(self, format: str, *args) -> None:
        # Log to stderr only if DASHBOARD_DEBUG environment variable is set
        if os.environ.get("DASHBOARD_DEBUG"):
            import sys
            sys.stderr.write(f"[{self.log_date_time_string()}] {format % args}\n")

    def _send_bytes(
        self,
        data: bytes,
        content_type: str,
        status: int = 200,
        *,
        download_name: str | None = None,
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Word-Format-API-Version", API_VERSION)
        if download_name:
            safe_name = Path(download_name).name.replace('"', "")
            self.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'",
        )
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload, status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send_bytes(data, "application/json; charset=utf-8", status)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > MAX_REQUEST_BYTES:
            raise ValueError(f"Request is empty or exceeds the {MAX_REQUEST_BYTES // (1024 * 1024)} MiB limit")
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def _request_host_is_trusted(self) -> bool:
        host_header = str(self.headers.get("Host") or "").strip()
        try:
            hostname = urlparse(f"//{host_header}").hostname
        except ValueError:
            return False
        return bool(hostname and hostname.lower() in self.trusted_hosts)

    def _reject_untrusted_host(self) -> bool:
        if self._request_host_is_trusted():
            return False
        self._send_json({"error": "untrusted host"}, HTTPStatus.FORBIDDEN)
        return True

    def do_GET(self) -> None:
        if self._reject_untrusted_host():
            return
        path = urlparse(self.path).path
        if path == "/api/presets":
            self._send_bytes(PRESETS_PATH.read_bytes(), "application/json; charset=utf-8")
            return
        if path == "/api/capabilities":
            self._send_json({
                "api_version": API_VERSION,
                "engine": "python-docx + targeted OOXML",
                "settings": [
                    {"path": token_path, "method": method}
                    for token_path, method in SUPPORTED_APPLICATION_PATHS.items()
                ],
            })
            return
        if path == "/api/session":
            self._send_json(public_dashboard_session(getattr(self.server, "dashboard_session", None)))
            return
        if path == "/api/session/status":
            self._send_json(public_dashboard_session_status(getattr(self.server, "dashboard_session", None)))
            return
        if path == "/api/session/result":
            session = getattr(self.server, "dashboard_session", None)
            if not session:
                self._send_json({"error": "No AI dashboard session is active"}, HTTPStatus.NOT_FOUND)
            elif session.get("result") is None:
                self._send_json({"status": "waiting", "session_id": session["id"]}, HTTPStatus.ACCEPTED)
            else:
                self._send_json(session["result"])
            return
        if path == "/api/verification/capabilities":
            runtime = _runtime_snapshot(self.server)
            self._send_json({
                "structural": {"available": True, "method": "python-docx + targeted OOXML + package re-open"},
                "visual": {
                    "available": runtime["visual_ready"],
                    "method": "AI renders every page and validates a model-produced visual review after modification",
                    "renderer_ready": runtime["renderer_ready"],
                    "pdf_rasterizer": runtime["pdf_rasterizer"],
                    "model_selection": ["auto", "current", "custom"],
                    "capability_field": "model.supports_image_input",
                    "unsupported_model_behavior": "skip_and_report",
                },
            })
            return
        if path == "/api/runtime/status":
            self._send_json(_runtime_snapshot(self.server))
            return
        if path == "/api/analysis/tasks/next":
            _cleanup_analysis_tasks(self.server)
            with self.server.dashboard_analysis_lock:
                pending = next(
                    (task for task in self.server.dashboard_analysis_tasks.values() if task["status"] == "waiting_for_ai"),
                    None,
                )
                if pending:
                    pending["status"] = "claimed"
                    pending["claim_token"] = uuid.uuid4().hex
                    pending["claimed_at"] = _utc_now()
                self._send_json(
                    _analysis_task_public(pending, include_context=True) if pending else {"status": "idle"},
                    HTTPStatus.OK if pending else HTTPStatus.ACCEPTED,
                )
            return
        if path.startswith("/api/analysis/tasks/"):
            task_id = path.removeprefix("/api/analysis/tasks/").strip("/")
            with self.server.dashboard_analysis_lock:
                task = self.server.dashboard_analysis_tasks.get(task_id)
                if not task:
                    self._send_json({"error": "Analysis task not found"}, HTTPStatus.NOT_FOUND)
                else:
                    self._send_json(_analysis_task_public(task))
            return
        files = {
            "/": UI_DIR / "index.html",
            "/index.html": UI_DIR / "index.html",
            "/app.js": UI_DIR / "app.js",
            "/styles.css": UI_DIR / "styles.css",
        }
        target = files.get(path)
        if not target or not target.is_file():
            self._send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type == "application/javascript":
            content_type += "; charset=utf-8"
        self._send_bytes(target.read_bytes(), content_type)

    def do_POST(self) -> None:
        if self._reject_untrusted_host():
            return
        path = urlparse(self.path).path
        try:
            payload = self._read_json()
            if path == "/api/analyze/text":
                text = str(payload.get("text", ""))
                if not text.strip():
                    raise ValueError("Formatting requirements are empty")
                self._send_json(parse_requirements(text))
                return
            if path == "/api/analyze/docx":
                result = self._analyze_upload(payload)
                self._send_json(result, HTTPStatus.ACCEPTED)
                return
            if path.startswith("/api/analysis/tasks/") and path.endswith("/result"):
                self._send_json(self._complete_analysis(path, payload))
                return
            if path == "/api/session/submit":
                self._send_json(self._submit_session(payload))
                return
            if path == "/api/runtime/download":
                self._send_json(_start_runtime_download(self.server), HTTPStatus.ACCEPTED)
                return
            if path == "/api/session/verification":
                self._send_json(self._record_visual_verification(payload))
                return
            if path == "/api/verification/visual":
                self._send_json(validate_visual_report(payload.get("report") or payload, payload.get("page_count")))
                return
            if path == "/api/validate":
                errors = validate_spec(payload.get("spec") or payload)
                self._send_json({"valid": not errors, "errors": errors})
                return
            self._send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
        except json.JSONDecodeError as exc:
            self._send_json({"error": "Invalid JSON in request body", "details": str(exc)}, HTTPStatus.BAD_REQUEST)
        except ValueError as exc:
            self._send_json({"error": "Invalid request data", "details": str(exc)}, HTTPStatus.BAD_REQUEST)
        except OSError as exc:
            self._send_json({"error": "File system error", "details": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
        except Exception as exc:
            self._send_json({"error": "Unexpected processing error", "details": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def _decode_upload(self, payload) -> tuple[str, bytes]:
        filename = Path(str(payload.get("filename", "document.docx"))).name
        if Path(filename).suffix.lower() not in {".docx", ".dotx"}:
            raise ValueError("Only DOCX and DOTX files are accepted")
        try:
            data = base64.b64decode(payload.get("data", ""), validate=True)
        except Exception as exc:
            raise ValueError("Upload is not valid base64 data") from exc
        if not data or len(data) > MAX_DOCUMENT_BYTES:
            raise ValueError("File is empty or exceeds the 30 MB limit")
        return filename, data

    def _analyze_upload(self, payload):
        filename, data = self._decode_upload(payload)
        source_sha256 = hashlib.sha256(data).hexdigest()
        task_id = uuid.uuid4().hex
        _cleanup_analysis_tasks(self.server, reserve_slot=True)
        with self.server.dashboard_analysis_lock:
            if len(self.server.dashboard_analysis_tasks) >= MAX_ANALYSIS_TASKS:
                raise ValueError("Analysis queue is full; wait for existing tasks to finish")
            self.server.dashboard_analysis_tasks[task_id] = {
                "id": task_id,
                "status": "preprocessing",
                "filename": filename,
                "source_sha256": source_sha256,
                "created_at": _utc_now(),
                "data": data,
                "result": None,
                "error": None,
            }
        try:
            with tempfile.TemporaryDirectory(prefix="word-format-master-") as temp:
                path = Path(temp) / filename
                path.write_bytes(data)
                result = analyze_docx(path)
                if path.suffix.lower() == ".docx":
                    try:
                        from docx import Document

                        citation_task = build_citation_task(
                            Document(str(path)),
                            filename=filename,
                            source_sha256=source_sha256,
                        )
                        result["citation_context"] = {
                            "status": "ready",
                            "paragraph_count": len(citation_task["paragraphs"]),
                            "reference_count": len(citation_task["references"]),
                            "task": citation_task,
                        }
                    except (KeyError, ValueError):
                        result["citation_context"] = {
                            "status": "unavailable",
                            "reason": "Strict OOXML templates require a supported parser before citation analysis",
                            "task": None,
                        }
                document_text = extract_document_text(path)
            with self.server.dashboard_analysis_lock:
                task = self.server.dashboard_analysis_tasks[task_id]
                task.update({
                    "status": "waiting_for_ai",
                    "document_text": document_text,
                    "baseline_analysis": result,
                })
            return {
                "status": "waiting_for_ai",
                "task_id": task_id,
                "message": "已交给当前对话 AI 分析，结果将自动回填。",
            }
        except Exception:
            with self.server.dashboard_analysis_lock:
                self.server.dashboard_analysis_tasks.pop(task_id, None)
            raise

    def _complete_analysis(self, path: str, payload: dict):
        task_id = path.removeprefix("/api/analysis/tasks/").removesuffix("/result").strip("/")
        spec = payload.get("spec") or payload.get("spec_patch")
        if not isinstance(spec, dict) or not spec:
            raise ValueError("Conversational AI result must contain a non-empty spec object")
        with self.server.dashboard_analysis_lock:
            task = self.server.dashboard_analysis_tasks.get(task_id)
            if not task:
                raise ValueError("Analysis task not found")
            if task["status"] == "completed":
                return _analysis_task_public(task)
            if task["status"] != "claimed" or payload.get("claim_token") != task.get("claim_token"):
                raise ValueError("Analysis task is not claimed by this worker")
            claim_token = task["claim_token"]
            task["status"] = "processing"
            data = task["data"]
            filename = task["filename"]
        try:
            with tempfile.TemporaryDirectory(prefix="word-format-master-ai-") as temp:
                source = Path(temp) / filename
                source.write_bytes(data)
                result = analyze_docx(source, semantic_spec=spec)
                result["citation_context"] = copy.deepcopy(task["baseline_analysis"].get("citation_context"))
            with self.server.dashboard_analysis_lock:
                if task.get("status") != "processing" or task.get("claim_token") != claim_token:
                    raise ValueError("Analysis task lease expired before the result was committed")
                task["status"] = "completed"
                task["completed_at"] = _utc_now()
                task["result"] = result
                for key in ("data", "document_text", "baseline_analysis", "claim_token", "claimed_at"):
                    task.pop(key, None)
            return _analysis_task_public(task)
        except Exception as exc:
            with self.server.dashboard_analysis_lock:
                if task.get("claim_token") == claim_token:
                    task["status"] = "failed"
                    task["completed_at"] = _utc_now()
                    task["error"] = str(exc)
                    for key in ("data", "document_text", "baseline_analysis", "claim_token", "claimed_at"):
                        task.pop(key, None)
            raise

    def _submit_session(self, payload):
        session = getattr(self.server, "dashboard_session", None)
        if not session:
            raise ValueError("No AI dashboard session is active")
        if payload.get("session_id") != session["id"]:
            raise ValueError("Dashboard session ID does not match")
        spec = payload.get("spec")
        if not isinstance(spec, dict):
            raise ValueError("A format specification is required")
        if spec.get("template_required") and spec.get("mode") != "official-template":
            raise ValueError("template_required is reserved for official-template workflows")
        if spec.get("mode") != "official-template":
            errors = validate_spec(spec)
            if errors:
                raise ValueError("; ".join(errors))
        lock = getattr(self.server, "dashboard_session_lock", None)
        if lock:
            with lock:
                return build_ai_handoff(
                    session,
                    payload,
                    skill_dir=SKILL_DIR,
                    registry=SUPPORTED_APPLICATION_PATHS,
                    dashboard_url=getattr(self.server, "dashboard_url", None),
                    renderers=detect_renderers(RUNTIME_ROOT),
                )
        return build_ai_handoff(
            session,
            payload,
            skill_dir=SKILL_DIR,
            registry=SUPPORTED_APPLICATION_PATHS,
            dashboard_url=getattr(self.server, "dashboard_url", None),
            renderers=detect_renderers(RUNTIME_ROOT),
        )

    def _record_visual_verification(self, payload):
        session = getattr(self.server, "dashboard_session", None)
        if not session:
            raise ValueError("No AI dashboard session is active")
        lock = getattr(self.server, "dashboard_session_lock", None)
        if lock:
            with lock:
                return record_visual_verification(session, payload)
        return record_visual_verification(session, payload)

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=0, help="Port; 0 selects an available port")
    parser.add_argument("--no-open", action="store_true", help="Do not open the default browser")
    parser.add_argument("--input", help="Preprocess this DOCX/DOTX and create an AI handoff session")
    parser.add_argument("--output", help="Requested output DOCX path for the AI handoff session")
    args = parser.parse_args()
    if not UI_DIR.is_dir() or not PRESETS_PATH.is_file():
        raise SystemExit("Dashboard assets are missing")
    server = ThreadingHTTPServer(("127.0.0.1", args.port), DashboardHandler)
    server.dashboard_session = prepare_dashboard_session(args.input, args.output) if args.input else None
    server.dashboard_session_lock = threading.Lock()
    server.dashboard_runtime_lock = threading.Lock()
    server.dashboard_runtime_job = {"status": "idle", "phase": "等待检测", "percent": 0}
    server.dashboard_analysis_tasks = {}
    server.dashboard_analysis_lock = threading.Lock()
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    server.dashboard_url = url
    print(f"Word Format Master: {url}", flush=True)
    if server.dashboard_session:
        print(f"AI dashboard session: {server.dashboard_session['id']}", flush=True)
    if not args.no_open:
        threading.Timer(0.25, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
