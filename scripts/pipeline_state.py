"""Hash-bound stage checkpoints and a process-scoped, nonblocking task lock."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import uuid

from json_output import write_json


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def now():
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def task_lock(path):
    # Keep the file: unlinking a lock file races with another process opening it.
    with Path(path).open("a+b") as stream:
        stream.seek(0, 2)
        if stream.tell() == 0:
            stream.write(b"0")
            stream.flush()
        stream.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise ValueError("此任务正在由另一个进程执行，请等待后恢复") from exc
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


class PipelineState:
    def __init__(self, path, identity, resume):
        self.path = Path(path)
        if resume:
            if not self.path.is_file():
                raise ValueError("缺少流水线状态；请选择新的任务目录重新执行")
            self.data = json.loads(self.path.read_text(encoding="utf-8"))
            if self.data.get("schema_version") != "1.0" or self.data.get("identity") != identity:
                raise ValueError("输入、参数或引擎摘要已变化，不能复用此任务状态；请另建任务目录")
        else:
            if self.path.exists():
                raise ValueError("任务状态已存在，请使用 --resume 或选择新的任务目录")
            self.data = {"schema_version": "1.0", "task_id": uuid.uuid4().hex,
                         "identity": identity, "stages": {}, "created_at": now()}
            self.save()

    def save(self):
        self.data["updated_at"] = now()
        write_json(self.path, self.data)

    def reusable(self, name, paths):
        stage = self.data["stages"].get(name, {})
        if stage.get("status") != "succeeded":
            return False
        expected = {str(Path(p).resolve()) for p in paths}
        artifacts = stage.get("artifacts", {})
        if set(artifacts) != expected or any(not Path(p).is_file() or digest(p) != value for p, value in artifacts.items()):
            downstream = False
            for key, value in self.data["stages"].items():
                if key == name:
                    downstream = True
                if downstream:
                    value["status"] = "invalidated"
            self.data["failure"] = {"stage": name, "code": "EVIDENCE_STALE",
                                    "detail": "阶段产物或证据已缺失或变化，需要重新核对"}
            self.save()
            raise ValueError(f"{name} 阶段产物或证据摘要不一致，拒绝恢复；请另建任务目录")
        return True

    def start(self, name):
        old = self.data["stages"].get(name, {})
        self.data["stages"][name] = {"status": "running", "attempts": old.get("attempts", 0) + 1,
                                    "started_at": now()}
        self.data.pop("failure", None)
        self.save()

    def succeeded(self, name, paths):
        self.data["stages"][name].update(status="succeeded", finished_at=now(),
                                          artifacts={str(Path(p).resolve()): digest(p) for p in paths})
        self.save()

    def failed(self, name, exc):
        self.data["stages"][name].update(status="failed", finished_at=now())
        self.data["failure"] = {"stage": name, "detail": str(exc), "code": "STAGE_FAILED"}
        self.save()
