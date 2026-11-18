#!/usr/bin/env python3
"""Safely clean regenerable local development caches and build outputs.

Removes contents of ``dev/artifacts``, ``dev/cache``, ``dev/output``, ``dev/tmp``,
any stray ``__pycache__`` folders under the repository, and bounded legacy conversion
caches in ``%TEMP%/word_format_legacy_cache``, while strictly preserving ``dev/templates``,
``dev/examples``, ``dev/reports``, ``dev/docs``, ``dev/tests``, ``dev/scripts``,
``dev/tools``, and ``dev/ci``.
"""

from __future__ import annotations

import argparse
import json
import shutil
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DEV_DIR = ROOT / "dev"

REGENERABLE_DEV_DIRS = ("artifacts", "cache", "output", "tmp")
PROTECTED_DEV_DIRS = frozenset(
    {
        "ci",
        "docs",
        "examples",
        "local-config",
        "reports",
        "scripts",
        "templates",
        "tests",
        "tools",
    }
)


def _dir_bytes(path: Path) -> int:
    total = 0
    if not path.exists():
        return 0
    if path.is_file():
        return path.stat().st_size
    for item in path.rglob("*"):
        if item.is_file() and not item.is_symlink():
            try:
                total += item.stat().st_size
            except OSError:
                pass
    return total


def clean_dev(*, dry_run: bool = False, include_temp_cache: bool = True) -> dict:
    removed_paths: list[str] = []
    reclaimed_bytes = 0

    for name in REGENERABLE_DEV_DIRS:
        if name in PROTECTED_DEV_DIRS:
            raise RuntimeError(f"Refusing to clean protected dev directory: {name}")
        target = DEV_DIR / name
        if not target.exists():
            continue
        if DEV_DIR not in target.resolve().parents:
            raise RuntimeError(f"Unsafe target path outside dev/: {target}")
        reclaimed_bytes += _dir_bytes(target)
        removed_paths.append(target.relative_to(ROOT).as_posix())
        if not dry_run:
            shutil.rmtree(target)

    # Remove any stray __pycache__ directories across the repository.
    for pycache in sorted(ROOT.rglob("__pycache__")):
        if not pycache.is_dir():
            continue
        rel = pycache.relative_to(ROOT).as_posix()
        if any(rel.startswith(f"dev/{d}/") or rel == f"dev/{d}" for d in REGENERABLE_DEV_DIRS):
            continue
        reclaimed_bytes += _dir_bytes(pycache)
        removed_paths.append(rel)
        if not dry_run:
            shutil.rmtree(pycache)

    if include_temp_cache:
        temp_cache = Path(tempfile.gettempdir()) / "word_format_legacy_cache"
        if temp_cache.exists():
            reclaimed_bytes += _dir_bytes(temp_cache)
            removed_paths.append(str(temp_cache))
            if not dry_run:
                shutil.rmtree(temp_cache, ignore_errors=True)

    return {
        "status": "dry_run" if dry_run else "cleaned",
        "removed": removed_paths,
        "reclaimed_bytes": reclaimed_bytes,
        "reclaimed_mb": round(reclaimed_bytes / (1024 * 1024), 2),
        "protected_preserved": [
            name for name in sorted(PROTECTED_DEV_DIRS) if (DEV_DIR / name).exists()
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").replace("%", "%%"))
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Report what would be removed without deleting files",
    )
    parser.add_argument(
        "--keep-temp-cache",
        action="store_true",
        help="Keep %%TEMP%%/word_format_legacy_cache",
    )
    args = parser.parse_args()
    summary = clean_dev(dry_run=args.dry_run, include_temp_cache=not args.keep_temp_cache)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
