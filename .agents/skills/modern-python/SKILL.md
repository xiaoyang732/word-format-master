---
name: modern-python
description: Configures Python projects with modern tooling (uv, ruff, ty/mypy, pytest). Use when creating projects, writing scripts, configuring pyproject.toml, linting/formatting code, or migrating legacy tooling.
---

# Modern Python

Guide for modern Python tooling, standards, and best practices.

## When to Use This Skill
- Configuring or checking Python project standards and conventions (PEP 8, type hints, modern syntax).
- Setting up or maintaining `pyproject.toml` and development toolchains.
- Running or configuring Ruff for high-speed linting, formatting, and auto-fixing.
- Enforcing type safety with ty / mypy across modules and functions.
- Migrating legacy tooling (flake8, black, isort, setup.py, requirements.txt) to `uv` and `ruff`.

## Core Best Practices & Principles
1. **Dependency & Environment Management (`uv`)**:
   - Use `uv add <pkg>` and `uv remove <pkg>` to manage dependencies.
   - Use `[dependency-groups]` (PEP 735) for development/test dependencies.
   - Never manually activate virtual environments—execute commands via `uv run <cmd>`.
2. **Linting and Formatting (`ruff`)**:
   - Ruff replaces flake8, black, isort, pyupgrade, pydocstyle.
   - Format code with `uv run ruff format .`
   - Check and auto-fix with `uv run ruff check --fix .`
   - Target version: Python 3.10+ (modern typing `X | None`, `list[str]`, etc.).
3. **Type Hints & Static Safety**:
   - Strictly annotate public function arguments and return types.
   - Eliminate bare `Any` where concrete types or generic protocols can be defined.
   - Run type checks via `uv run ty check <path>` or `uv run mypy <path>`.
4. **Clean Code & Error Handling**:
   - Never use bare `except:`; specify caught exception classes.
   - Use f-strings for string formatting.
   - Guard against `None` and attribute errors before deep property access (especially in XML/DOM tree traversal).

## Quick Command Reference
| Tool / Task | Command |
| :--- | :--- |
| **Lint & Auto-fix** | `uv run ruff check --fix .` |
| **Code Formatting** | `uv run ruff format .` |
| **Check Format Only** | `uv run ruff format --check .` |
| **Type Checking** | `uv run ty check .` 或 `uv run mypy .` |
| **Add Dev Tool** | `uv add --group dev ruff pytest` |
| **Sync All Groups** | `uv sync --all-groups` |

## References
- [ruff-config.md](./references/ruff-config.md) - Complete Ruff linting and formatting rules & configuration.
- [uv-commands.md](./references/uv-commands.md) - `uv` package management and execution guide.
- [testing.md](./references/testing.md) - Pytest and coverage setup guide.
