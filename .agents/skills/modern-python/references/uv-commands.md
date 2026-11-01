# uv Command Reference

`uv` is an extremely fast Python package and project manager written in Rust. It replaces pip, virtualenv, pip-tools, pipx, and pyenv.

**Key principle:** Always use `uv run` to execute commands. Never manually activate virtual environments.

## Common Project Commands

| Command | Description |
|---------|-------------|
| `uv add <pkg>` | Add dependency to project |
| `uv add <pkg> --group dev` | Add to development dependency group |
| `uv remove <pkg>` | Remove dependency |
| `uv sync` | Install dependencies (creates/updates venv) |
| `uv sync --all-groups` | Install all dependency groups |
| `uv run <cmd>` | Run command in project venv |
| `uv run python script.py` | Run Python script within environment |
| `uv run pytest` | Run pytest suite |
| `uv run ruff check .` | Run Ruff linter |
| `uv run ruff format .` | Run Ruff code formatter |
