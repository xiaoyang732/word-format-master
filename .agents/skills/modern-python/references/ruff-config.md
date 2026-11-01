# Ruff Configuration Reference

Ruff is an extremely fast Python linter and formatter written in Rust. It replaces flake8, black, isort, pyupgrade, pydocstyle, and many other tools.

## Basic Setup

Add to `pyproject.toml`:

```toml
[tool.ruff]
line-length = 100
target-version = "py311"
src = ["src", "scripts"]

[tool.ruff.lint]
select = ["ALL"]
ignore = [
    "D",        # pydocstyle
    "COM812",   # trailing comma (formatter conflict)
    "ISC001",   # string concat (formatter conflict)
]

[tool.ruff.format]
quote-style = "double"
indent-style = "space"
docstring-code-format = true
```

## Running Ruff

```bash
# Lint
uv run ruff check .
uv run ruff check --fix .        # Auto-fix
uv run ruff check --fix --unsafe-fixes .  # Including unsafe fixes

# Format
uv run ruff format .
uv run ruff format --check .     # Check only
uv run ruff format --diff .      # Show diff
```

## Rule Categories

Using `select = ["ALL"]` enables all rules. Common categories:

| Code | Category | Description |
|------|----------|-------------|
| `E`, `W` | pycodestyle | Style errors and warnings |
| `F` | Pyflakes | Logical errors |
| `I` | isort | Import sorting |
| `N` | pep8-naming | Naming conventions |
| `D` | pydocstyle | Docstring conventions |
| `UP` | pyupgrade | Python upgrade suggestions |
| `B` | flake8-bugbear | Bug detection |
| `S` | flake8-bandit | Security issues |
| `A` | flake8-builtins | Built-in shadowing |
| `C4` | flake8-comprehensions | Comprehension improvements |
| `DTZ` | flake8-datetimez | Timezone-aware datetime |
| `T10` | flake8-debugger | Debugger statements |
| `T20` | flake8-print | Print statements |
| `PT` | flake8-pytest-style | Pytest style |
| `Q` | flake8-quotes | Quote consistency |
| `SIM` | flake8-simplify | Simplification suggestions |
| `TID` | flake8-tidy-imports | Import hygiene |
| `ARG` | flake8-unused-arguments | Unused arguments |
| `ERA` | eradicate | Commented-out code |
| `PL` | Pylint | Pylint rules |
| `RUF` | Ruff-specific | Ruff's own rules |
| `ANN` | flake8-annotations | Type annotation checks |

## Recommended Ignores

### Always Ignore (Formatter Conflicts)

```toml
ignore = [
    "COM812",   # missing-trailing-comma
    "ISC001",   # single-line-implicit-string-concatenation
]
```

### Common Ignores

```toml
ignore = [
    "D",        # Docstrings (enable selectively)
    "ANN401",   # Dynamically typed Any
    "TD002",    # Missing TODO author
    "TD003",    # Missing TODO link
    "FIX002",   # Line contains TODO
]
```

## Formatter Configuration

```toml
[tool.ruff.format]
quote-style = "double"           # or "single"
indent-style = "space"           # or "tab"
skip-magic-trailing-comma = false
line-ending = "auto"             # or "lf", "crlf"
docstring-code-format = true
docstring-code-line-length = 80
```
