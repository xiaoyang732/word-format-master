# Testing with pytest

Configuration and best practices for pytest with coverage enforcement.

## pyproject.toml Configuration

```toml
[tool.pytest.ini_options]
testpaths = ["tests", "_dev/tests"]
addopts = [
    "-ra",
    "--strict-markers",
    "--tb=short",
    "-v"
]
markers = [
    "slow: marks tests as slow",
    "regression: marks regression tests",
    "integration: marks integration tests",
]
```

## Running Tests

```bash
# Run all tests
uv run pytest

# Run with verbose output
uv run pytest -v

# Run specific test file
uv run pytest _dev/tests/test_format_core.py

# Stop on first failure
uv run pytest -x

# Run only failed tests from last run
uv run pytest --lf
```
