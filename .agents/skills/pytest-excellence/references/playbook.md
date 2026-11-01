# pytest — Advanced Implementation Playbook

## §1 — Production Configuration

```toml
# pyproject.toml
[tool.pytest.ini_options]
testpaths = ["tests", "_dev/tests"]
addopts = "-v --strict-markers --tb=short"
markers = [
    "slow: marks tests as slow",
    "regression: marks regression tests",
    "integration: marks integration tests",
    "smoke: smoke tests",
]

[tool.coverage.run]
source = ["scripts", "src"]
omit = ["tests/*", "_dev/tests/*"]

[tool.coverage.report]
fail_under = 80
show_missing = true
```

## §2 — Fixtures (Scoping, Factories, Teardown)

```python
# conftest.py — shared fixtures
import pytest

# Function-scoped: created per test (default), auto-cleanup
@pytest.fixture
def sample_docx(tmp_path):
    from docx import Document
    doc = Document()
    doc.add_heading("测试标题", level=1)
    doc.add_paragraph("测试正文内容")
    file_path = tmp_path / "test.docx"
    doc.save(str(file_path))
    yield file_path
    # Teardown logic runs automatically after test

# Factory fixture — create multiple dynamic instances
@pytest.fixture
def docx_factory(tmp_path):
    created = []
    def _create_docx(name="sample.docx", paragraphs=3):
        from docx import Document
        doc = Document()
        for i in range(paragraphs):
            doc.add_paragraph(f"段落 {i+1}")
        file_path = tmp_path / name
        doc.save(str(file_path))
        created.append(file_path)
        return file_path
    yield _create_docx

# Autouse fixture — runs for every test in module
@pytest.fixture(autouse=True)
def reset_registry_cache():
    # Setup
    yield
    # Cleanup state
```

## §3 — Parameterized Tests

```python
import pytest

@pytest.mark.parametrize("heading_level,expected_style", [
    (1, "Heading 1"),
    (2, "Heading 2"),
    (3, "Heading 3"),
], ids=["h1", "h2", "h3"])
def test_heading_styles(heading_level, expected_style):
    # Arrange & Act
    style_name = get_heading_style(heading_level)
    # Assert
    assert style_name == expected_style
```

## §4 — Mocking with pytest-mock

```python
def test_external_api_call(mocker):
    mock_get = mocker.patch("requests.get")
    mock_get.return_value.status_code = 200
    mock_get.return_value.json.return_value = {"status": "ok"}
    
    response = fetch_remote_template("http://example.com")
    assert response["status"] == "ok"
    mock_get.assert_called_once()
```

## §5 — Best Practices Checklist

- ✅ Use fixtures over setup/teardown methods
- ✅ Use `conftest.py` for shared fixtures (auto-discovered)
- ✅ Use `tmp_path` for file operations (built-in, auto-cleanup)
- ✅ Use `monkeypatch` for env vars and attribute patching
- ✅ Use `pytest-mock` (`mocker` fixture) over `unittest.mock`
- ✅ Use `@pytest.mark.parametrize` for data-driven tests
- ✅ Register all custom markers in configuration
- ✅ Use `--strict-markers` to catch typos in marker names
- ✅ Use `--tb=short` for concise tracebacks
- ✅ Name files `test_*.py` and functions `test_*`
- ✅ Use `pytest.raises(match=...)` for precise error checking
