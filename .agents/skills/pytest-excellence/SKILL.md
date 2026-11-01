---
name: pytest-excellence
description: Generates production-grade pytest tests in Python with AAA pattern, fixtures, parametrize, markers, mocking, and conftest patterns. Use when writing, refactoring, or running Python tests.
---

# Pytest Excellence Testing Standards

## When to Use This Skill
- Writing unit, regression, or integration tests for Python codebases.
- Structuring tests with the Arrange-Act-Assert (AAA) pattern.
- Designing modular fixtures, factory fixtures, and lifecycle cleanup (`yield`).
- Running data-driven tests using `@pytest.mark.parametrize`.
- Mocking external calls and isolating file system tests via `tmp_path` and `pytest-mock`.

## Core Standards

### 1. Structure: Arrange-Act-Assert (AAA)
Always keep tests readable and segmented into three clear phases:
- **Arrange**: Prepare inputs, instantiate objects, set up fixtures or mocks.
- **Act**: Invoke the single method or function under test.
- **Assert**: Verify outcomes against expected values, state changes, or exception assertions.

### 2. Fixture Hygiene
- Place shared fixtures in `conftest.py` (auto-discovered).
- Use `yield` fixtures for teardown and cleanup; never leave dangling resources.
- Use built-in `tmp_path` for all document and file-system tests (guarantees isolation and auto-cleanup).
- Avoid global mutable test state; use function-scoped fixtures for test isolation.

### 3. Parametrization
- Use `@pytest.mark.parametrize` for data-driven test cases and boundary conditions.
- Provide descriptive `ids` for test variants to make test failure logs immediately understandable.

### 4. Mocking & Isolation
- Prefer `pytest-mock` (`mocker` fixture) over raw `unittest.mock.patch` for cleaner fixture lifecycle integration.
- Use `monkeypatch` for environment variables and configuration overrides.

### 5. Exception & Warning Testing
- Always verify the exception type and match the error message:
  ```python
  with pytest.raises(ValueError, match="Invalid style"):
      apply_formatting(bad_input)
  ```

## Anti-Patterns to Strictly Avoid
| Anti-Pattern | Recommended Practice | Reason |
| :--- | :--- | :--- |
| `self.assertEqual(...)` | `assert x == y` | Pytest assertion rewriting provides superior diff output. |
| Setup in `__init__` | `@pytest.fixture` | Proper lifecycle management. |
| Global state mutations | Fixtures with `yield` | Prevents test interference and flaky runs. |
| Massive, monolithic tests | Small, focused single-behavior tests | Easy root-cause identification when regressions occur. |
| Testing without asserts | Meaningful assertions on return and state | Avoids "coverage theater". |

## Quick Command Reference
| Task | Command |
| :--- | :--- |
| **Run All Tests** | `pytest` |
| **Run Verbose** | `pytest -v` |
| **Stop on First Failure** | `pytest -x` |
| **Run Last Failed** | `pytest --lf` |
| **Run by Marker** | `pytest -m "not slow"` |
| **Run Specific Test** | `pytest tests/test_core.py::test_specific_case` |
| **Coverage Report** | `pytest --cov=src --cov-report=term-missing` |

## References
- [playbook.md](./references/playbook.md) - 11 deep production patterns: configuration, fixtures, parametrize, mocking, async, class organization, debugging, and 15-item checklist.
