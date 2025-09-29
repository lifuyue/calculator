# Repository Guidelines

## Project Structure & Module Organization
- **`glycoenum/`** holds the Python package. Core modules: `types.py`, `formula_parser.py`, `mass_calc.py`, `enumerator.py`, `config.py`, and the Typer CLI in `cli.py`. Default unit data lives in `glycoenum/defaults/units.toml`.
- **`docs/`** contains architecture notes; add design RFCs or API drafts here. Keep any new contributor guides alongside `architecture.md`.
- Tests should mirror module layout under a future `tests/` package (e.g., `tests/test_formula_parser.py`). Avoid mixing fixtures with production code.

## Build, Test, and Development Commands
- `python3 -m venv .venv && source .venv/bin/activate` — standard virtualenv bootstrap.
- `pip install -e .[dev]` — install the package in editable mode once a `dev` extra is defined.
- `python -m glycoenum.cli run --help` — quick smoke-test of CLI wiring.
- `python -m compileall glycoenum` — lightweight syntax check already used in CI scaffolding; run before commits when you touch multiple modules.

## Coding Style & Naming Conventions
- Target Python 3.10+ with type hints. Prefer `dataclasses` for structured data and keep module-level constants UPPER_SNAKE.
- Formatting: adopt Black (88 cols) and Ruff for linting once configured. Use single quotes by default; reserve double quotes for docstrings.
- Filenames follow snake_case; Typer commands remain lowercase verbs (e.g., `run`).

## Testing Guidelines
- Use `pytest` with `pytest-cov` for coverage. Name test files `test_*.py` and focus on parsing edge cases, mass computations, and CLI integration.
- Add representative fixtures for common unit sets (default six units, custom modifier) to keep tests readable.
- Aim for >90% coverage in `glycoenum/formula_parser.py` and `glycoenum/mass_calc.py`, as they feed downstream calculations.

## Commit & Pull Request Guidelines
- Craft imperative commit subjects scoped to the domain: `Implement formula dehydration guard` or `Wire Typer CLI output formatter`.
- Pull requests should outline intent, list verification steps (`python -m compileall glycoenum`, future `pytest` runs), and mention config or data changes. Attach CLI output snippets or CSV samples when behavior changes.

## Security & Configuration Tips
- Never commit real `.env` files. When sharing CLI examples, redact customer data.
- Document new config keys in `docs/` and ensure defaults live in `glycoenum/defaults/` so the CLI stays reproducible.
