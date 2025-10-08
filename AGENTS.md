# Repository Guidelines

## Project Structure & Module Organization
- `glycoenum/` contains the application package: `gui.py` hosts the Tkinter interface, `cli.py` simply forwards legacy entry points to the GUI, `formula.py` handles composition math, `mass.py` covers mass models and adducts, `permute.py` streams multiset permutations, and `__main__.py` exposes `python -m glycoenum`.
- Root-level `pyproject.toml` defines the entry point and build metadata; `build.bat` wraps the PyInstaller flow. Place any future docs under `docs/` and Python tests beside the package in `tests/` (e.g., `tests/test_formula.py`).
- Runtime assets are minimal—no user-writable data lives in the package. Keep optional examples or CSV fixtures in `docs/` to avoid shipping them with the library.

## Build, Test, and Development Commands
- `python3 -m venv .venv && source .venv/bin/activate` – standard virtualenv bootstrap.
- `pip install -e .` – install in editable mode for local iteration.
- `python3 -m glycoenum.gui` – smoke-check GUI launch from the module entry point.
- `python3 -m compileall glycoenum` – quick syntax verification before commits.
- Packaging: `pyinstaller --noconfirm --onefile --windowed glycoenum/gui.py -n OligosaccharidePrediction` (mirrors the documented release command).

## Coding Style & Naming Conventions
- Target Python 3.10+, prefer type hints on public functions. Constants remain UPPER_SNAKE; module names use snake_case.
- Follow Black (88 cols) and Ruff defaults once configured; align string quoting with double quotes for user-facing text and docstrings.
- CLI flags stay lowercase with hyphen separators (`--mass-model`, `--max-rows`). Function names should be verbs (`format_hill`, `iter_unique_permutations`).

## Testing Guidelines
- Adopt `pytest` for unit and GUI integration tests. Mirror module paths: `tests/test_permutes.py`, `tests/test_gui.py`, etc.
- Cover parsing edge cases (mixed case symbols, invalid tokens), mass override errors, adduct math, and truncation warnings. Include at least one golden CSV snapshot test using temporary files.
- Aim for ≥90% line coverage in `formula.py` and `mass.py`; document unavoidable gaps in test docstrings.

## Commit & Pull Request Guidelines
- Commits use imperative subjects scoped to the domain (`Add multiset stream generator`, `Refine mass override errors`). Group related refactors into a single logical change.
- PRs must describe intent, list verification steps (e.g., `python3 -m compileall glycoenum`, planned `pytest` suite), and attach relevant CLI output snippets. Note changes to packaging scripts or default constants so downstream consumers can react.

## Security & Packaging Notes
- Do not commit generated executables or virtualenv folders. Example outputs should omit PII and rely on synthetic compositions.
- When adjusting mass tables or modifiers, update the README and mention PyInstaller rebuild steps so release automation stays predictable.
