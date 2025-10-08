# Oligosaccharide Prediction

Oligosaccharide Prediction is a desktop application for enumerating glycan unit permutations, applying dehydration and terminal modifiers, and visualising the resulting molecular formulas alongside theoretical masses. The project now ships with a native GUI instead of the previous CLI workflow.

## Getting Started

```bash
python -m venv .venv
.venv\Scripts\activate  # PowerShell: .\.venv\Scripts\Activate.ps1
pip install -e .
```

Launch the interface with any of the following commands:

```bash
python -m glycoenum
python -m glycoenum.gui
oligosaccharide-prediction
```

## Using the Application

- Enter non-negative counts for the six supported monosaccharide units (Hex, deoxyhex, pent, HexN, UA, HexNAc).
- Click **Calculate sequences** to enumerate permutations, compute dehydrated and modified formulas, and report the theoretical mass using the CLI defaults (monoisotopic masses, neutral adduct, four decimal places).
- Use **Export XLSX...** to save the currently displayed permutations to an Excel workbook with predicted compound, pre-/post-derivatization molecular formulas, calculated mass, and theoretical m/z columns.
- Select **Generate summary...** to build the full 2-10 unit permutation summary workbooks on demand; choose an output directory and the app will write chunked XLSX files alongside a manifest using the same column schema.

Totals must satisfy 2 <= sum(units) <= 10. Each calculation removes (n - 1) * H2O and applies a single 2 PMP terminal modifier (C20H18N4O).

## Building

The repository includes a PyInstaller spec for bundling a standalone executable:

```bash
pyinstaller --noconfirm --onefile --windowed glycoenum/gui.py -n OligosaccharidePrediction
```

On Windows you can run `build.bat`; the script installs dependencies and emits `dist\Oligosaccharide prediction.exe`.

## Development Notes

- Core composition math lives in `glycoenum/formula.py`, mass handling in `glycoenum/mass.py`, and permutation streaming in `glycoenum/permute.py`.
- Tests belong under `tests/`, mirroring module structure (`tests/test_gui.py`, `tests/test_formula.py`, etc.).
- When publishing changes to mass tables or modifiers, update the README and flag any packaging adjustments so downstream consumers can refresh PyInstaller bundles.
