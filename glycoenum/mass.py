"""Mass calculation helpers for the glycoenum CLI."""

from __future__ import annotations

from typing import Mapping

DEFAULT_MASS_TABLES: dict[str, dict[str, float]] = {
    "monoisotopic": {
        "C": 12.0,
        "H": 1.00782503223,
        "N": 14.00307400443,
        "O": 15.99491461957,
        "Na": 22.9897692820,
    },
    "average": {
        "C": 12.0107,
        "H": 1.00794,
        "N": 14.0067,
        "O": 15.9994,
        "Na": 22.98976928,
    },
}


def build_mass_table(
    model: str,
    overrides: Mapping[str, float] | None = None,
) -> dict[str, float]:
    """Return the atom-mass table for *model* with optional overrides."""
    key = model.strip().lower()
    if key not in DEFAULT_MASS_TABLES:
        raise ValueError(f"Unknown mass model '{model}'")

    table = {element: float(value) for element, value in DEFAULT_MASS_TABLES[key].items()}
    if overrides:
        for element, value in overrides.items():
            symbol = element[0].upper() + element[1:].lower()
            table[symbol] = float(value)
    return table


def calculate_mass(counts: Mapping[str, int], masses: Mapping[str, float]) -> float:
    """Compute the molecular mass for the provided composition."""
    total = 0.0
    for element, amount in counts.items():
        if element not in masses:
            raise KeyError(f"Missing mass for element '{element}'")
        total += masses[element] * int(amount)
    return total


def apply_adduct(base_mass: float, adduct: str, masses: Mapping[str, float]) -> float:
    """Apply a supported adduct expression to *base_mass*."""
    text = (adduct or "").strip()
    if not text or text.lower() == "neutral":
        return base_mass

    normalized = text.lower()
    if normalized == "[m+h]+":
        if "H" not in masses:
            raise KeyError("Mass table missing 'H' required for [M+H]+ adduct")
        return base_mass + masses["H"]
    if normalized == "[m+na]+":
        if "Na" not in masses:
            raise KeyError("Mass table missing 'Na' required for [M+Na]+ adduct")
        return base_mass + masses["Na"]

    raise ValueError(f"Unsupported adduct '{adduct}'")
