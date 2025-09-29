"""Core formula utilities for the glycoenum CLI.

The helpers in this module parse molecular formulas, combine element counts,
apply dehydration offsets, and format counts using Hill notation. All
functions return new dictionaries and never mutate caller-owned mappings.
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Mapping

_ELEMENT_RE = re.compile(r"([A-Za-z][a-z]?)(\d*)")


def parse_formula(formula: str) -> dict[str, int]:
    """Parse a molecular formula string into an element → count mapping."""
    if formula is None:
        raise ValueError("Formula cannot be None")

    cleaned = "".join(str(formula).split())
    if not cleaned:
        raise ValueError("Formula cannot be empty")

    pos = 0
    counts: Counter[str] = Counter()
    while pos < len(cleaned):
        match = _ELEMENT_RE.match(cleaned, pos)
        if not match:
            snippet = cleaned[pos : pos + 5]
            raise ValueError(f"Invalid token '{snippet}' in formula '{formula}'")
        symbol = _canonical_symbol(match.group(1))
        number_group = match.group(2)
        amount = int(number_group) if number_group else 1
        if amount < 0:
            raise ValueError("Element counts must be non-negative integers")
        counts[symbol] += amount
        pos = match.end()

    return _strip_zeros(dict(counts))


def scale_counts(counts: Mapping[str, int], factor: int) -> dict[str, int]:
    """Multiply all element counts by *factor* (must be non-negative)."""
    if factor < 0:
        raise ValueError("Scale factor must be non-negative")
    result: dict[str, int] = {}
    for element, value in counts.items():
        _validate_count(element, value)
        result[_canonical_symbol(element)] = int(value) * factor
    return _strip_zeros(result)


def dehydrate(counts: Mapping[str, int], n: int) -> dict[str, int]:
    """Remove (n−1) molecules of H₂O from the provided composition."""
    if n < 1:
        raise ValueError("Polymer length must be at least 1")

    adjusted = Counter({
        _canonical_symbol(elem): int(amount)
        for elem, amount in counts.items()
    })

    if n == 1:
        return _strip_zeros(dict(adjusted))

    loss = n - 1
    adjusted["H"] -= 2 * loss
    adjusted["O"] -= loss

    _ensure_non_negative(adjusted, context="after dehydration")
    return _strip_zeros(dict(adjusted))


def add_modifier(counts: Mapping[str, int], modifier: str) -> dict[str, int]:
    """Add a modifier formula exactly once to the given composition."""
    result = Counter({
        _canonical_symbol(elem): int(amount)
        for elem, amount in counts.items()
    })
    result.update(parse_formula(modifier))
    _ensure_non_negative(result, context="after modifier addition")
    return _strip_zeros(dict(result))


def format_hill(counts: Mapping[str, int]) -> str:
    """Format element counts according to Hill notation."""
    normalized = {
        _canonical_symbol(elem): int(amount)
        for elem, amount in counts.items()
        if int(amount) != 0
    }
    if not normalized:
        return "0"

    def sort_key(element: str) -> tuple[int, str]:
        if element == "C":
            return (0, element)
        if element == "H":
            return (1, element)
        return (2, element)

    parts: list[str] = []
    for element in sorted(normalized, key=sort_key):
        value = normalized[element]
        parts.append(element if value == 1 else f"{element}{value}")
    return "".join(parts)


def _canonical_symbol(symbol: str) -> str:
    if not symbol:
        raise ValueError("Element symbol cannot be empty")
    return symbol[0].upper() + symbol[1:].lower()


def _strip_zeros(data: Mapping[str, int]) -> dict[str, int]:
    return {k: int(v) for k, v in data.items() if int(v) != 0}


def _ensure_non_negative(data: Mapping[str, int], *, context: str) -> None:
    negatives = {k: v for k, v in data.items() if v < 0}
    if negatives:
        details = ", ".join(f"{elem}={val}" for elem, val in negatives.items())
        raise ValueError(f"Negative counts {context}: {details}")


def _validate_count(element: str, value: int) -> None:
    if int(value) < 0:
        raise ValueError(f"Negative count for element '{element}' is not allowed")
