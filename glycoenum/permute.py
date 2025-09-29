"""Permutation utilities for enumerating unique glycan sequences."""

from __future__ import annotations

from math import factorial
from typing import Iterator, Mapping, Sequence, Tuple


def permutation_count(multiset: Mapping[str, int]) -> int:
    """Return the number of unique permutations for the given multiset."""
    counts = [int(v) for v in multiset.values() if int(v) > 0]
    total = sum(counts)
    if total == 0:
        return 0
    numerator = factorial(total)
    denominator = 1
    for count in counts:
        denominator *= factorial(count)
    return numerator // denominator


def iter_unique_permutations(multiset: Mapping[str, int]) -> Iterator[Tuple[str, ...]]:
    """Yield unique permutations for the supplied multiset in lexicographic order."""
    items: list[tuple[str, int]] = [
        (label, int(count)) for label, count in sorted(multiset.items()) if int(count) > 0
    ]
    if not items:
        return

    labels = [label for label, _ in items]
    counts = [count for _, count in items]
    total_length = sum(counts)
    buffer: list[str] = [""] * total_length

    def backtrack(depth: int) -> Iterator[Tuple[str, ...]]:
        if depth == total_length:
            yield tuple(buffer)
            return
        for idx, label in enumerate(labels):
            if counts[idx] == 0:
                continue
            counts[idx] -= 1
            buffer[depth] = label
            yield from backtrack(depth + 1)
            counts[idx] += 1

    yield from backtrack(0)
