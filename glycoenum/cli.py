from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter
from pathlib import Path
from typing import Iterable

from glycoenum import __version__
from glycoenum.formula import add_modifier, dehydrate, format_hill, parse_formula, scale_counts
from glycoenum.mass import apply_adduct, build_mass_table, calculate_mass
from glycoenum.permute import iter_unique_permutations, permutation_count

UNIT_ORDER = ["Hex", "deoxyhex", "pent", "HexN", "UA", "HexNAc"]
UNIT_OPTION_DESTS = {
    "Hex": "hex",
    "deoxyhex": "deoxyhex",
    "pent": "pent",
    "HexN": "hexn",
    "UA": "ua",
    "HexNAc": "hexnac",
}
UNIT_FORMULAS = {
    "Hex": "C6H12O6",
    "deoxyhex": "C6H12O5",
    "pent": "C5H10O5",
    "HexN": "C6H13NO5",
    "UA": "C6H10O7",
    "HexNAc": "C8H15NO6",
}
TERMINAL_MODIFIER = "C20H18N4O"
ADDUCT_CHOICES = ["neutral", "[M+H]+", "[M+Na]+"]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="glycoenum",
        description=(
            "Enumerate glycan sequence permutations, compute dehydrated and modified "
            "formulas, and report theoretical masses."
        ),
    )
    parser.add_argument(
        "counts",
        nargs="*",
        type=int,
        help="Positional form: Hex deoxyhex pent HexN UA HexNAc (non-negative integers).",
    )
    parser.add_argument(
        "--hex",
        dest="hex",
        type=int,
        help="Count for Hex units.",
    )
    parser.add_argument(
        "--deoxyhex",
        dest="deoxyhex",
        type=int,
        help="Count for deoxyhex units.",
    )
    parser.add_argument(
        "--pent",
        dest="pent",
        type=int,
        help="Count for pent units.",
    )
    parser.add_argument(
        "--hexn",
        dest="hexn",
        type=int,
        help="Count for HexN units.",
    )
    parser.add_argument(
        "--ua",
        dest="ua",
        type=int,
        help="Count for UA units.",
    )
    parser.add_argument(
        "--hexnac",
        dest="hexnac",
        type=int,
        help="Count for HexNAc units.",
    )
    parser.add_argument(
        "--adduct",
        default="neutral",
        help="Adduct mode: neutral (default), [M+H]+, [M+Na]+.",
    )
    parser.add_argument(
        "--mass-model",
        dest="mass_model",
        default="monoisotopic",
        choices=["monoisotopic", "average"],
        help="Mass model to use (monoisotopic or average).",
    )
    parser.add_argument(
        "--masses",
        help="Override atomic masses, e.g. C=12.0,H=1.007825.",
    )
    parser.add_argument(
        "--decimals",
        type=int,
        default=4,
        help="Decimal places for the theoretical mass column (default: 4).",
    )
    parser.add_argument(
        "--csv",
        help="Write output CSV to this path; otherwise emit to stdout.",
    )
    parser.add_argument(
        "--max-rows",
        dest="max_rows",
        type=int,
        help="Optional row cap; stop early and warn if exceeded.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"glycoenum {__version__}",
    )
    return parser


def main(argv: Iterable[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not _has_any_counts(args):
        if sys.stdin.isatty():
            print('请输入六个非负整数，对应 Hex, deoxyhex, pent, HexN, UA, HexNAc。', file=sys.stdout)
            print('使用空格分隔多个数值后按 Enter 确认（例如：3 1 0 2 0 0）。按 Ctrl+C 退出。', file=sys.stdout)
            print('', file=sys.stdout)
            args.counts = _prompt_positional_counts()
        else:
            parser.print_help()
            print('\nProvide counts using positional arguments (Hex deoxyhex pent HexN UA HexNAc) or named options such as --hex.', file=sys.stdout)
            return

    try:
        counts_map = _resolve_counts(args, parser)
        overrides = _parse_mass_overrides(args.masses)
        decimals = _validate_decimals(args.decimals)
        max_rows = _validate_max_rows(args.max_rows)
        adduct = _normalize_adduct(args.adduct)
    except ValueError as exc:
        print(exc, file=sys.stderr)
        sys.exit(1)

    total_units = sum(counts_map.values())
    if not 2 <= total_units <= 10:
        print("Total unit count must be between 2 and 10 (inclusive).", file=sys.stderr)
        sys.exit(1)

    unit_compositions = {
        name: scale_counts(parse_formula(formula), counts_map[name])
        for name, formula in UNIT_FORMULAS.items()
    }

    pooled = Counter()
    for composition in unit_compositions.values():
        pooled.update(composition)

    dehydrated = dehydrate(pooled, total_units)
    try:
        final_counts = add_modifier(dehydrated, TERMINAL_MODIFIER)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    base_formula = format_hill(dehydrated)
    final_formula = format_hill(final_counts)

    try:
        masses = build_mass_table(args.mass_model, overrides)
        neutral_mass = calculate_mass(final_counts, masses)
        theoretical_mass = apply_adduct(neutral_mass, adduct, masses)
    except (ValueError, KeyError) as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    permutations = {name: count for name, count in counts_map.items() if count > 0}
    total_permutations = permutation_count(permutations)

    row_limit = max_rows if max_rows is not None else total_permutations
    cap = min(row_limit, total_permutations)

    destination = Path(args.csv) if args.csv else None
    try:
        rows_written = _emit_output(
            permutations,
            base_formula,
            final_formula,
            theoretical_mass,
            decimals,
            cap,
            destination,
        )
    except OSError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    if cap < total_permutations:
        message = (
            f"[warn] rows truncated at {cap} (total would be {total_permutations})."
        )
        print(message, file=sys.stderr)

    if destination is None:
        sys.stdout.flush()
        _pause_before_exit()


def _has_any_counts(args: argparse.Namespace) -> bool:
    if args.counts:
        return True
    for dest in UNIT_OPTION_DESTS.values():
        if getattr(args, dest) is not None:
            return True
    return False


def _pause_before_exit() -> None:
    if not getattr(sys, 'frozen', False):
        return
    try:
        input('按 Enter 键退出...')
    except EOFError:
        pass


def _prompt_positional_counts() -> list[int]:
    while True:
        try:
            raw = input('请输入六个非负整数，以空格分隔（Hex deoxyhex pent HexN UA HexNAc）：').strip()
        except EOFError:
            raise SystemExit(1)
        if not raw:
            print('输入不能为空，请重新输入。', file=sys.stdout)
            continue
        parts = raw.replace(',', ' ').split()
        if len(parts) != len(UNIT_ORDER):
            print(f'需要 {len(UNIT_ORDER)} 个数值，请重新输入。', file=sys.stdout)
            continue
        try:
            values = [int(part) for part in parts]
        except ValueError:
            print('请只输入非负整数。', file=sys.stdout)
            continue
        if any(value < 0 for value in values):
            print('所有数值必须为非负整数。', file=sys.stdout)
            continue
        print('', file=sys.stdout)
        return values


def _resolve_counts(args: argparse.Namespace, parser: argparse.ArgumentParser) -> dict[str, int]:
    counts_map: dict[str, int] = {}
    positional = list(args.counts)
    if positional:
        if len(positional) != len(UNIT_ORDER):
            parser.error(
                "Exactly six positional counts required: Hex deoxyhex pent HexN UA HexNAc"
            )
        counts_map.update(dict(zip(UNIT_ORDER, positional)))

    for name in UNIT_ORDER:
        dest = UNIT_OPTION_DESTS[name]
        option_value = getattr(args, dest)
        if option_value is None:
            continue
        if name in counts_map and counts_map[name] != option_value:
            parser.error(f"Conflicting counts for {name} (positional vs option)")
        counts_map[name] = option_value

    missing = [name for name in UNIT_ORDER if name not in counts_map]
    if missing:
        parser.error(
            "Missing counts for: "
            + ", ".join(missing)
            + ". Provide either positional counts or named options."
        )

    for name, value in counts_map.items():
        if value is None:
            parser.error(f"Count for {name} must be an integer")
        if value < 0:
            parser.error(f"Counts must be non-negative. {name}={value}")

    return {name: int(counts_map[name]) for name in UNIT_ORDER}


def _parse_mass_overrides(text: str | None) -> dict[str, float]:
    if not text:
        return {}
    overrides: dict[str, float] = {}
    entries = [segment.strip() for segment in text.split(",") if segment.strip()]
    if not entries:
        raise ValueError("Mass override string is empty")
    for entry in entries:
        if "=" not in entry:
            raise ValueError(
                f"Mass override '{entry}' must be in ELEMENT=value format"
            )
        element, value = entry.split("=", 1)
        symbol = element.strip()
        if not symbol:
            raise ValueError("Element symbol cannot be empty in overrides")
        overrides[symbol[0].upper() + symbol[1:].lower()] = float(value)
    return overrides


def _validate_decimals(value: int) -> int:
    if value < 0:
        raise ValueError("--decimals must be zero or a positive integer")
    return int(value)


def _validate_max_rows(value: int | None) -> int | None:
    if value is None:
        return None
    if value <= 0:
        raise ValueError("--max-rows must be a positive integer")
    return int(value)


def _normalize_adduct(adduct: str) -> str:
    normalized = (adduct or "").strip()
    if not normalized:
        return "neutral"
    lowered = normalized.lower()
    for choice in ADDUCT_CHOICES:
        if lowered == choice.lower():
            return choice
    raise ValueError(
        "--adduct must be one of neutral, [M+H]+, [M+Na]+ (case insensitive)."
    )


def _emit_output(
    permutations: dict[str, int],
    base_formula: str,
    final_formula: str,
    theoretical_mass: float,
    decimals: int,
    cap: int,
    destination: Path | None,
) -> int:
    rows = 0
    header = ["compound", "分子式", "最终分子式", "理论"]

    if destination:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with destination.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(header)
            rows = _write_rows(writer, permutations, base_formula, final_formula, theoretical_mass, decimals, cap)
    else:
        writer = csv.writer(sys.stdout, lineterminator="\n")
        writer.writerow(header)
        rows = _write_rows(writer, permutations, base_formula, final_formula, theoretical_mass, decimals, cap)
    return rows


def _write_rows(
    writer: csv.writer,
    permutations: dict[str, int],
    base_formula: str,
    final_formula: str,
    theoretical_mass: float,
    decimals: int,
    cap: int,
) -> int:
    formatted_mass = f"{theoretical_mass:.{decimals}f}"
    emitted = 0
    if cap <= 0:
        return emitted
    for emitted, sequence in enumerate(iter_unique_permutations(permutations), start=1):
        writer.writerow(
            [
                "-".join(sequence),
                base_formula,
                final_formula,
                formatted_mass,
            ]
        )
        if emitted >= cap:
            break
    return emitted


if __name__ == "__main__":  # pragma: no cover
    main()
