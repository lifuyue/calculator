from __future__ import annotations

import argparse
import csv
import json
import sys
import tempfile
from collections import Counter
from itertools import chain, islice
from pathlib import Path
from typing import Iterable, Iterator, TextIO
import zipfile

from xml.sax.saxutils import escape

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
OUTPUT_HEADER = ["compound", "分子式", "最终分子式", "理论"]
XLSX_FILENAME = "glycoenum_output.xlsx"
SUMMARY_BASENAME = "glycoenum_summary"
SUMMARY_MANIFEST_NAME = f"{SUMMARY_BASENAME}_manifest.json"
SUMMARY_MIN_TOTAL_UNITS = 2
SUMMARY_MAX_TOTAL_UNITS = 10
SUMMARY_DECIMALS = 4
SUMMARY_ROWS_PER_WORKBOOK = 1_048_575


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

    _maybe_generate_summary_workbook()

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

    mass_text = f"{theoretical_mass:.{decimals}f}"

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
            mass_text,
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

    _maybe_generate_xlsx(
        permutations,
        base_formula,
        final_formula,
        mass_text,
        rows_written,
        total_permutations,
    )

    if destination is None:
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
    mass_text: str,
    cap: int,
    destination: Path | None,
) -> int:
    rows = 0

    if destination:
        destination.parent.mkdir(parents=True, exist_ok=True)
        with destination.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(OUTPUT_HEADER)
            rows = _write_rows(
                writer,
                permutations,
                base_formula,
                final_formula,
                mass_text,
                cap,
            )
    else:
        writer = csv.writer(sys.stdout, lineterminator="\n")
        writer.writerow(OUTPUT_HEADER)
        rows = _write_rows(
            writer,
            permutations,
            base_formula,
            final_formula,
            mass_text,
            cap,
        )
    return rows


def _write_rows(
    writer: csv.writer,
    permutations: dict[str, int],
    base_formula: str,
    final_formula: str,
    mass_text: str,
    cap: int,
) -> int:
    emitted = 0
    if cap <= 0:
        return emitted
    for emitted, sequence in enumerate(iter_unique_permutations(permutations), start=1):
        writer.writerow(
            [
                "-".join(sequence),
                base_formula,
                final_formula,
                mass_text,
            ]
        )
        if emitted >= cap:
            break
    return emitted


def _maybe_generate_xlsx(
    permutations: dict[str, int],
    base_formula: str,
    final_formula: str,
    mass_text: str,
    rows_written: int,
    total_permutations: int,
) -> None:
    if rows_written <= 0:
        return
    if not sys.stdin.isatty():
        return
    try:
        answer = input("是否在当前目录生成 XLSX 表格？(Y/N): ").strip().lower()
    except EOFError:
        return
    if answer not in {"y", "yes"}:
        return

    target = Path.cwd() / XLSX_FILENAME
    try:
        exported = _write_xlsx(
            permutations,
            base_formula,
            final_formula,
            mass_text,
            rows_written,
            target,
        )
    except (OSError, ValueError) as exc:
        print(f"写入 XLSX 失败: {exc}", file=sys.stderr)
        return

    note = ""
    if total_permutations > rows_written:
        note = f"（已截断，仅导出前 {rows_written} 行）"
    print(
        f"[info] 已在当前目录生成 {target.name}，共 {exported} 行{note}",
        file=sys.stderr,
    )


def _maybe_generate_summary_workbook() -> None:
    target_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
    manifest_path = target_dir / SUMMARY_MANIFEST_NAME

    manifest = _read_summary_manifest(manifest_path)
    if manifest and _manifest_complete(manifest, target_dir):
        files = manifest.get("files", [])
        total_rows = manifest.get("total_rows", 0)
        print(
            f"[info] 检测到已存在 {len(files)} 个总表文件（共 {total_rows} 行），跳过生成。",
            file=sys.stderr,
        )
        return

    # 清理旧文件，准备重新生成
    for leftover in target_dir.glob(f"{SUMMARY_BASENAME}*.xlsx"):
        try:
            leftover.unlink()
        except OSError:
            pass
    try:
        manifest_path.unlink()
    except OSError:
        pass

    total_expected = sum(6 ** n for n in range(SUMMARY_MIN_TOTAL_UNITS, SUMMARY_MAX_TOTAL_UNITS + 1))
    print(
        (
            f"[info] 正在生成 {SUMMARY_MIN_TOTAL_UNITS}-{SUMMARY_MAX_TOTAL_UNITS} 单元所有排列组合，总计"
            f" {total_expected} 行数据，可能耗时较长..."
        ),
        file=sys.stderr,
    )

    try:
        manifest = _build_summary_workbooks(target_dir)
    except (OSError, ValueError) as exc:
        print(f"[warn] 总表生成失败: {exc}", file=sys.stderr)
        for leftover in target_dir.glob(f"{SUMMARY_BASENAME}*.xlsx"):
            try:
                leftover.unlink()
            except OSError:
                pass
        return

    try:
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass

    files = manifest.get("files", [])
    total_rows = manifest.get("total_rows", 0)
    print(
        f"[info] 已生成 {len(files)} 个总表文件（共 {total_rows} 行）。",
        file=sys.stderr,
    )


def _build_summary_workbooks(target_dir: Path) -> dict[str, object]:
    parsed_units = {name: parse_formula(formula) for name, formula in UNIT_FORMULAS.items()}
    masses = build_mass_table("monoisotopic", {})
    header = OUTPUT_HEADER

    row_iter = _iter_all_permutation_rows(parsed_units, masses)
    chunk_limit = SUMMARY_ROWS_PER_WORKBOOK

    chunk_files: list[str] = []
    total_rows = 0
    chunk_index = 1

    while True:
        chunk_name = _summary_chunk_name(chunk_index)
        chunk_path = target_dir / chunk_name
        rows_written, wrote = _write_summary_chunk(chunk_path, header, row_iter, chunk_limit)
        if not wrote:
            break
        chunk_files.append(chunk_name)
        total_rows += rows_written
        chunk_index += 1

    if not chunk_files:
        raise ValueError("未生成任何总表数据")

    return {
        "files": chunk_files,
        "total_rows": total_rows,
        "rows_per_workbook": chunk_limit,
        "unit_range": [SUMMARY_MIN_TOTAL_UNITS, SUMMARY_MAX_TOTAL_UNITS],
    }


def _write_summary_chunk(
    path: Path,
    header: Iterable[str],
    row_iter: Iterator[list[str]],
    chunk_limit: int,
) -> tuple[int, bool]:
    rows = islice(row_iter, chunk_limit)
    try:
        first_row = next(rows)
    except StopIteration:
        return 0, False

    chunk_rows = chain([first_row], rows)
    written = _write_table_xlsx(path, header, chunk_rows, sheet_name="glycoenum")
    return written, True


def _iter_all_permutation_rows(
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> Iterator[list[str]]:
    for total_units in range(SUMMARY_MIN_TOTAL_UNITS, SUMMARY_MAX_TOTAL_UNITS + 1):
        yield from _iter_permutation_rows_for_total(total_units, parsed_units, masses)


def _iter_permutation_rows_for_total(
    total_units: int,
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> Iterator[list[str]]:
    for counts in _iter_compositions(total_units, len(UNIT_ORDER)):
        counts_map = {UNIT_ORDER[i]: counts[i] for i in range(len(UNIT_ORDER))}
        permutations = {name: value for name, value in counts_map.items() if value}
        if not permutations:
            continue

        base_formula, final_formula, mass_text = _summarize_formula(
            total_units,
            counts_map,
            parsed_units,
            masses,
        )

        for sequence in iter_unique_permutations(permutations):
            yield [
                "-".join(sequence),
                base_formula,
                final_formula,
                mass_text,
            ]


def _summarize_formula(
    total_units: int,
    counts_map: dict[str, int],
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> tuple[str, str, str]:
    pooled = Counter()
    for name, value in counts_map.items():
        if value:
            pooled.update(scale_counts(parsed_units[name], value))

    dehydrated = dehydrate(pooled, total_units)
    final_counts = add_modifier(dehydrated, TERMINAL_MODIFIER)
    base_formula = format_hill(dehydrated)
    final_formula = format_hill(final_counts)

    neutral_mass = calculate_mass(final_counts, masses)
    theoretical_mass = apply_adduct(neutral_mass, "neutral", masses)

    return base_formula, final_formula, f"{theoretical_mass:.{SUMMARY_DECIMALS}f}"


def _iter_compositions(total: int, dimension: int) -> Iterator[tuple[int, ...]]:
    counts = [0] * dimension

    def backtrack(index: int, remaining: int) -> Iterator[tuple[int, ...]]:
        if index == dimension - 1:
            counts[index] = remaining
            yield tuple(counts)
            return
        for value in range(remaining + 1):
            counts[index] = value
            yield from backtrack(index + 1, remaining - value)

    yield from backtrack(0, total)


def _summary_chunk_name(chunk_index: int) -> str:
    if chunk_index <= 1:
        return f"{SUMMARY_BASENAME}.xlsx"
    return f"{SUMMARY_BASENAME}_part{chunk_index}.xlsx"


def _read_summary_manifest(path: Path) -> dict[str, object] | None:
    try:
        raw = path.read_text(encoding="utf-8")
    except OSError:
        return None
    try:
        data = json.loads(raw)
    except (ValueError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    return data


def _manifest_complete(manifest: dict[str, object], target_dir: Path) -> bool:
    files = manifest.get("files")
    if not isinstance(files, list) or not files:
        return False
    for entry in files:
        if not isinstance(entry, str):
            return False
        if not (target_dir / entry).exists():
            return False
    return True


CONTENT_TYPES_XML = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
"""

WORKBOOK_XML_TEMPLATE = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"{sheet_name}\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>
"""

WORKBOOK_RELS_XML = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
</Relationships>
"""

STYLES_XML = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>
  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>
  <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>
  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>
  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>
</styleSheet>
"""


def _write_xlsx(
    permutations: dict[str, int],
    base_formula: str,
    final_formula: str,
    mass_text: str,
    rows_limit: int,
    path: Path,
) -> int:
    rows_limit = max(0, rows_limit)
    if rows_limit <= 0:
        rows_iter: Iterable[Iterable[str]] = ()
    else:
        rows_iter = (
            ["-".join(sequence), base_formula, final_formula, mass_text]
            for sequence in islice(iter_unique_permutations(permutations), rows_limit)
        )
    return _write_table_xlsx(path, OUTPUT_HEADER, rows_iter, sheet_name="glycoenum")


def _write_table_xlsx(
    path: Path,
    header: Iterable[str],
    rows: Iterable[Iterable[str]],
    sheet_name: str,
) -> int:
    sheet_path: Path | None = None
    header_row = [str(value) for value in header]
    rows_written = 0

    with tempfile.NamedTemporaryFile("w", encoding="utf-8", newline="", delete=False) as sheet_tmp:
        sheet_path = Path(sheet_tmp.name)
        sheet_tmp.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        sheet_tmp.write(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
        )
        sheet_tmp.write("  <sheetData>\n")
        _write_sheet_row(sheet_tmp, 1, header_row)

        row_index = 2
        for values in rows:
            _write_sheet_row(sheet_tmp, row_index, values)
            rows_written += 1
            row_index += 1

        sheet_tmp.write("  </sheetData>\n")
        sheet_tmp.write("</worksheet>\n")

    try:
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
            workbook.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
            workbook.writestr("_rels/.rels", ROOT_RELS_XML)
            workbook.writestr("xl/workbook.xml", _build_workbook_xml(sheet_name))
            workbook.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
            workbook.writestr("xl/styles.xml", STYLES_XML)
            workbook.write(sheet_path, "xl/worksheets/sheet1.xml")
    finally:
        if sheet_path is not None:
            sheet_path.unlink(missing_ok=True)

    return rows_written


def _build_workbook_xml(sheet_name: str) -> str:
    safe_name = escape(sheet_name, {"\"": "&quot;"})
    return WORKBOOK_XML_TEMPLATE.format(sheet_name=safe_name)


def _write_sheet_row(stream: TextIO, row_index: int, values: Iterable[str]) -> None:
    stream.write(f"    <row r=\"{row_index}\">")
    for column_index, value in enumerate(values, start=1):
        cell_ref = f"{_column_letter(column_index)}{row_index}"
        text = escape(str(value))
        stream.write(
            f"<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{text}</t></is></c>"
        )
    stream.write("</row>\n")


def _column_letter(index: int) -> str:
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":  # pragma: no cover
    main()
