from __future__ import annotations

import json
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass
from itertools import chain, islice
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, TextIO
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from xml.sax.saxutils import escape

import license_manager
import sys

from glycoenum.formula import add_modifier, dehydrate, format_hill, parse_formula, scale_counts
from glycoenum.mass import apply_adduct, build_mass_table, calculate_mass
from glycoenum.permute import iter_unique_permutations, permutation_count

WINDOW_TITLE = "Oligosaccharide prediction"
UNIT_ORDER = ["Hex", "deoxyhex", "pent", "HexN", "UA", "HexNAc"]
UNIT_FORMULAS = {
    "Hex": "C6H12O6",
    "deoxyhex": "C6H12O5",
    "pent": "C5H10O5",
    "HexN": "C6H13NO5",
    "UA": "C6H10O7",
    "HexNAc": "C8H15NO6",
}
TERMINAL_MODIFIER = "C20H18N4O"
MIN_TOTAL_UNITS = 2
MAX_TOTAL_UNITS = 10
DEFAULT_MASS_MODEL = "monoisotopic"
DEFAULT_ADDUCT = "neutral"
DEFAULT_DECIMALS = 4
EXPORT_HEADER = [
    "Predicted compound",
    "Pre-derivatization molecular formula",
    "Post-derivatization molecular formula",
    "Calculated mass",
    "Theoretical m/z",
]
XLSX_FILENAME = "Oligosaccharide_prediction_output.xlsx"
WORKBOOK_SHEET_NAME = "glycoenum"
SUMMARY_BASENAME = "Oligosaccharide_prediction_summary"
SUMMARY_MANIFEST_NAME = f"{SUMMARY_BASENAME}_manifest.json"
SUMMARY_MIN_TOTAL_UNITS = 2
SUMMARY_MAX_TOTAL_UNITS = 10
SUMMARY_DECIMALS = 4
SUMMARY_ROWS_PER_WORKBOOK = 1_048_575


@dataclass
class CalculationResult:
    base_formula: str
    final_formula: str
    neutral_mass: float
    theoretical_mass: float
    proton_mass: float
    total_permutations: int
    permutation_counts: Dict[str, int]
    sequences: List[str]
    decimals: int

    @property
    def displayed_count(self) -> int:
        return len(self.sequences)

    @property
    def formatted_mass(self) -> str:
        return f"{self.theoretical_mass:.{self.decimals}f}"

    @property
    def theoretical_mz(self) -> float:
        return self.theoretical_mass + self.proton_mass

    @property
    def formatted_mz(self) -> str:
        return f"{self.theoretical_mz:.{self.decimals}f}"

    @property
    def truncated(self) -> bool:
        return self.displayed_count < self.total_permutations


class OligosaccharideApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("1120x720")
        self.minsize(960, 640)
        self._configure_theme()

        self.unit_vars: Dict[str, tk.IntVar] = {
            name: tk.IntVar(value=0) for name in UNIT_ORDER
        }

        self.summary_vars: Dict[str, tk.StringVar] = {
            "pre_formula": tk.StringVar(value="-"),
            "post_formula": tk.StringVar(value="-"),
            "calculated_mass": tk.StringVar(value="-"),
            "theoretical_mz": tk.StringVar(value="-"),
            "total_results": tk.StringVar(value="-"),
            "status": tk.StringVar(value="Waiting for input"),
        }

        self.current_result: CalculationResult | None = None

        self._build_layout()

    def _configure_theme(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("TkDefaultFont", 18, "bold"))
        style.configure("Header.TLabel", font=("TkDefaultFont", 11, "bold"))

    def _build_layout(self) -> None:
        container = ttk.Frame(self, padding=(20, 16, 20, 16))
        container.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(5, weight=1)

        title = ttk.Label(
            container,
            text="Oligosaccharide prediction",
            style="Title.TLabel",
        )
        title.grid(row=0, column=0, sticky="w")

        description = ttk.Label(
            container,
            text=(
                "Enter the number of units for each monosaccharide, then choose "
                "\"Calculate sequences\" to enumerate every unique arrangement, apply "
                "dehydration and terminal derivatization, and report the resulting masses. "
                "Use \"Export XLSX...\" to save the current table or \"Generate summary...\" "
                "to build the full 2-10 unit reference workbooks."
            ),
            wraplength=720,
        )
        description.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 12))

        defaults_label = ttk.Label(
            container,
            text=(
                "Using defaults: monoisotopic mass model, neutral adduct, four decimal places."
                " All permutations are enumerated; large totals may take time to display."
            ),
            foreground="#555555",
            wraplength=720,
        )
        defaults_label.grid(row=2, column=0, sticky="w", pady=(0, 12))

        top_section = ttk.Frame(container)
        top_section.grid(row=3, column=0, sticky="ew")
        top_section.columnconfigure(0, weight=1)

        self._build_composition_frame(top_section)

        self._build_action_frame(container)
        self._build_results_frame(container)

    def _build_composition_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Composition", padding=12)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        frame.columnconfigure(1, weight=1)

        for row, name in enumerate(UNIT_ORDER):
            ttk.Label(frame, text=f"{name} units:", style="Header.TLabel").grid(
                row=row,
                column=0,
                sticky="w",
                padx=(0, 8),
                pady=4,
            )
            spinbox = ttk.Spinbox(
                frame,
                from_=0,
                to=20,
                textvariable=self.unit_vars[name],
                width=8,
            )
            spinbox.grid(row=row, column=1, sticky="w", pady=4)

    def _build_action_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent)
        frame.grid(row=4, column=0, sticky="ew", pady=(18, 12))
        frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(frame, textvariable=self.summary_vars["status"])
        self.status_label.grid(row=0, column=0, sticky="w")

        action_group = ttk.Frame(frame)
        action_group.grid(row=0, column=1, sticky="e")

        calculate_button = ttk.Button(
            action_group,
            text="Calculate sequences",
            command=self._handle_calculate,
        )
        calculate_button.grid(row=0, column=0, padx=(0, 8))

        export_button = ttk.Button(
            action_group,
            text="Export XLSX...",
            command=self._handle_export,
        )
        export_button.grid(row=0, column=1)
        self.export_button = export_button
        self.export_button.state(["disabled"])

        summary_button = ttk.Button(
            action_group,
            text="Generate summary...",
            command=self._handle_generate_summary,
        )
        summary_button.grid(row=0, column=2, padx=(8, 0))

    def _build_results_frame(self, parent: ttk.Frame) -> None:
        frame = ttk.LabelFrame(parent, text="Results", padding=12)
        frame.grid(row=5, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        summary = ttk.Frame(frame)
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        summary.columnconfigure(1, weight=1)

        summary_fields = [
            ("Pre-derivatization molecular formula", "pre_formula"),
            ("Post-derivatization molecular formula", "post_formula"),
            ("Calculated mass", "calculated_mass"),
            ("Theoretical m/z", "theoretical_mz"),
            ("Total results", "total_results"),
        ]

        for row, (label, key) in enumerate(summary_fields):
            ttk.Label(summary, text=label, style="Header.TLabel").grid(
                row=row, column=0, sticky="w", pady=(0 if row == 0 else 6, 0)
            )
            ttk.Label(
                summary,
                textvariable=self.summary_vars[key],
            ).grid(row=row, column=1, sticky="w", padx=(12, 0), pady=(0 if row == 0 else 6, 0))

        table_frame = ttk.Frame(frame)
        table_frame.grid(row=1, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = ("index", "sequence")
        tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=12,
        )
        tree.heading("index", text="#")
        tree.column("index", width=60, anchor="center")
        tree.heading("sequence", text="Sequence")
        tree.column("sequence", width=240, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        self.tree = tree

        scrollbar = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

    def _handle_calculate(self) -> None:
        try:
            counts = {name: int(var.get()) for name, var in self.unit_vars.items()}
        except tk.TclError:
            messagebox.showerror(
                WINDOW_TITLE,
                "Unit counts must be valid integers.",
                parent=self,
            )
            return

        if any(value < 0 for value in counts.values()):
            messagebox.showerror(
                WINDOW_TITLE,
                "Unit counts cannot be negative.",
                parent=self,
            )
            return

        total_units = sum(counts.values())
        if total_units < MIN_TOTAL_UNITS or total_units > MAX_TOTAL_UNITS:
            messagebox.showwarning(
                WINDOW_TITLE,
                (
                    f"Total unit count must be between {MIN_TOTAL_UNITS} and "
                    f"{MAX_TOTAL_UNITS}. Current total: {total_units}."
                ),
                parent=self,
            )
            return

        try:
            result = self._calculate(counts)
        except (ValueError, KeyError) as exc:
            messagebox.showerror(WINDOW_TITLE, str(exc), parent=self)
            return

        self._present_result(result)

    def _calculate(self, counts: Dict[str, int]) -> CalculationResult:
        compositions = {
            name: scale_counts(parse_formula(formula), counts[name])
            for name, formula in UNIT_FORMULAS.items()
        }

        pooled = Counter()
        for composition in compositions.values():
            pooled.update(composition)

        total_units = sum(counts.values())
        dehydrated = dehydrate(pooled, total_units)
        final_counts = add_modifier(dehydrated, TERMINAL_MODIFIER)

        base_formula = format_hill(dehydrated)
        final_formula = format_hill(final_counts)

        masses = build_mass_table(DEFAULT_MASS_MODEL, {})
        neutral_mass = calculate_mass(final_counts, masses)
        theoretical_mass = apply_adduct(neutral_mass, DEFAULT_ADDUCT, masses)
        proton_mass = masses.get("H")
        if proton_mass is None:
            raise KeyError("Mass table missing 'H' required for theoretical m/z calculation")

        permutations = {name: count for name, count in counts.items() if count > 0}
        total_permutations = permutation_count(permutations)
        sequences: List[str] = []
        if total_permutations > 0:
            generator = iter_unique_permutations(permutations)
            sequences = ["-".join(sequence) for sequence in generator]

        return CalculationResult(
            base_formula=base_formula,
            final_formula=final_formula,
            neutral_mass=neutral_mass,
            theoretical_mass=theoretical_mass,
            proton_mass=float(proton_mass),
            total_permutations=total_permutations,
            permutation_counts=permutations,
            sequences=sequences,
            decimals=DEFAULT_DECIMALS,
        )

    def _present_result(self, result: CalculationResult) -> None:
        self.current_result = result
        self.tree.delete(*self.tree.get_children())
        for index, sequence in enumerate(result.sequences, start=1):
            self.tree.insert("", "end", values=(index, sequence))

        self.summary_vars["pre_formula"].set(result.base_formula)
        self.summary_vars["post_formula"].set(result.final_formula)
        self.summary_vars["calculated_mass"].set(result.formatted_mass)
        self.summary_vars["theoretical_mz"].set(result.formatted_mz)
        self.summary_vars["total_results"].set(str(result.total_permutations))

        if result.total_permutations == 0:
            status_text = "No sequences available"
        elif result.truncated:
            status_text = (
                f"Showing {result.displayed_count} of {result.total_permutations} sequences"
            )
        else:
            status_text = f"Showing all {result.total_permutations} sequences"

        self.summary_vars["status"].set(status_text)
        self.export_button.state(["!disabled"] if result.displayed_count else ["disabled"])

    def _handle_export(self) -> None:
        if not self.current_result or not self.current_result.sequences:
            messagebox.showinfo(
                WINDOW_TITLE,
                "Calculate sequences before exporting.",
                parent=self,
            )
            return

        filename = filedialog.asksaveasfilename(
            parent=self,
            title="Export sequences",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
            initialfile=XLSX_FILENAME,
        )
        if not filename:
            return

        path = Path(filename)
        try:
            rows_written = _write_result_xlsx(self.current_result, path)
        except (OSError, ValueError, KeyError) as exc:
            messagebox.showerror(
                WINDOW_TITLE,
                f"Failed to export workbook: {exc}",
                parent=self,
            )
            return

        messagebox.showinfo(
            WINDOW_TITLE,
            f"Exported {rows_written} sequences to {path.name}.",
            parent=self,
        )

    def _handle_generate_summary(self) -> None:
        confirm = messagebox.askyesno(
            WINDOW_TITLE,
            (
                "Generate summary workbooks for all unit compositions (2-10 total units)?\n"
                "This may take a long time and produce large files."
            ),
            parent=self,
        )
        if not confirm:
            return

        directory = filedialog.askdirectory(
            parent=self,
            title="Select output folder",
        )
        if not directory:
            return

        target_dir = Path(directory)
        self.summary_vars["status"].set("Generating summary workbooks...")
        self.update_idletasks()

        try:
            manifest = generate_summary_workbooks(target_dir)
        except (OSError, ValueError) as exc:
            messagebox.showerror(
                WINDOW_TITLE,
                f"Failed to generate summary: {exc}",
                parent=self,
            )
            self.summary_vars["status"].set("Summary generation failed")
            return

        files = manifest.get("files", [])
        total_rows = manifest.get("total_rows", 0)
        self.summary_vars["status"].set(f"Summary generated ({len(files)} file(s))")
        messagebox.showinfo(
            WINDOW_TITLE,
            (
                f"Created {len(files)} summary workbook(s) "
                f"with {total_rows} total rows in {target_dir}."
            ),
            parent=self,
        )


def _iter_result_rows(result: CalculationResult) -> Iterable[List[str]]:
    base_formula = result.base_formula
    final_formula = result.final_formula
    mass_text = result.formatted_mass
    mz_text = result.formatted_mz

    if result.sequences:
        for sequence in result.sequences:
            text = sequence if isinstance(sequence, str) else "-".join(sequence)
            yield [text, base_formula, final_formula, mass_text, mz_text]
    else:
        for sequence in iter_unique_permutations(result.permutation_counts):
            yield [
                "-".join(sequence),
                base_formula,
                final_formula,
                mass_text,
                mz_text,
            ]


def _write_result_xlsx(result: CalculationResult, path: Path) -> int:
    rows = _iter_result_rows(result)
    return _write_table_xlsx(path, EXPORT_HEADER, rows, sheet_name=WORKBOOK_SHEET_NAME)


def generate_summary_workbooks(target_dir: Path) -> dict[str, object]:
    target_dir = target_dir.resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    manifest_path = target_dir / SUMMARY_MANIFEST_NAME
    for leftover in target_dir.glob(f"{SUMMARY_BASENAME}*.xlsx"):
        leftover.unlink(missing_ok=True)
    manifest_path.unlink(missing_ok=True)

    parsed_units = {name: parse_formula(formula) for name, formula in UNIT_FORMULAS.items()}
    masses = build_mass_table(DEFAULT_MASS_MODEL, {})
    header = EXPORT_HEADER

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
        raise ValueError("No summary data was generated.")

    manifest = {
        "files": chunk_files,
        "total_rows": total_rows,
        "rows_per_workbook": chunk_limit,
        "unit_range": [SUMMARY_MIN_TOTAL_UNITS, SUMMARY_MAX_TOTAL_UNITS],
        "mass_model": DEFAULT_MASS_MODEL,
        "adduct": DEFAULT_ADDUCT,
        "decimals": SUMMARY_DECIMALS,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def _write_summary_chunk(
    path: Path,
    header: Iterable[str],
    row_iter: Iterator[List[str]],
    chunk_limit: int,
) -> tuple[int, bool]:
    rows = islice(row_iter, chunk_limit)
    try:
        first_row = next(rows)
    except StopIteration:
        return 0, False

    chunk_rows = chain([first_row], rows)
    written = _write_table_xlsx(path, header, chunk_rows, sheet_name=WORKBOOK_SHEET_NAME)
    return written, True


def _iter_all_permutation_rows(
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> Iterator[List[str]]:
    for total_units in range(SUMMARY_MIN_TOTAL_UNITS, SUMMARY_MAX_TOTAL_UNITS + 1):
        yield from _iter_permutation_rows_for_total(total_units, parsed_units, masses)


def _iter_permutation_rows_for_total(
    total_units: int,
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> Iterator[List[str]]:
    for counts in _iter_compositions(total_units, len(UNIT_ORDER)):
        counts_map = {UNIT_ORDER[i]: counts[i] for i in range(len(UNIT_ORDER))}
        permutations = {name: value for name, value in counts_map.items() if value}
        if not permutations:
            continue

        base_formula, final_formula, mass_text, mz_text = _summarize_formula(
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
                mz_text,
            ]


def _summarize_formula(
    total_units: int,
    counts_map: dict[str, int],
    parsed_units: dict[str, dict[str, int]],
    masses: dict[str, float],
) -> tuple[str, str, str, str]:
    pooled = Counter()
    for name, value in counts_map.items():
        if value:
            pooled.update(scale_counts(parsed_units[name], value))

    dehydrated = dehydrate(pooled, total_units)
    final_counts = add_modifier(dehydrated, TERMINAL_MODIFIER)
    base_formula = format_hill(dehydrated)
    final_formula = format_hill(final_counts)

    neutral_mass = calculate_mass(final_counts, masses)
    theoretical_mass = apply_adduct(neutral_mass, DEFAULT_ADDUCT, masses)
    proton_mass = masses.get("H")
    if proton_mass is None:
        raise KeyError("Mass table missing 'H' required for theoretical m/z calculation")
    mass_text = f"{theoretical_mass:.{SUMMARY_DECIMALS}f}"
    mz_text = f"{theoretical_mass + float(proton_mass):.{SUMMARY_DECIMALS}f}"

    return base_formula, final_formula, mass_text, mz_text


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


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

WORKBOOK_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


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
        sheet_tmp.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        sheet_tmp.write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
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
    safe_name = escape(sheet_name, {'"': "&quot;"})
    return WORKBOOK_XML_TEMPLATE.format(sheet_name=safe_name)


def _write_sheet_row(stream: TextIO, row_index: int, values: Iterable[str]) -> None:
    stream.write(f'    <row r="{row_index}">')
    for column_index, value in enumerate(values, start=1):
        cell_ref = f"{_column_letter(column_index)}{row_index}"
        text = escape(str(value))
        stream.write(f'<c r="{cell_ref}" t="inlineStr"><is><t>{text}</t></is></c>')
    stream.write("</row>\n")


def _column_letter(index: int) -> str:
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def main() -> None:
    ok, message = license_manager.activate_if_needed()
    if not ok:
        try:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror(WINDOW_TITLE, f"License error: {message}")
            temp_root.destroy()
        except Exception:
            print(f"License error: {message}")
        sys.exit(1)

    app = OligosaccharideApp()
    app.mainloop()


if __name__ == "__main__":  # pragma: no cover
    main()
