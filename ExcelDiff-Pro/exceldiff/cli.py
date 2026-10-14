"""Optional command-line interface for scheduled / local Excel reconciliation.

Example:
    python -m exceldiff.cli orders.xlsx payments.xlsx --key-a B --key-b C --value-a E --value-b F
"""
from __future__ import annotations

import argparse
from decimal import Decimal, InvalidOperation
from pathlib import Path
import sys

from openpyxl.utils.cell import column_index_from_string

from .core import CompareConfig, compare_tables
from .io import DataImportError, list_sheets, load_table
from .presets import load_profile
from .report import export_report


def letter(value: str) -> int:
    try:
        idx = column_index_from_string(value.strip().upper()) - 1
        if idx > 99:
            raise ValueError()
        return idx
    except (ValueError, TypeError) as exc:
        raise argparse.ArgumentTypeError("Use an Excel column letter from A through CV (max 100 columns).") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="ExcelDiff Pro: compare two Excel/CSV files with column-letter mapping")
    parser.add_argument("source_a", type=Path, help="First XLSX/CSV file")
    parser.add_argument("source_b", type=Path, help="Second XLSX/CSV file")
    parser.add_argument("--sheet-a", help="Worksheet name for Excel A (CSV uses CSV Data)")
    parser.add_argument("--sheet-b", help="Worksheet name for Excel B")
    parser.add_argument("--key-a", type=letter, help="Matching key column in source A, e.g. B")
    parser.add_argument("--key-b", type=letter, help="Matching key column in source B, e.g. C")
    parser.add_argument("--value-a", type=letter, help="Comparison column in A, e.g. E")
    parser.add_argument("--value-b", type=letter, help="Comparison column in B, e.g. F")
    parser.add_argument("--second-key-a", type=letter, help="Second key column in A, optional")
    parser.add_argument("--second-key-b", type=letter, help="Second key column in B, optional")
    parser.add_argument("--start-a", type=int, default=2, help="First data row in A (default: 2)")
    parser.add_argument("--start-b", type=int, default=2, help="First data row in B (default: 2)")
    parser.add_argument("--mode", choices=["number", "text", "presence"], default="number")
    parser.add_argument("--tolerance", default="0.01", help="Maximum absolute numeric difference")
    parser.add_argument("--decimal-a", choices=[".", ","], default=".")
    parser.add_argument("--decimal-b", choices=[".", ","], default=".")
    parser.add_argument("--sum-duplicates", action="store_true", help="Group & SUM repeated keys (numeric mode only)")
    parser.add_argument("--preset", type=Path, help="Use a mapping JSON profile exported by the web app")
    parser.add_argument("--output", type=Path, default=Path("ExcelDiff_Comparison_Report.xlsx"))
    return parser


def run(argv: list[str] | None = None) -> int:
    p = build_parser()
    args = p.parse_args(argv)
    try:
        if args.preset:
            config, sheets = load_profile(args.preset.read_bytes())
            selected_a = args.sheet_a or sheets["a"]
            selected_b = args.sheet_b or sheets["b"]
        else:
            if args.key_a is None or args.key_b is None:
                p.error("--key-a and --key-b are required without --preset")
            if args.mode != "presence" and (args.value_a is None or args.value_b is None):
                p.error("--value-a and --value-b are required in number and text modes")
            try:
                tolerance = Decimal(args.tolerance)
            except InvalidOperation as exc:
                raise ValueError("Invalid tolerance. Use 0 or a positive decimal number.") from exc
            config = CompareConfig(
                keys_a=(args.key_a,) + ((args.second_key_a,) if args.second_key_a is not None else ()),
                keys_b=(args.key_b,) + ((args.second_key_b,) if args.second_key_b is not None else ()),
                value_a=args.value_a, value_b=args.value_b,
                start_a=args.start_a, start_b=args.start_b, compare_type=args.mode,
                tolerance=tolerance, decimal_a=args.decimal_a, decimal_b=args.decimal_b,
                aggregate=args.sum_duplicates,
            )
            selected_a = args.sheet_a
            selected_b = args.sheet_b
        content_a = args.source_a.read_bytes()
        content_b = args.source_b.read_bytes()
        if selected_a is None:
            selected_a = list_sheets(content_a, args.source_a.name)[0]
        if selected_b is None:
            selected_b = list_sheets(content_b, args.source_b.name)[0]
        a = load_table(content_a, args.source_a.name, selected_a)
        b = load_table(content_b, args.source_b.name, selected_b)
        result = compare_tables(a, b, config)
        file_bytes = export_report(a, b, result, config)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(file_bytes)
        print(f"ExcelDiff Pro: {result.metrics['matched']} matched, "
              f"{len(result.exceptions)} exceptions, {result.metrics['total_results']} result rows")
        print(f"Report saved: {args.output.resolve()}")
        for warning in result.warnings:
            print("Warning: " + warning, file=sys.stderr)
        return 0
    except (OSError, DataImportError, ValueError) as exc:
        print("Error: " + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(run())
