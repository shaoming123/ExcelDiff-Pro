"""Deterministic, side-effect-free Excel reconciliation engine.

No SQL, file system writes, background jobs or network calls.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
import re
from typing import Any
import unicodedata

from .io import DataTable, is_blank, present

MATCHED = "MATCHED"
DIFFERENT = "DIFFERENT"
ONLY_A = "ONLY IN A"
ONLY_B = "ONLY IN B"
DUPLICATE = "DUPLICATE KEY"
INVALID_KEY_A = "EMPTY KEY A"
INVALID_KEY_B = "EMPTY KEY B"
INVALID_VALUE = "INVALID VALUE"
ALL_STATUSES = (MATCHED, DIFFERENT, ONLY_A, ONLY_B, DUPLICATE, INVALID_KEY_A, INVALID_KEY_B, INVALID_VALUE)


@dataclass(frozen=True)
class CompareConfig:
    keys_a: tuple[int, ...] = (1,)
    keys_b: tuple[int, ...] = (2,)
    value_a: int | None = 4
    value_b: int | None = 5
    start_a: int = 2
    start_b: int = 2
    compare_type: str = "number"  # number / text / presence
    tolerance: Decimal = Decimal("0.01")
    decimal_a: str = "."
    decimal_b: str = "."
    trim_keys: bool = True
    ignore_key_case: bool = True
    ignore_text_case: bool = False
    trim_text: bool = True
    aggregate: bool = False  # Group payments by key then SUM: number comparison only.
    extra_a: tuple[int, ...] = ()
    extra_b: tuple[int, ...] = ()


@dataclass
class ComparisonResult:
    rows: list[dict[str, Any]]
    metrics: dict[str, int]
    warnings: list[str] = field(default_factory=list)
    config: CompareConfig | None = None

    @property
    def exceptions(self) -> list[dict[str, Any]]:
        return [row for row in self.rows if row["Result"] != MATCHED]


@dataclass(frozen=True)
class _Line:
    excel_row: int
    key: tuple[str, ...]
    key_label: str
    value: Any
    raw: tuple[Any, ...]


def validate_config(a: DataTable, b: DataTable, cfg: CompareConfig) -> None:
    if not cfg.keys_a or len(cfg.keys_a) != len(cfg.keys_b) or len(cfg.keys_a) > 3:
        raise ValueError("Select 1 to 3 matching columns on each side, with equal counts.")
    if len(set(cfg.keys_a)) != len(cfg.keys_a) or len(set(cfg.keys_b)) != len(cfg.keys_b):
        raise ValueError("Each matching column can only be selected once per file.")
    for table, columns, start in ((a, cfg.keys_a + cfg.extra_a, cfg.start_a), (b, cfg.keys_b + cfg.extra_b, cfg.start_b)):
        if not (1 <= start <= table.height):
            raise ValueError(f"Starting row must be between 1 and {table.height} for {table.filename}.")
        if any(i < 0 or i >= table.width for i in columns):
            raise ValueError(f"One or more selected columns are outside {table.filename}.")
    if cfg.compare_type not in ("number", "text", "presence"):
        raise ValueError("Choose number, text or presence comparison.")
    if cfg.compare_type != "presence":
        for table, value_col, keys in ((a, cfg.value_a, cfg.keys_a), (b, cfg.value_b, cfg.keys_b)):
            if value_col is None or not (0 <= value_col < table.width):
                raise ValueError(f"Select a valid comparison column for {table.filename}.")
            if value_col in keys:
                raise ValueError("The comparison value must not be one of the matching keys.")
    if cfg.aggregate and cfg.compare_type != "number":
        raise ValueError("Group & SUM is available for numeric comparisons only.")
    if cfg.tolerance.is_nan() or cfg.tolerance < 0 or cfg.tolerance.is_infinite():
        raise ValueError("Tolerance must be a finite, non-negative number.")
    if cfg.decimal_a not in (".", ",") or cfg.decimal_b not in (".", ","):
        raise ValueError("Decimal separator must be . or ,")


def normalize_key(value: Any, *, trim: bool, ignore_case: bool) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)
    text = unicodedata.normalize("NFKC", text)
    if trim:
        text = text.strip()
    if ignore_case:
        text = text.casefold()
    return text


def parse_number(value: Any, decimal_sep: str = ".") -> Decimal | None:
    """Parse currency-like values, rejecting ambiguous formats rather than guessing."""
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            val = Decimal(str(value))
            return val if val.is_finite() else None
        except InvalidOperation:
            return None
    if not isinstance(value, str):
        return None
    text = value.strip().replace("\u00a0", " ")
    neg = text.startswith("(") and text.endswith(")")
    if neg:
        text = text[1:-1].strip()
    # Explicitly allow common currency labels around the number.
    # Support both "RM 75.00" and "RM75.00" without accepting arbitrary words.
    currencies = r"(?:MYR|RM|USD|EUR|SGD|GBP|AUD)"
    text = re.sub(rf"(?i)^{currencies}\s*", "", text)
    text = re.sub(rf"(?i)\s*{currencies}$", "", text)
    for sym in ("$", "€", "£", "¥", "₹"):
        text = text.replace(sym, "")
    text = text.strip()
    if " " in text:
        # Spaces as thousands separators are accepted only in correctly grouped numbers.
        text = re.sub(r"(?<=\d) (?=\d{3}(?:\D|$))", "", text)
    grouping = "," if decimal_sep == "." else "."
    escaped_group = re.escape(grouping)
    escaped_decimal = re.escape(decimal_sep)
    numeric_pattern = rf"[+-]?(?:\d{{1,3}}(?:{escaped_group}\d{{3}})+|\d+)(?:{escaped_decimal}\d+)?"
    if not re.fullmatch(numeric_pattern, text):
        return None
    normalized = text.replace(grouping, "").replace(decimal_sep, ".")
    try:
        parsed = Decimal(normalized)
        return -parsed if neg else parsed
    except InvalidOperation:
        return None


def compare_text(a: Any, b: Any, *, trim: bool, ignore_case: bool) -> bool:
    if is_blank(a) or is_blank(b):
        return is_blank(a) and is_blank(b)
    left = unicodedata.normalize("NFKC", str(a))
    right = unicodedata.normalize("NFKC", str(b))
    if trim:
        left, right = left.strip(), right.strip()
    if ignore_case:
        left, right = left.casefold(), right.casefold()
    return left == right


def build_lines(table: DataTable, key_cols: tuple[int, ...], value_col: int | None, start: int, cfg: CompareConfig) -> tuple[dict[tuple[str, ...], list[_Line]], list[_Line]]:
    grouped: dict[tuple[str, ...], list[_Line]] = defaultdict(list)
    invalid: list[_Line] = []
    for excel_row, raw in table.data_rows(start):
        raw_keys = [raw[col] for col in key_cols]
        normalized = tuple(normalize_key(v, trim=cfg.trim_keys, ignore_case=cfg.ignore_key_case) for v in raw_keys)
        label = "  |  ".join(present(v, 50) or "(empty)" for v in raw_keys)
        value = raw[value_col] if value_col is not None else None
        line = _Line(excel_row=excel_row, key=normalized, key_label=label, value=value, raw=raw)
        if any(not k for k in normalized):
            invalid.append(line)
        else:
            grouped[normalized].append(line)
    return grouped, invalid


def inspect_table(table: DataTable, key_cols: tuple[int, ...], start: int, cfg: CompareConfig) -> dict[str, int]:
    group, invalid = build_lines(table, key_cols, None, start, cfg)
    return {
        "data_rows": len(table.data_rows(start)),
        "valid_keys": len(group),
        "empty_keys": len(invalid),
        "duplicate_keys": sum(1 for v in group.values() if len(v) > 1),
        "duplicate_rows": sum(len(v) for v in group.values() if len(v) > 1),
    }


def mapping_diagnostics(a: DataTable, b: DataTable, cfg: CompareConfig) -> dict[str, Any]:
    """Pre-flight checks for suspicious mappings, without pretending to know business semantics."""
    validate_config(a, b, cfg)
    groups_a, invalid_a = build_lines(a, cfg.keys_a, cfg.value_a, cfg.start_a, cfg)
    groups_b, invalid_b = build_lines(b, cfg.keys_b, cfg.value_b, cfg.start_b, cfg)
    intersect = set(groups_a) & set(groups_b)
    unique_a = len(groups_a)
    unique_b = len(groups_b)
    details: dict[str, Any] = {
        "matched_unique_keys": len(intersect),
        "unique_a": unique_a,
        "unique_b": unique_b,
        "overlap_a_percent": round(len(intersect) / max(1, unique_a) * 100, 1),
        "overlap_b_percent": round(len(intersect) / max(1, unique_b) * 100, 1),
        "empty_key_rows_a": len(invalid_a),
        "empty_key_rows_b": len(invalid_b),
        "preview": [],
        "numeric_invalid_a": 0,
        "numeric_invalid_b": 0,
    }
    for key in list(intersect)[:8]:
        # Display only first one of each key for preview; duplicates are separately warned.
        left, right = groups_a[key][0], groups_b[key][0]
        details["preview"].append({
            "Key": left.key_label,
            "A row": left.excel_row,
            "A value": present(left.value, 55),
            "B row": right.excel_row,
            "B value": present(right.value, 55),
        })
    if cfg.compare_type == "number":
        for side, groups, separator in (
            ("a", groups_a, cfg.decimal_a),
            ("b", groups_b, cfg.decimal_b),
        ):
            details[f"numeric_invalid_{side}"] = sum(
                parse_number(line.value, separator) is None
                for lines in groups.values() for line in lines
            )
    return details


def _str_value(lines: list[_Line]) -> str:
    if not lines:
        return ""
    if len(lines) == 1:
        return present(lines[0].value, 200)
    return "; ".join(present(line.value, 50) for line in lines[:10]) + (" …" if len(lines) > 10 else "")


def _num_sum(lines: list[_Line], decimal_sep: str) -> Decimal | None:
    values = [parse_number(line.value, decimal_sep) for line in lines]
    if not values or any(value is None for value in values):
        return None
    return sum(values, Decimal("0"))  # type: ignore[arg-type]


def _extra_fields(lines: list[_Line], cols: tuple[int, ...], side: str) -> dict[str, str]:
    from openpyxl.utils import get_column_letter
    result = {}
    for col in cols:
        seen = []
        for line in lines:
            s = present(line.raw[col], 140)
            if s and s not in seen:
                seen.append(s)
            if len(seen) >= 3:
                break
        result[f"{side} · {get_column_letter(col + 1)}"] = " | ".join(seen)
    return result


def _make_row(key: str, a_lines: list[_Line], b_lines: list[_Line], status: str, cfg: CompareConfig, *, difference: str = "", value_a: str | None = None, value_b: str | None = None) -> dict[str, Any]:
    row = {
        "Key": key,
        "Result": status,
        "A · Excel rows": ", ".join(str(v.excel_row) for v in a_lines[:15]) + ("…" if len(a_lines) > 15 else ""),
        "B · Excel rows": ", ".join(str(v.excel_row) for v in b_lines[:15]) + ("…" if len(b_lines) > 15 else ""),
        "A · Value": _str_value(a_lines) if value_a is None else value_a,
        "B · Value": _str_value(b_lines) if value_b is None else value_b,
        "Difference (A - B)": difference,
    }
    row.update(_extra_fields(a_lines, cfg.extra_a, "A"))
    row.update(_extra_fields(b_lines, cfg.extra_b, "B"))
    return row


def compare_tables(a: DataTable, b: DataTable, cfg: CompareConfig) -> ComparisonResult:
    validate_config(a, b, cfg)
    a_groups, a_invalid = build_lines(a, cfg.keys_a, cfg.value_a if cfg.compare_type != "presence" else None, cfg.start_a, cfg)
    b_groups, b_invalid = build_lines(b, cfg.keys_b, cfg.value_b if cfg.compare_type != "presence" else None, cfg.start_b, cfg)
    if not a_groups and not a_invalid:
        raise ValueError("File A has no data from the selected starting row.")
    if not b_groups and not b_invalid:
        raise ValueError("File B has no data from the selected starting row.")

    rows: list[dict[str, Any]] = []
    keys = list(dict.fromkeys([*a_groups.keys(), *b_groups.keys()]))
    for key in keys:
        aa, bb = a_groups.get(key, []), b_groups.get(key, [])
        label = (aa[0] if aa else bb[0]).key_label
        if (len(aa) > 1 or len(bb) > 1) and not cfg.aggregate:
            rows.append(_make_row(label, aa, bb, DUPLICATE, cfg))
            continue
        if not aa:
            rows.append(_make_row(label, aa, bb, ONLY_B, cfg))
            continue
        if not bb:
            rows.append(_make_row(label, aa, bb, ONLY_A, cfg))
            continue
        if cfg.compare_type == "presence":
            rows.append(_make_row(label, aa, bb, MATCHED, cfg))
        elif cfg.compare_type == "text":
            equal = compare_text(aa[0].value, bb[0].value, trim=cfg.trim_text, ignore_case=cfg.ignore_text_case)
            rows.append(_make_row(label, aa, bb, MATCHED if equal else DIFFERENT, cfg))
        else:
            num_a = _num_sum(aa, cfg.decimal_a)
            num_b = _num_sum(bb, cfg.decimal_b)
            if num_a is None or num_b is None:
                rows.append(_make_row(label, aa, bb, INVALID_VALUE, cfg, value_a="Invalid" if num_a is None else str(num_a), value_b="Invalid" if num_b is None else str(num_b)))
            else:
                delta = num_a - num_b
                status = MATCHED if abs(delta) <= cfg.tolerance else DIFFERENT
                rows.append(_make_row(label, aa, bb, status, cfg, difference=str(delta), value_a=str(num_a), value_b=str(num_b)))

    for line in a_invalid:
        rows.append(_make_row(line.key_label, [line], [], INVALID_KEY_A, cfg))
    for line in b_invalid:
        rows.append(_make_row(line.key_label, [], [line], INVALID_KEY_B, cfg))

    counts = Counter(row["Result"] for row in rows)
    metrics = {
        "file_a_rows": len(a.data_rows(cfg.start_a)),
        "file_b_rows": len(b.data_rows(cfg.start_b)),
        "keys_checked": len(keys),
        "matched": counts[MATCHED],
        "different": counts[DIFFERENT],
        "only_a": counts[ONLY_A],
        "only_b": counts[ONLY_B],
        "duplicate": counts[DUPLICATE],
        "invalid_key": counts[INVALID_KEY_A] + counts[INVALID_KEY_B],
        "invalid_value": counts[INVALID_VALUE],
        "total_results": len(rows),
    }
    warnings = []
    if cfg.aggregate:
        grouped_key_count = sum(len(v) > 1 for v in a_groups.values()) + sum(len(v) > 1 for v in b_groups.values())
        if grouped_key_count:
            warnings.append(f"Group & SUM was enabled. Repeated rows were summed for {grouped_key_count} source key group(s). Verify this is the intended business rule.")
    if metrics["duplicate"]:
        warnings.append(f"{metrics['duplicate']} key(s) contain duplicate rows. " + ("Values were aggregated by key." if cfg.aggregate else "They were flagged, not silently matched."))
    if metrics["invalid_key"]:
        warnings.append(f"{metrics['invalid_key']} row(s) have an empty matching key and were not matched.")
    if metrics["invalid_value"]:
        warnings.append(f"{metrics['invalid_value']} pair(s) contain unparseable numeric values. Verify decimal separators and source cells.")
    return ComparisonResult(rows=rows, metrics=metrics, warnings=warnings, config=cfg)
