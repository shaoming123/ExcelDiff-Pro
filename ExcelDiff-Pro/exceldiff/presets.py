"""Export/import a portable, database-free comparison mapping profile.

A profile stores column positions and rules, never uploaded spreadsheet data.
"""
from __future__ import annotations

from dataclasses import asdict
from decimal import Decimal, InvalidOperation
import json
from typing import Any

from .core import CompareConfig

PROFILE_VERSION = 1
MAX_PROFILE_BYTES = 16 * 1024


def export_profile(config: CompareConfig, sheet_a: str, sheet_b: str) -> bytes:
    profile = {
        "app": "ExcelDiff Pro",
        "profile_version": PROFILE_VERSION,
        "note": "Recheck the column letters and sample values if file layouts change.",
        "sheets": {"a": sheet_a, "b": sheet_b},
        "settings": {**asdict(config), "tolerance": str(config.tolerance)},
    }
    return (json.dumps(profile, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def load_profile(payload: bytes) -> tuple[CompareConfig, dict[str, str]]:
    if len(payload) > MAX_PROFILE_BYTES:
        raise ValueError("Mapping profile is too large (16 KB maximum).")
    try:
        value = json.loads(payload.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Mapping profile must be valid UTF-8 JSON.") from exc
    if not isinstance(value, dict) or value.get("app") != "ExcelDiff Pro" or value.get("profile_version") != PROFILE_VERSION:
        raise ValueError("Unrecognized mapping profile or unsupported version.")
    raw = value.get("settings")
    sheets = value.get("sheets")
    if not isinstance(raw, dict) or not isinstance(sheets, dict):
        raise ValueError("Mapping profile is missing settings or sheets.")
    allowed = set(CompareConfig.__dataclass_fields__)
    if set(raw) != allowed:
        raise ValueError("Mapping profile settings are missing or contain unknown fields.")
    try:
        for attr in ("keys_a", "keys_b", "extra_a", "extra_b"):
            seq = raw[attr]
            if not isinstance(seq, list) or len(seq) > 25 or not all(type(v) is int and 0 <= v < 100 for v in seq):
                raise ValueError(f"Invalid column list: {attr}.")
            raw[attr] = tuple(seq)
        for attr in ("value_a", "value_b"):
            if raw[attr] is not None and (type(raw[attr]) is not int or not 0 <= raw[attr] < 100):
                raise ValueError(f"Invalid column number: {attr}.")
        for attr in ("start_a", "start_b"):
            if type(raw[attr]) is not int or not 1 <= raw[attr] <= 30_000:
                raise ValueError(f"Invalid starting row: {attr}.")
        for attr in ("trim_keys", "ignore_key_case", "ignore_text_case", "trim_text", "aggregate"):
            if type(raw[attr]) is not bool:
                raise ValueError(f"Invalid boolean: {attr}.")
        if raw["compare_type"] not in ("number", "text", "presence"):
            raise ValueError("Invalid comparison mode.")
        if raw["decimal_a"] not in (".", ",") or raw["decimal_b"] not in (".", ","):
            raise ValueError("Invalid decimal separator.")
        # A web profile must fit the exact controls the UI can represent.
        if not (1 <= len(raw["keys_a"]) <= 2) or len(raw["keys_a"]) != len(raw["keys_b"]):
            raise ValueError("Mapping profile must have one or two matching keys on both sides.")
        for side in ("a", "b"):
            keys = raw[f"keys_{side}"]
            if len(set(keys)) != len(keys):
                raise ValueError("Mapping profile contains repeated matching column numbers.")
            val = raw[f"value_{side}"]
            if raw["compare_type"] != "presence" and (val is None or val in keys):
                raise ValueError("A value column cannot be missing or identical to a matching key.")
        if raw["aggregate"] and raw["compare_type"] != "number":
            raise ValueError("Group & SUM is only valid for numeric comparison.")
        raw["tolerance"] = Decimal(str(raw["tolerance"]))
        if not raw["tolerance"].is_finite() or raw["tolerance"] < 0:
            raise ValueError("Invalid tolerance.")
        for k in ("a", "b"):
            if not isinstance(sheets.get(k), str) or len(sheets[k]) > 150:
                raise ValueError("Invalid sheet name.")
        cfg = CompareConfig(**raw)
    except (TypeError, InvalidOperation, OverflowError) as exc:
        raise ValueError("Invalid mapping profile configuration.") from exc
    return cfg, {"a": sheets["a"], "b": sheets["b"]}
