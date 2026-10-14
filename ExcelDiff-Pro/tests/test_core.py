from dataclasses import replace
from decimal import Decimal
import pytest
from exceldiff.core import (
    CompareConfig, compare_tables, inspect_table, parse_number, normalize_key,
    MATCHED, DIFFERENT, ONLY_A, ONLY_B, DUPLICATE,
    INVALID_KEY_A, INVALID_KEY_B, INVALID_VALUE,
)
from exceldiff.io import DataTable


def t(name, rows):
    width = max(len(x) for x in rows)
    return DataTable(name, "Sheet1", tuple(tuple(x) + (None,) * (width - len(x)) for x in rows), width)


def config(**kwargs):
    base = CompareConfig(keys_a=(0,), keys_b=(0,), value_a=1, value_b=1, start_a=2, start_b=2)
    return replace(base, **kwargs)


def compare(a, b, **kwargs):
    return compare_tables(t("A.xlsx", a), t("B.xlsx", b), config(**kwargs))


def test_basic_matched_different_missing_both_sides():
    r = compare([["key", "value"],["A",100],["B",200],["C",300]], [["key", "value"],["A",100],["B",180],["D",50]])
    assert r.metrics["matched"] == 1
    assert r.metrics["different"] == 1
    assert r.metrics["only_a"] == 1
    assert r.metrics["only_b"] == 1
    assert r.metrics["total_results"] == 4
    assert next(x for x in r.rows if x["Key"] == "B")["Difference (A - B)"] == "20"


def test_duplicate_keys_flagged_not_matched():
    r = compare([["k","v"],["A",10],["A",12]], [["k","v"],["A",22]])
    assert r.rows[0]["Result"] == DUPLICATE
    assert r.metrics["matched"] == 0
    assert r.rows[0]["A · Excel rows"] == "2, 3"


def test_aggregate_duplicate_sum():
    r = compare([["k","v"],["A",10],["A",12]], [["k","v"],["A",22]], aggregate=True)
    assert r.rows[0]["Result"] == MATCHED
    assert r.rows[0]["A · Value"] == "22"


def test_invalid_keys_never_match_each_other():
    r = compare([["k","v"],[None,10],["  ",20],["A",30]], [["k","v"],[None,10],["A",30]])
    assert r.metrics["matched"] == 1
    assert r.metrics["invalid_key"] == 3
    assert {x["Result"] for x in r.exceptions} == {INVALID_KEY_A,INVALID_KEY_B}


def test_invalid_number_not_silent_zero():
    r = compare([["k","v"],["A","N/A"]], [["k","v"],["A",0]])
    assert r.rows[0]["Result"] == INVALID_VALUE


def test_amount_tolerance():
    r = compare([["k","v"],["A","10.001"]], [["k","v"],["A","10.01"]], tolerance=Decimal("0.01"))
    assert r.rows[0]["Result"] == MATCHED
    r2 = compare([["k","v"],["A","10.001"]], [["k","v"],["A","10.01"]], tolerance=Decimal("0.001"))
    assert r2.rows[0]["Result"] == DIFFERENT


def test_text_compare_case_and_trim():
    a = [["k","v"],["A"," Paid  "]]
    b = [["k","v"],["A","paid"]]
    assert compare(a,b,compare_type="text").rows[0]["Result"] == DIFFERENT
    assert compare(a,b,compare_type="text",ignore_text_case=True).rows[0]["Result"] == MATCHED


def test_presence_mode_needs_no_amount():
    r = compare([["k"],["A"],["C"]],[["k"],["A"]],compare_type="presence",value_a=None,value_b=None)
    assert r.metrics["matched"] == 1
    assert r.metrics["only_a"] == 1


def test_key_normalization_preserves_leading_zeros():
    assert normalize_key("0012",trim=True,ignore_case=True) != normalize_key(12,trim=True,ignore_case=True)
    assert normalize_key(" ABc ",trim=True,ignore_case=True) == "abc"


def test_composite_keys_are_not_concatenated_unsafely():
    a = t("A",[["x","y","amount"],["AB","C",10],["A","BC",20]])
    b = t("B",[["x","y","amount"],["A","BC",20],["AB","C",10]])
    cfg = CompareConfig(keys_a=(0,1),keys_b=(0,1),value_a=2,value_b=2,start_a=2,start_b=2)
    r = compare_tables(a,b,cfg)
    assert r.metrics["matched"] == 2


def test_numbers_supported():
    assert parse_number("RM 1,234.50") == Decimal("1234.50")
    assert parse_number("(1,234.50)") == Decimal("-1234.50")
    assert parse_number("1.234,50",",") == Decimal("1234.50")
    assert parse_number("RM75.00") == Decimal("75.00")
    assert parse_number("1,2,3") is None
    assert parse_number("N/A") is None
    assert parse_number("12.34", ",") is None or parse_number("12.34",",") == Decimal("1234")  # locale ambiguity is explicit


def test_no_data_rejected():
    with pytest.raises(ValueError):
        compare([["k","v"]],[["k","v"],["A",5]])


def test_invalid_configuration_rejected():
    with pytest.raises(ValueError):
        compare([["k","v"],["A",1]],[["k","v"],["A",1]],value_a=0)
    with pytest.raises(ValueError):
        compare([["k","v"],["A",1]],[["k","v"],["A",1]],tolerance=Decimal("-1"))


def test_extra_columns_in_result():
    r=compare([["k","v","name"],["A",5,"Ali"]],[["k","v","status"],["A",5,"Paid"]], extra_a=(2,),extra_b=(2,))
    assert r.rows[0]["A · C"]=="Ali"
    assert r.rows[0]["B · C"]=="Paid"


def test_row_numbers_original_positions():
    r=compare([["k","v"],["A",5],["",None],["B",7]], [["k","v"],["A",5],["B",7]])
    assert next(x for x in r.rows if x["Key"]=="B")["A · Excel rows"]=="4"


def test_diagnostics():
    table=t("X",[["k","v"],["a",1],["a",2],["",3],["b",4]])
    d=inspect_table(table,(0,),2,config())
    assert d["duplicate_keys"]==1
    assert d["empty_keys"]==1
    assert d["data_rows"]==4


def test_rmless_space_currency_formats():
    for text in ("RM75.00", "RM 75.00", "MYR75.00", "75.00 RM"):
        assert parse_number(text) == Decimal("75.00")
    assert parse_number("RMNOTANUMBER") is None


def test_mapping_diagnostics_preview_and_low_overlap():
    from exceldiff.core import mapping_diagnostics
    a = t("A", [["ID","Amount"],["A", 10],["B", 20],["X", "BAD"]])
    b = t("B", [["Reference","Paid"],["A", 10],["C", 20]])
    d = mapping_diagnostics(a,b,config())
    assert d["matched_unique_keys"] == 1
    assert d["unique_a"] == 3
    assert d["numeric_invalid_a"] == 1
    assert d["preview"][0]["A row"] == 2


def test_aggregate_reports_its_usage():
    result = compare([['k','v'],['A',1],['A',2]],[['k','v'],['A',3]],aggregate=True)
    assert result.metrics['matched'] == 1
    assert any('Group & SUM' in message for message in result.warnings)
