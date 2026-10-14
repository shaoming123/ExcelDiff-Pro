import json
from dataclasses import replace
from decimal import Decimal
import pytest
from exceldiff.core import CompareConfig
from exceldiff.presets import export_profile, load_profile


def test_round_trip_bilingual_filenames_and_settings():
    cfg = replace(CompareConfig(),keys_a=(1,2),keys_b=(2,0),tolerance=Decimal('0.001'),extra_a=(3,))
    raw=export_profile(cfg,'订单表','Paid')
    loaded,sheets=load_profile(raw)
    assert loaded == cfg
    assert sheets=={'a':'订单表','b':'Paid'}


def test_reject_bad_mapping_profiles():
    with pytest.raises(ValueError):
        load_profile(b'{}')
    with pytest.raises(ValueError):
        load_profile(b'a'*20000)
    raw=json.loads(export_profile(CompareConfig(),"A","B"))
    raw['settings']['keys_a']=[True]
    with pytest.raises(ValueError):
        load_profile(json.dumps(raw).encode())
    raw=json.loads(export_profile(CompareConfig(),"A","B"))
    raw['settings']['tolerance'] = 'NaN'
    with pytest.raises(ValueError):
        load_profile(json.dumps(raw).encode())


def test_profile_rejects_missing_matching_keys_and_value_conflicts():
    obj=json.loads(export_profile(CompareConfig(),"A","B"))
    obj['settings']['keys_a'] = []
    obj['settings']['keys_b'] = []
    with pytest.raises(ValueError,match="matching keys"):
        load_profile(json.dumps(obj).encode())
    obj=json.loads(export_profile(CompareConfig(),"A","B"))
    obj['settings']['value_a'] = 1
    with pytest.raises(ValueError,match="value column"):
        load_profile(json.dumps(obj).encode())
