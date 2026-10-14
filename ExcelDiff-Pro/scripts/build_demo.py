"""Developer-only fixture maker. Run from repository root."""
from pathlib import Path
from exceldiff.sample import demo_bytes
from exceldiff.io import load_table
from exceldiff.core import CompareConfig, compare_tables
from exceldiff.report import export_report
from exceldiff.presets import export_profile

out=Path('examples');out.mkdir(exist_ok=True)
a,b=demo_bytes()
(out/'sample_orders.xlsx').write_bytes(a)
(out/'sample_payments.xlsx').write_bytes(b)
cfg=CompareConfig(extra_a=(2,),extra_b=(4,))
res=compare_tables(load_table(a,'sample_orders.xlsx','Orders'),load_table(b,'sample_payments.xlsx','Payments'),cfg)
(out/'sample_comparison_report.xlsx').write_bytes(export_report(load_table(a,'sample_orders.xlsx','Orders'),load_table(b,'sample_payments.xlsx','Payments'),res,cfg))
(out/'sample_mapping_profile.json').write_bytes(export_profile(cfg,'Orders','Payments'))
print('Saved sample XLSX reports, demo input files and profile')
