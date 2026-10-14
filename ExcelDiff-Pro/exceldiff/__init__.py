"""ExcelDiff Pro - local/in-memory spreadsheet comparison engine."""
from .core import CompareConfig, ComparisonResult, compare_tables, inspect_table, mapping_diagnostics
from .io import DataTable, DataImportError, column_options, list_sheets, load_table

__all__ = ["CompareConfig", "ComparisonResult", "compare_tables", "inspect_table", "mapping_diagnostics", "DataTable", "DataImportError", "column_options", "list_sheets", "load_table"]
