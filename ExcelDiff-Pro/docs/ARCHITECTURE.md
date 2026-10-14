# Architecture & security / 架构与安全

## English

ExcelDiff Pro uses a **single Python process** with Streamlit for the UI. There is no database or background queue.

```text
Browser uploads (2 XLSX / CSV files)
    |
    v
Streamlit file_uploader  [HTTPS required on public deployments]
    |
    v
File checks            [extension, ZIP-size, row/column and byte limits]
    |
    v
Excel/CSV parser       [openpyxl read_only + data_only / standard csv module]
    |
    v
Session-scoped tables  [server RAM, cleared on session end/reset]
    |
    v
Mapping + validation   [column letters, start rows, duplicate/empty checks]
    |
    v
Deterministic engine   [composite keys, normalization, Decimal arithmetic]
    |
    v
In-memory report       [openpyxl, formula injection protection]
    |
    v
Download XLSX / CSV    [no workbook persistence / no SQL]
```

### Important invariants

1. Physical column letters refer to zero-based offsets in the imported tabular data (`B` = index 1). Row numbers remain 1-based Excel positions.
2. Missing keys never match other missing keys. Duplicates are flagged rather than silently matched; aggregating them requires explicit opt-in.
3. Number parsing uses `Decimal` and an explicit decimal separator. Bad numeric text is flagged `INVALID VALUE`.
4. The comparison uses one matching key per source (or two composite columns) and one optional value pair. There is no fuzzy matching.
5. Only known matching rules are applied. No AI inference, uncontrolled network calls or customer-data analytics.
6. The saved mapping JSON contains config and worksheet names, not uploaded file data. Reinspect columns if source layouts change.
7. Spreadsheet output strings are escaped against accidental Excel formula execution.

### Boundaries and risk

In-memory is not the same as end-to-end encrypted or entirely local. Clients upload over HTTP(S) to the **hosting server**. A multiuser host must have proper TLS, isolation, authentication as appropriate, rate limiting and monitoring. Streamlit sessions eventually expire, but the app must not promise a deletion time or forensic erasure. File validation reduces attack surface but does not make parsing arbitrary untrusted XLSX perfectly secure. No independent security review has been performed.

### Tests

`tests/test_core.py` validates matching and numerical edge cases, `tests/test_io_report.py` validates actual XLSX/CSV loading and output reports, `tests/test_presets.py` checks mapping serialization validation, `tests/test_cli.py` checks batch end-to-end output.

## 中文（简体）

这是一个 **Streamlit + Python** 单体架构。两个 Excel / CSV 从浏览器上传到服务端，由 openpyxl 和 CSV Reader 读取到**会话内存**。用户选择 Column Letter 和 Row Start 后，系统检查 Key、重复值与金额格式，运行确定性的对账逻辑，再把结果用 openpyxl 生成内存中的 Excel 报告供下载。

**没有数据库、没有用户账号、没有队列，也不会主动把 Excel 转存到磁盘。** 不过，浏览器需要把文件发送到部署服务器，因此使用公网服务时必须考虑 HTTPS、上传内容安全、用户隔离、权限控制与主机日志政策。

主要验证规则：

- B/E 等字母代表 Excel 真实栏位位置；行号保留为原始工作表的行号。
- 空 Key 不会相互匹配；重复 Key 默认要求人工检查，只有用户明确选择才会加总。
- 金额用 `Decimal` 处理，并要求指定小数点格式；错误金额会被标记。
- 目前每次比对一个数值栏位对，也支持第二个 Matching Key（例如 `Order ID + Currency`）。
- JSON Mapping Profile 仅储存配置，不包含 Excel 内容。若 Excel 插入新 Column，需要重新检查。
- 导出 Excel / CSV 会防护以 `=`, `+`, `-`, `@` 开头的危险公式型文字。

已加入多种自动测试，但还需要在你真实的托管环境中做功能测试、压力测试和必要的安全审核，才建议用于敏感或高价值业务数据。
