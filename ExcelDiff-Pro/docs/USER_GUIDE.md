# ExcelDiff Pro — User Guide / 使用说明

## English

**Purpose:** Identify missing, different and duplicate records between two spreadsheets even when their headers are different.

**Default language:** English. Select **简体中文** from the **Language / 语言** dropdown in the sidebar.

**Need advanced features?** Scroll to the **Work With Me** section or use the sidebar contact box to discuss specialized mapping, advanced payment matching, automation, SQL/API integration or deployment. Send questions to **greatfoodprefer@gmail.com**. Do not email confidential workbooks; use mock data when describing your needs.

### Example

| Source A (Orders.xlsx) | Excel column | Source B (Payments.xlsx) | Excel column |
|---|---|---|---|
| Order No | B | Reference | C |
| Total (RM) | E | Paid Amount | F |

1. Start the app, click **Try sample** or upload Orders.xlsx and Payments.xlsx.
2. Select worksheet for each source. CSV will appear as `CSV Data`.
3. Select the first data row (default Row 2 when Row 1 contains headers).
4. Pick **A Matching Key = B**, **B Matching Key = C**.
5. Select **Compare amount / numbers** and **A Value = E**, **B Value = F**.
6. Inspect actual values, empty key counts, duplicates, matched key percentage and the paired sample preview.
7. Click **Run comparison**. Filter results; download a four-sheet Excel report.

### Status meanings

| Status | Interpretation |
|---|---|
| MATCHED | Both sources have that key; selected values agree within tolerance. |
| DIFFERENT | Both sources have the key, but values differ. |
| ONLY IN A | Key exists only in A. |
| ONLY IN B | Key exists only in B. |
| DUPLICATE KEY | Repeated key in at least one side; user review needed unless Group & SUM is explicitly enabled. |
| EMPTY KEY A/B | Row lacks a required key. |
| INVALID VALUE | Unable to parse selected numeric field safely. |

### More examples

- **Text comparison:** compare `Payment Status` against a second system's `Status` using matching transaction IDs.
- **Presence comparison:** compare two ID lists without any amount columns.
- **Split payments:** if B has several payments per Order ID, enable **Group & SUM** only when this is the correct business logic.
- **Composite IDs:** match on Order ID and currency to avoid cross-currency accidental matches.
- **Saved mapping:** download a JSON profile and import it next month. Check column positions before rerunning.

### Known limits

Only structured `.xlsx` and `.csv`; first version compares one value pair at a time. No direct `.xls`, macro execution, password-protected workbooks, OCR/PDF imports, arbitrary merged-header schemas, or automatic foreign-exchange conversion. Formulas are read using their last stored values, if any. Numeric identifiers can lose leading zeros when Excel stores them as numbers (e.g. `0012` becomes `12`). Preview data before running.

## 中文（简体）

**系统用途：** 上传两份 Excel / CSV，找出缺少记录、金额不一致、重复订单和错误值。即使两份文件的 Column Header 不同，也可以直接按 Excel 栏位字母匹配。

**系统默认语言为英文**，可以在左侧栏 `Language / 语言` 切换为简体中文。

**需要定制开发？** 在左侧栏和页面下方的 `定制开发服务` 区块可联系开发者，讨论专属格式、复杂付款规则、批次处理、Python / SQL 自动化或 API 集成。邮箱：**greatfoodprefer@gmail.com**。建议使用虚构资料描述需求，不要直接发送机密财务 Excel。

### 使用流程

1. 启动 App，点击 **试用示例** 或上传两份真实文件。
2. 为 A / B 选择 Worksheet；CSV 的 Sheet 名称会显示 `CSV Data`。
3. 选择数据开始行。例如 Row 1 是 Header，则从 Row 2 开始。
4. 选择 **Excel A Matching Key = B**、**Excel B Matching Key = C**。
5. 比较模式选择「比较金额 / 数字」，A 选 **E**，B 选 **F**。
6. 检查系统显示的 Sample、空编号、重复编号、两份 Excel 的匹配率。
7. 点击「开始比较」查看结果并下载包含四个工作表的报告。

### 结果解读

| Status | 中文解释 |
|---|---|
| MATCHED | 两份 Excel 都找到记录，且比较值一致（误差范围内）。 |
| DIFFERENT | 两边都找到编号，但所选值不一样。 |
| ONLY IN A | 这笔记录只有 Excel A 存在。 |
| ONLY IN B | 这笔记录只有 Excel B 存在。 |
| DUPLICATE KEY | 至少一边有重复编号，需要人工处理。 |
| EMPTY KEY A/B | 缺少匹配编号，不能拿空编号去匹配空编号。 |
| INVALID VALUE | 选择的金额栏位出现不能安全解析的值。 |

### 其他用途

- **文字比较：** 比较两个系统中订单状态（例如 Paid、Pending）。
- **只比较是否存在：** 比较两份 Customer ID、Invoice No 清单，看看遗漏了什么。
- **分期 / 多笔付款：** 若一张订单对应多笔付款，可以选择 Group & SUM，但一定要符合真实业务规则。
- **组合编号：** 如果同一个订单号可能出现在不同币种，可以同时匹配 Order ID + Currency。
- **保存 Mapping Profile：** 下载 JSON，以后重复使用。注意 Excel 插入栏位后，原本 B/E 可能已经改变。

### 注意限制

系统目前不支持 `.xls`、PDF OCR、加密文件、宏、非表格式的复杂合并栏位或自动汇率转换。每次可比较一个值栏位对。对 Excel 公式，读取的是工作簿最后一次保存的缓存计算结果，不会自动重算。若编号以数字形式储存，`0012` 可能已被 Excel 变为 `12`；请在运行前核对预览。
