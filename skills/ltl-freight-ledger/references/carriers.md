# 四家承运商 · 导出格式与字段映射速查

> 本文件是 `SKILL.md` 的补充参考，供新接手本 skill 的 agent 快速对齐口径。
> 台账固定列位：**A = PRO#**、**B = pick up date**、**C = Due date**、**F = Orig. amount**。

## 总览

| 承运商 | 台账 sheet | 用户提供什么 | PRO# 来源 | 金额来源 | 日期写哪列 |
|---|---|---|---|---|---|
| XPO | `XPO` | 导出的 `InvoiceData*.csv` | `Invoice` 列第一段 | `Invoice Amount` | C 列（Due date，**文本**） |
| DYLT（Daylight） | `DYLT` | 导出的 `MyDaylight*.xlsx` | `Probill` | `Total Cost` | B 列（Pick up date，**真日期**） |
| ABF | `ABF` | 导出的 `Shipments*.xlsx` | `Pro #`（**带前导零**） | `Gross Amount` | 不写日期 |
| TForce | `TForce` | 页面表格复制成的 TSV | `PRO` | `Balance`（⚠️ 见下） | 不写日期 |

## XPO

- 导出文件列名：`PRO Number` / `Invoice` / `Date Due` / `Invoice Amount`
- `Invoice` 形如 `273804801 202601 015`：**第一段 9 位数字才是发票号**，后两段是批次/序号。
- 已用 516 条真实数据验证：取第一段 == `PRO Number` 去掉横杠，**100% 一致**。
- 日期格式 `%m/%d/%Y`，写入 C 列时必须写**文本**（COM 写 datetime 会因时区偏早一天）。

## DYLT（Daylight）

- 导出文件列名：`Probill` / `Pickup Date` / `Total Cost`
- `Probill` 本身就是 PRO#，无需提取。
- 日期格式 `%Y-%m-%d`，写入 B 列用**真日期**，格式 `[$-409]yyyy-mm-dd;@`，
  且值写成 `datetime(y, m, d, 12, 0, 0)`（正午，规避时区折算）。

## ABF

- 导出文件列名：`Pro #` / `Pickup Date` / `Gross Amount`
- **`Pro #` 带前导零**（如 `070951951`）：写入 A 列前必须把单元格 `NumberFormat` 设为 `"@"`，
  值按**字符串**写，**绝对不能 `int()` 转换**，否则前导零丢失。
- 不登记日期（`Pickup Date` 仅用于时间窗筛选）。

## TForce

页面表格**不能导出**，用户复制页面内容粘成 TSV，表头为：

```
PRO  BOL  Customer Number  Company  PO  Ship Date  P/C  Curr  Balance  Origin  Dest  Pmt
```

### ⚠️ 金额语义（最容易踩的坑）

| 概念 | 含义 |
|---|---|
| 页面 `Balance` | **剩余待支付金额**，会随每一笔付款递减，**不是发票原始总额** |
| 台账 `Orig. amount` | 登记的是**原始金额** |
| 页面 `Pmt` | `Y` = 已支付过（至少一笔）；空白 = 尚无付款记录 |

**铁律**：绝不能用页面 `Balance` 覆盖台账里已有的金额。
判重只看 PRO#——**已在台账就跳过，与 Pmt 状态无关**。

### 预览表附带列

TForce 预览表会额外输出 `Pmt` 和 `BOL` 两列，**仅供人工 review**（一眼看出哪些已付），
不写入台账。

## 通用规则

- **去重键**：台账 A 列 `PRO#`（字符串比对，读 COM 单元格时先 `int(float(v))`）。
- **时间窗**：导出文件通常是全量历史，必须按日期过滤；
  DYLT/ABF 按 `Pickup Date`，XPO 按 `Due date`，TForce 按 `Ship Date`。
- **PRO# 一律按字符串处理**，避免前导零与科学计数法问题。
