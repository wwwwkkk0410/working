---
name: ltl-freight-ledger
description: 每周把 XPO/ABF/TForce/DYLT 四家零担承运商的运单数据登记到本地运费台账 Excel。用户提供 3 个可导出的表格文件（XPO / DYLT / ABF）+ TForce 页面复制文本，自动解析、去重、预览、写入台账。当用户提到"更新台账""运单台账""运费台账""XPO/ABF/TForce/DYLT"时使用。
compatibility: "Windows + WPS 表格或 Excel + Python 3（openpyxl、pywin32）。写入台账依赖 WPS/Excel COM，非 Windows 环境无法写入。豆包（豆包工作）需手动上传，且必须在「本地电脑模式」下运行。"
agent_created: true
---

# 零担运输运费台账自动更新

## 何时使用

用户每周需要更新一张 Excel 运费台账，数据来自四家零担承运商（LTL）。
当用户说"更新这周台账""帮我跑一下运费台账""表格我发你了"时触发本流程。

## 工作模式：用户供货，我记账

**本 skill 不登录任何网站，不需要任何账号密码。**

用户每轮提供：

| 承运商 | 用户提供什么 | 台账登记内容 |
|---|---|---|
| XPO | 导出的 `InvoiceData.csv` | PRO#（A）+ Due date（C）+ 金额（F） |
| DYLT（Daylight） | 导出的 `MyDaylight*.xlsx` | PRO#（A）+ Pick up date（B）+ 金额（F） |
| ABF | 导出的 `Shipments*.xlsx` | PRO#（A）+ 金额（F），**不登记日期** |
| TForce | 页面表格复制成的 TSV 文本 | PRO#（A）+ 金额（F），**不登记日期** |

用户把文件放进工作区 `运费台账/inbox/`，或在对话里直接把文件发给我。
TForce 由于页面不能导出，用户复制页面表格内容贴成 TSV 文件即可。

## 核心原则

1. **台账里的公式和嵌入图片绝对不能破坏**。写入用 WPS COM（`KET.Application`），`openpyxl` 只用于只读分析。
2. **每次写台账前先备份**到 `运费台账/backups/`。
3. **先输出「新增预览表」，用户确认后才落盘**。
4. **各承运商的日期策略不同**：
   - XPO：登记 `Due date`（C 列），存为文本 `09/23/2026`。
   - DYLT：登记 `Pick up date`（B 列），存为真日期。
   - ABF / TForce：**不登记日期**——日期列不碰（其中 B 列可能带公式，绝不能覆盖）。
5. **PRO# 以导出文件为准**：XPO 需从 Invoice 长串数字提取第一段；其余三家直接取 PRO# 列。
6. **去重键是台账 A 列的 PRO#**：已在台账 → 跳过，不重复登记。
7. **只登记已送达（Delivered）的发票**：ABF / DYLT 解析时自动筛掉未送达记录——ABF 按 `Delivery Date` 或 `Delivered To Broker Date` 任一非空判定；DYLT 按 `Status = DELIVERED` 判定；XPO / TForce 导出无配送状态字段，不筛选。被筛掉的未送达单等送达后再随下一轮导出入账。

## 工作流程

### 阶段 1：收料

1. 确认四份输入都到位：XPO、DYLT、ABF 三个文件 + TForce 粘贴文本。
2. 缺哪家就向用户要哪家；三家文件格式不对时给出提示（见"已知坑点"）。

### 阶段 2：解析与去重

使用 `scripts/parse_export.py`：

```bash
python scripts/parse_export.py --carrier XPO    --export "inbox/InvoiceData.csv"
python scripts/parse_export.py --carrier DYLT   --export "inbox/MyDaylight.xlsx"
python scripts/parse_export.py --carrier ABF    --export "inbox/Shipments.xlsx"
python scripts/parse_export.py --carrier TForce --export "inbox/TForce_page.tsv"
```

- 读取 csv / xlsx / tsv，自动识别分隔符，支持 UTF-8 BOM、GBK、负数金额、多种日期格式。
- 时间窗默认取 `config.json` 的 `default_since`（本机为 `2026-08-01`），可用 `--since` 覆盖；
  DYLT/ABF 按 Pickup Date，XPO 按 Due date，TForce 按 Ship Date。
- 用台账 A 列 `PRO#` 去重。
- 输出 `<outdir>/<承运商>_台账待录入_YYYYMMDD.xlsx` 预览表（outdir 默认取 config，缺省是 `<当前目录>/ltl_output`）。
- TForce 预览表额外带 `Pmt`、`BOL` 两列（仅预览，不写入台账）。

### 阶段 3：写入台账

用户确认预览表后，使用 `scripts/write_ledger.py`：

```bash
python scripts/write_ledger.py --carrier XPO    --preview "output/XPO_台账待录入_YYYYMMDD.xlsx"
python scripts/write_ledger.py --carrier DYLT   --preview "output/DYLT_台账待录入_YYYYMMDD.xlsx"
python scripts/write_ledger.py --carrier ABF    --preview "output/ABF_台账待录入_YYYYMMDD.xlsx"
python scripts/write_ledger.py --carrier TForce --preview "output/TForce_台账待录入_YYYYMMDD.xlsx"
```

- 备份原台账到 `<台账所在目录>/backups/`（可用 `--backup-dir` 或 config 的 `backup_dir` 覆盖）。
- 通过 WPS COM（`KET.Application`）打开台账。
- 只写目标列：XPO 写 A/C/F，DYLT 写 A/B/F，ABF/TForce 写 A/F。
- 写入后用 `openpyxl` 只读复核，逐条比对 PRO#/金额/日期。

## 字段映射（已验证）

| 承运商 | 台账 A 列 PRO# | 日期来源 | 台账金额列 F | 说明 |
|---|---|---|---|---|
| XPO | `Invoice` 第一段 = `PRO Number` 去横杠 | `Due date` C 列 | `Invoice Amount` | Invoice 列形如 `273804801 202601 015` |
| DYLT | `Probill` | `pick up date` B 列 | `Total Cost` | 日期为真日期 |
| ABF | `Pro #` | 不登记 | `Gross Amount` | PRO# 可能带前导零，按文本写入 |
| TForce | 页面 `PRO` 列 | 不登记 | 页面 `Balance` 列 | Balance 语义见下 ⚠️ |

### TForce 金额语义（务必清楚）

- **页面 `Balance` = "剩余待支付金额"，不是发票总额**。会随每一笔付款递减。
- 台账里登记的是**原始金额**；**绝不能用页面 Balance 覆盖台账里已有的金额**。
- 页面 `Pmt` 列：`Y` = 已经支付过；空白 = 还没付款记录。
- 判定逻辑只看 PRO#：已在台账就跳过，与 Pmt 状态无关。

## 首次配置（每台机器一次）

台账路径、备份目录、输出目录、默认时间窗**都不写死在代码里**，按以下顺序解析（先命中者胜出）：

1. 命令行参数：`--ledger` / `--backup-dir` / `--outdir` / `--since`
2. 环境变量：`LTL_LEDGER` / `LTL_BACKUP_DIR` / `LTL_OUTDIR` / `LTL_SINCE`
3. **skill 目录下的 `config.json`**（推荐方式）
4. `~/.workbuddy/secrets/ltl-sites.json`（旧结构，向后兼容）

在 skill 目录建一份 `config.json` 即可（模板：`assets/config.example.json`）：

```json
{
  "ledger": "D:/路径/运费台账.xlsx",
  "backup_dir": "D:/路径/backups",
  "outdir": "D:/路径/output",
  "default_since": "2026-08-01"
}
```

`config.json` 属于本机私有配置，已在 `.gitignore` 里排除，**不要提交 git、也不要分享给别人**。

自检命令：`python scripts/ltl_config.py`（打印当前解析到的各项路径）。
解析逻辑集中在 `scripts/ltl_config.py`，要找/改路径只需看这一个文件。

## 跨 Agent 使用（Claude Code / Cursor / Codex / Windsurf / 豆包）

本 skill 遵循 **Agent Skills 开放标准**（`SKILL.md` + `scripts/` + `references/` + `assets/`），
其他支持该标准的工具可以直接识别。

一键安装到别的工具：

```bash
python scripts/install_skill.py --list                        # 看看能装到哪些地方
python scripts/install_skill.py --target claude,cursor,codex  # 用户级：所有项目可用
python scripts/install_skill.py --target all                  # 全部「自动安装」目标
python scripts/install_skill.py --target agents --project "D:/我的项目"   # 项目级：随 git 走
python scripts/install_skill.py --target claude --mode symlink           # 建软链，改一处到处生效
python scripts/install_skill.py --target doubao --dest "D:/上传包目录"    # 豆包：生成手动上传包
```

各工具的技能目录：

| 工具 | 用户级（全局） | 项目级 |
|---|---|---|
| Claude Code | `~/.claude/skills/` | `.claude/skills/` |
| Cursor | `~/.cursor/skills/` | `.cursor/skills/` |
| Codex (OpenAI) | `~/.codex/skills/` | `.codex/skills/` |
| Windsurf | `~/.windsurf/skills/` | `.windsurf/skills/` |
| 通用标准 | `~/.agents/skills/` | `.agents/skills/` |
| WorkBuddy | `~/.workbuddy/skills/` | `.workbuddy/skills/` |
| 豆包 / 豆包工作 | ❌ 无本地目录 | ❌ 无本地目录（只能手动上传，见下） |

> Cursor 为了兼容会额外读取 `~/.claude/skills` 与 `~/.codex/skills`，所以装了 Claude Code
> 或 Codex 后，Cursor 通常也能直接看到本 skill。

### 豆包（豆包工作）

豆包**没有本地技能目录**，技能存在云端，只能手动上传：

1. 生成上传包：`python scripts/install_skill.py --target doubao --dest "<输出目录>"`
   （产出 `ltl-freight-ledger/` 目录 + 同名 `.zip`，默认包含 `config.json`；不想带就加 `--no-config`）
2. 打开豆包 → 左栏「**技能·连接器·伙伴**」→「**我的技能**」→ 右上角 **+** →「**上传技能**」
3. 把上面生成的**目录或 zip 拖进去**；上传后**新技能需重开会话**才生效。

⚠️ **执行环境必须选「本地电脑模式」**。豆包有两个执行环境：

| 环境 | 能否操作本机台账 | 说明 |
|---|---|---|
| 本地电脑模式 | ✅ 有机会 | 运行在你自己电脑上，可访问本地文件与本地软件 |
| 云电脑模式 | ❌ 不行 | 独立的云端 Linux 沙箱（2 核 4G），**读不到本机台账，也没有 WPS COM** |

即使切到本地电脑模式，豆包沙箱里是否装了 `pywin32` + WPS 也需要实测；
若脚本跑不通，本 skill 在豆包侧仍可作为「工作流说明书 + 字段映射参考」被读懂
（`SKILL.md` 与 `references/carriers.md` 本身就写全了四家的口径）。

安装器会自动排除 `config.json` / `output/` / `backups/` / `__pycache__` 等本机私有内容，
装过去的是一份干净 skill；在那边首次运行需要按上面的说明补一份 `config.json`。

**唯一的硬门槛是运行环境**：Windows + WPS 表格（或 Excel）+ `openpyxl` + `pywin32`。
台账里有 WPS 的 `DISPIMG` 嵌入图片和 `_xlfn.XLOOKUP` 外部引用，**必须用 COM 写入**。
因此在 Mac / Linux 上的 agent，本 skill 只能当"操作说明书 + 字段映射参考"来读，无法真正写台账。

## 已知坑点

- **COM 日期时区坑**：不能把 `datetime.date` 直接赋给 COM 单元格，会被时区折算导致日期偏早一天。`pick up date` 列写 `datetime.datetime(year, month, day, 12, 0, 0)`；`Due date` 列直接写文本。
- **COM 读数字返回 float 字符串**：幂等去重时要把单元格值先 `int(float(v))` 再比对。
- **openpyxl 不能保存这个台账**：会破坏 WPS 的 `DISPIMG` 嵌入图片和 `_xlfn.XLOOKUP` 外部引用。
- **导出文件可能是全量历史**：必须按时间窗过滤，不能直接追加。
- **PRO# 前导零**：ABF 部分 PRO# 以 0 开头，写入 A 列时必须用字符串，不能用 `int()`。
- **TForce Balance 不是发票总额**：页面 Balance 是剩余待支付金额，会随付款递减；台账记录原始金额，绝不用 Balance 覆盖已有值。
- **所有路径一律走 `scripts/ltl_config.py`**：脚本里不再有硬编码的绝对路径，换机器只改 `config.json`。若在别处看到写死的 `C:\Users\...` 路径，说明是旧版本，需同步。
- **输出目录不能落在 skill 目录内**：`ltl_config.resolve_outdir` 会自动拦截并改道到工作目录，避免把预览表写进 skill。

## 文件结构

```
ltl-freight-ledger/
├── SKILL.md                  # 本文件
├── config.json               # 【本机私有，不入库】台账/备份/输出路径
├── .gitignore                # 排除 config.json、output/、backups/
├── assets/
│   └── config.example.json   # 配置模板（复制成根目录 config.json）
├── references/
│   └── carriers.md           # 四家承运商导出格式、字段映射、TForce 金额语义详解
└── scripts/
    ├── ltl_config.py         # 路径解析（跨 agent 可移植的核心）
    ├── parse_export.py       # 解析导出文件 → 预览表
    ├── write_ledger.py       # 预览表 → COM 写入台账 + 复核
    └── install_skill.py      # 安装到 Claude Code / Cursor / Codex / Windsurf；并打包豆包上传包
```

需要某家承运商的导出列名、字段映射细节或 TForce 金额语义时，
按需读取 `references/carriers.md`。

## 待扩展

1. 稳定后改为每周定时 automation（用户仍需手动导出文件）。
2. 若用户后续希望自动化取数，再单独评估登录自动化的安全方案。
