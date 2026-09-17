# -*- coding: utf-8 -*-
"""
运输台账 · 承运商导出文件解析器
=================================
作用：把承运商网站导出的运单/发票明细，转成可以直接录进台账的三列：
      PRO#(发票编号) / 日期 / 金额，并与台账已有记录比对，标出哪些是新增。

已接入：XPO / DYLT(Daylight) / ABF / TForce
TForce 用法：页面表格不能导出，把页面文本粘贴成 TSV 文件即可，
            列名：PRO / BOL / Customer Number / Company / PO / Ship Date / P/C / Curr / Balance / Origin / Dest / Pmt

用法：
    python parse_export.py --carrier XPO --export "<导出的csv或xlsx>" [--since 2026-08-01]

    台账路径与输出目录不写死在代码里，按以下顺序解析（详见 scripts/ltl_config.py）：
      --ledger/--outdir 参数  >  环境变量 LTL_LEDGER/LTL_OUTDIR  >  skill 目录 config.json

输出：
    <outdir>/<承运商>_台账待录入_<日期>.xlsx   三列待录入数据
    控制台报告                                  总数 / 新增 / 已在台账 的统计

跨 Agent：本脚本不依赖 WorkBuddy 专有目录，Claude Code / Cursor / Codex /
    Windsurf 等支持 Agent Skills 标准的 agent 均可直接调用。
"""
import argparse
import csv
import datetime
import os
import re
import sys
import unicodedata

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.exit("需要 openpyxl：pip install openpyxl")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 同一目录下的配置解析模块（台账/输出/时间窗都从这里取，无硬编码路径）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ltl_config

# ----------------------------------------------------------------------------
# 承运商配置：每家公司的差异都集中在这里，加新公司只需要加一段配置
# ----------------------------------------------------------------------------
CARRIER_PROFILES = {
    "XPO": {
        "sheet": "XPO",                     # 台账里的工作表名
        "ledger_key_col": "PRO#",           # 台账里用来判重的列
        # ---- 导出文件的列名 ----
        "export_key_col": "PRO Number",
        "export_invoice_col": "Invoice",    # 长串数字，需要提取
        "export_date_col": "Date Due",
        "export_amount_col": "Invoice Amount",
        # ---- 目标台账列（列名 -> 用途） ----
        "target": [
            ("PRO#", "key"),
            ("Due date", "date"),
            ("Orig. amount", "amount"),
        ],
        "date_format": "%m/%d/%Y",          # 导出文件里的日期格式
    },
    "DYLT": {
        "sheet": "DYLT",
        "ledger_key_col": "PRO#",
        "export_key_col": "Probill",
        "export_invoice_col": "Probill",    # 没有长串数字，直接用 Probill
        "export_date_col": "Pickup Date",
        "export_amount_col": "Total Cost",
        "target": [
            ("PRO#", "key"),
            ("pick up date", "date"),        # Daylight 需要登记 Pickup Date
            ("Orig. amount", "amount"),
        ],
        "delivered": {"col": "Status", "equals": "DELIVERED"},   # 只登记已送达
        "date_format": "%Y-%m-%d",
    },
    "ABF": {
        "sheet": "ABF",
        "ledger_key_col": "PRO#",
        "export_key_col": "Pro #",
        "export_invoice_col": "Pro #",      # 直接用 Pro #
        "export_date_col": "Pickup Date",   # 不登记日期，但用于时间窗筛选
        "export_amount_col": "Gross Amount",
        "target": [
            ("PRO#", "key"),
            ("Orig. amount", "amount"),       # ABF 不登记日期
        ],
        "delivered": {"cols_any": ["Delivery Date", "Delivered To Broker Date"]},  # 只登记已送达
        "date_format": "%Y-%m-%d",
    },
    "TForce": {
        "sheet": "TForce",
        "ledger_key_col": "PRO#",
        "export_key_col": "PRO",
        "export_invoice_col": "PRO",         # 没有 Invoice 长串，直接用 PRO
        "export_date_col": "Ship Date",      # 不登记日期，但用于时间窗筛选
        "export_amount_col": "Balance",
        "target": [
            ("PRO#", "key"),
            ("Orig. amount", "amount"),       # TForce 不登记日期
        ],
        "preview_extras": ["Pmt", "BOL"],     # 仅预览用，标记是否已付 + BOL 便于核对
        "date_format": "%Y-%m-%d",
    },
}

# 已知需要保留原样的台账列（含公式），永不写入
PROTECTED_COLUMNS_HINT = ["pick up date", "平台报价", "比报价高", "开票编号",
                          "Mia 认领", "发票", "报价单", "供应商回复", "处理状态", "处理结果"]


# ----------------------------------------------------------------------------
# 基础解析
# ----------------------------------------------------------------------------
def extract_pro(invoice_text, pro_text=""):
    """从导出文件的发票编号里提取真正的发票号。

    XPO 的 Invoice 列形如 "273804801 202601 015"：
    第 1 段 9 位数字才是发票号，后两段是批次/序号。
    已用 516 条真实数据验证：取第 1 段 == PRO Number 去掉横杠，100% 一致。
    """
    inv = (invoice_text or "").strip()
    first = re.split(r"[\s\-]+", inv)[0] if inv else ""
    if first.isdigit():
        return first
    pro = re.sub(r"[^0-9A-Za-z]", "", (pro_text or ""))
    return pro


def parse_money(s):
    """'$2,347.85' -> 2347.85 ; '($944.41)' -> -944.41 ; '' -> None"""
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace("$", "").replace(",", "").replace(" ", "")
    if not t or t in {"-", "—"}:
        return None
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def parse_date(s, fmt):
    if isinstance(s, datetime.datetime):
        return s.date()
    if isinstance(s, datetime.date):
        return s
    t = str(s or "").strip()
    if not t:
        return None
    for f in (fmt, "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(t[:19], f).date()
        except ValueError:
            continue
    return None


def clean(s):
    if s is None:
        return ""
    return str(s).replace("\u200b", "").strip()


# ----------------------------------------------------------------------------
# 读导出文件（csv / xlsx 都支持）
# ----------------------------------------------------------------------------
def read_export(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = [[clean(c) for c in r] for r in ws.iter_rows(values_only=True)]
        wb.close()
    else:
        # 承运商导出的 csv 常见带 UTF-8 BOM；TForce 粘贴的是 TSV（制表符分隔）
        with open(path, "rb") as f:
            sample = f.read(4096)
        delim = "\t" if sample.count(b"\t") > sample.count(b",") else ","
        for enc in ("utf-8-sig", "utf-8", "gbk", "utf-16"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    rows = [[clean(c) for c in r] for r in csv.reader(f, delimiter=delim)]
                break
            except UnicodeDecodeError:
                continue
        else:
            raise RuntimeError("无法识别导出文件的编码，请另存为 UTF-8 或 xlsx")
    rows = [r for r in rows if any(r)]
    if not rows:
        raise RuntimeError("导出文件是空的")
    header = rows[0]
    return header, rows[1:]


def _delivered_ok(profile, header, row):
    """只登记已送达(Delivered)的发票；未配置 delivered 的承运商不筛选。"""
    cfg = profile.get("delivered")
    if not cfg:
        return True, ""
    if "equals" in cfg:
        col = cfg["col"]
        if col not in header:
            return False, "缺配送状态列 %s" % col
        v = str(row[header.index(col)]).strip().upper() if header.index(col) < len(row) else ""
        return v == str(cfg["equals"]).strip().upper(), "配送状态=%s" % (v or "空")
    if "cols_any" in cfg:
        for col in cfg["cols_any"]:
            if (col in header and header.index(col) < len(row)
                    and row[header.index(col)] is not None
                    and str(row[header.index(col)]).strip()):
                return True, ""
        return False, "未送达"
    return True, ""


def build_records(profile, header, data_rows):
    need = [profile["export_key_col"], profile["export_invoice_col"],
            profile["export_date_col"], profile["export_amount_col"]]
    missing = [c for c in need if c not in header]
    if missing:
        raise RuntimeError(
            "导出文件缺少列：%s\n实际列名：%s" % ("、".join(missing), header))
    ix = {c: header.index(c) for c in need}

    records, skipped = [], []
    for r in data_rows:
        ok, why = _delivered_ok(profile, header, r)
        if not ok:
            skipped.append((r, why or "未送达"))
            continue
        pro_raw = r[ix[profile["export_key_col"]]] if ix[profile["export_key_col"]] < len(r) else ""
        inv_raw = r[ix[profile["export_invoice_col"]]] if ix[profile["export_invoice_col"]] < len(r) else ""
        pro = extract_pro(inv_raw, pro_raw)
        due = parse_date(r[ix[profile["export_date_col"]]], profile["date_format"])
        amt = parse_money(r[ix[profile["export_amount_col"]]])
        if not pro:
            skipped.append((r, "取不到发票编号"))
            continue
        records.append({"pro": pro, "date": due, "amount": amt,
                        "raw_invoice": inv_raw, "raw_pro": pro_raw})
    return records, skipped


# ----------------------------------------------------------------------------
# 读台账已有记录
# ----------------------------------------------------------------------------
def read_ledger_existing(ledger_path, profile):
    wb = load_workbook(ledger_path, read_only=True, data_only=True)
    if profile["sheet"] not in wb.sheetnames:
        wb.close()
        raise RuntimeError("台账里找不到工作表 %s，实际有：%s"
                           % (profile["sheet"], wb.sheetnames))
    ws = wb[profile["sheet"]]
    header = [clean(c) for c in next(ws.iter_rows(values_only=True))]
    if profile["ledger_key_col"] not in header:
        wb.close()
        raise RuntimeError("台账 %s 表里找不到 %s 列" % (profile["sheet"], profile["ledger_key_col"]))
    ki = header.index(profile["ledger_key_col"])
    existing, last_row = set(), 1
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        v = clean(r[ki]) if ki < len(r) else ""
        if v:
            existing.add(v)
            last_row = i
    wb.close()
    return existing, last_row, header


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carrier", default="XPO")
    ap.add_argument("--export", required=True, help="承运商导出的 csv / xlsx")
    ap.add_argument("--ledger", default=None,
                    help="台账 xlsx；不传则依次取 LTL_LEDGER 环境变量 / config.json")
    ap.add_argument("--outdir", default=None,
                    help="预览表输出目录；默认 <当前目录>/ltl_output")
    ap.add_argument("--all", action="store_true", help="输出的表格里包含已在台账中的记录")
    ap.add_argument("--since", default=None,
                    help="只保留日期 >= 该日期的记录，YYYY-MM-DD；也可写 since-last 表示台账已有记录的最新日期之后")
    args = ap.parse_args()

    args.ledger = ltl_config.resolve_ledger(args.ledger)
    args.outdir = ltl_config.resolve_outdir(args.outdir)
    args.since = ltl_config.resolve_since(args.since)

    profile = CARRIER_PROFILES[args.carrier]
    header, data_rows = read_export(args.export)
    records, skipped = build_records(profile, header, data_rows)
    existing, last_row, ledger_header = read_ledger_existing(args.ledger, profile)

    new = [r for r in records if r["pro"] not in existing]
    already = [r for r in records if r["pro"] in existing]

    # ---- 时间窗过滤 ----
    cutoff = None
    if args.since:
        if args.since == "since-last":
            dts = [r["date"] for r in records if r["pro"] in existing and r["date"]]
            cutoff = max(dts) if dts else None
        else:
            cutoff = datetime.datetime.strptime(args.since, "%Y-%m-%d").date()
    if cutoff:
        records = [r for r in records if r["date"] and r["date"] >= cutoff]
        new = [r for r in records if r["pro"] not in existing]
        already = [r for r in records if r["pro"] in existing]

    print("=" * 62)
    print("承运商：%s      台账工作表：%s" % (args.carrier, profile["sheet"]))
    print("导出文件：%s" % os.path.basename(args.export))
    if cutoff:
        print("时间窗  ：>= %s" % cutoff.isoformat())
    print("=" * 62)
    print("导出记录数        ：%d" % len(records))
    print("台账已有 PRO# 数  ：%d（数据写到第 %d 行）" % (len(existing), last_row))
    print("已在台账（跳过）  ：%d" % len(already))
    print("新增（可录入）    ：%d" % len(new))
    if skipped:
        print("无法解析被跳过    ：%d" % len(skipped))
        n_und = sum(1 for _, why in skipped if why == "未送达")
        if n_und:
            print("其中未送达被筛掉  ：%d" % n_und)

    # 校验：面额/日期为空的情况
    bad = [r for r in new if r["amount"] is None or r["date"] is None]
    if bad:
        print("⚠ 新增记录里有 %d 条金额或日期为空，需人工确认：" % len(bad))
        for r in bad[:5]:
            print("    %s  date=%s  amount=%s  raw=%s"
                  % (r["pro"], r["date"], r["amount"], r["raw_invoice"]))

    # 新增按月分布，方便判断该取哪一段
    from collections import Counter
    dist = Counter(r["date"].strftime("%Y-%m") for r in new if r["date"])
    if dist:
        print("\n新增记录的日期按月分布：")
        for k in sorted(dist):
            print("    %s  %4d 条" % (k, dist[k]))

    # 写出
    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    stamp = datetime.date.today().strftime("%Y%m%d")
    out_path = os.path.join(outdir, "%s_台账待录入_%s.xlsx" % (args.carrier, stamp))

    wb = Workbook()
    ws = wb.active
    ws.title = "待录入"
    target_cols = profile["target"]
    preview_extras = profile.get("preview_extras", [])
    ws.append(["状态"] + [c for c, _ in target_cols] + preview_extras + ["导出文件里的原始发票编号"])
    rows_out = (records if args.all else new)
    rows_out = sorted(rows_out, key=lambda r: (r["date"] or datetime.date(1900, 1, 1), r["pro"]))

    def target_value(r, use):
        if use == "key":
            return r["pro"]
        if use == "date":
            return r["date"].strftime("%Y-%m-%d") if r["date"] else ""
        if use == "amount":
            return r["amount"] if r["amount"] is not None else ""
        return ""

    extra_ix = {c: header.index(c) for c in preview_extras if c in header}
    key_ix = header.index(profile["export_key_col"])

    def lookup_extra(pro, c):
        """回到原始 data_rows 找该 PRO 对应的额外字段值，仅预览用。"""
        if c not in extra_ix:
            return ""
        for row in data_rows:
            if (len(row) > key_ix and clean(row[key_ix]) == pro):
                return clean(row[extra_ix[c]]) if extra_ix[c] < len(row) else ""
        return ""

    for r in rows_out:
        ws.append(
            ["已在台账" if r["pro"] in existing else "新增"]
            + [target_value(r, use) for _, use in target_cols]
            + [lookup_extra(r["pro"], c) for c in preview_extras]
            + [r["raw_invoice"]]
        )
    widths = [10, 14, 13, 13] + [12] * len(preview_extras) + [26]
    for i, w in enumerate(widths[:len(target_cols) + 2 + len(preview_extras)], start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"
    wb.save(out_path)

    print("\n已生成：%s" % out_path)
    print("提示：台账里 %d 行以下就是可以追加新记录的位置。" % last_row)


if __name__ == "__main__":
    main()
