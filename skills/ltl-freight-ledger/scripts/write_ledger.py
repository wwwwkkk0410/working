# -*- coding: utf-8 -*-
"""
把 parse_export.py 生成的「待录入」预览表写入台账。
走 WPS/Excel COM，保住公式与嵌入图片。
只写目标列，运行前自动备份台账。

用法：
    python write_ledger.py --carrier DYLT --preview output/DYLT_台账待录入_20260911.xlsx
    python write_ledger.py --carrier ABF --preview output/ABF_台账待录入_20260911.xlsx

台账路径与备份目录不写死在代码里，按以下顺序解析（详见 scripts/ltl_config.py）：
    --ledger/--backup-dir 参数  >  环境变量 LTL_LEDGER/LTL_BACKUP_DIR  >  skill 目录 config.json
    备份目录默认是 <台账所在目录>/backups

运行环境：Windows + 已装 WPS 表格或 Excel + pywin32（pip install pywin32）。
    非 Windows 环境无法写入本台账（台账含 WPS 的 DISPIMG 嵌入图与 XLOOKUP，
    用 openpyxl 保存会破坏），此时本脚本会给出明确报错。
"""
import argparse
import datetime
import os
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 让 import 能拿到同一目录下的 parse_export / ltl_config
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ltl_config
from parse_export import CARRIER_PROFILES, clean

# 所有台账表的列结构都一样，列号固定如下：
# A=PRO#, B=pick up date, C=Due date, F=Orig. amount
COL_MAP = {
    ("PRO#", "key"): 1,
    ("pick up date", "date"): 2,
    ("Due date", "date"): 3,
    ("Orig. amount", "amount"): 6,
}


def read_preview(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header = [clean(c) for c in rows[0]]
    recs = []
    for r in rows[1:]:
        if not r:
            continue
        d = dict(zip(header, r))
        if clean(d.get("状态")) != "新增":
            continue
        rec = {"pro": clean(d.get("PRO#", ""))}
        # 日期：可能列名是 pick up date 或 Due date，也可能没有
        for col_name in ("pick up date", "Due date"):
            if col_name in d and d[col_name]:
                v = d[col_name]
                if isinstance(v, datetime.datetime):
                    rec["date"] = v.date()
                elif isinstance(v, datetime.date):
                    rec["date"] = v
                elif isinstance(v, str) and v.strip():
                    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
                        try:
                            rec["date"] = datetime.datetime.strptime(v.strip(), fmt).date()
                            break
                        except ValueError:
                            continue
        # 金额
        amt = d.get("Orig. amount")
        if amt is None or amt == "":
            amt = d.get("Orig. amount")
        try:
            rec["amount"] = float(amt)
        except (TypeError, ValueError):
            rec["amount"] = None
        recs.append(rec)
    return recs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--carrier", required=True)
    ap.add_argument("--preview", required=True, help="parse_export.py 生成的待录入 xlsx")
    ap.add_argument("--ledger", default=None,
                    help="台账 xlsx；不传则走 LTL_LEDGER 环境变量 / config.json")
    ap.add_argument("--backup-dir", default=None,
                    help="备份目录；默认 <台账所在目录>/backups")
    args = ap.parse_args()

    args.ledger = ltl_config.resolve_ledger(args.ledger)
    args.backup_dir = ltl_config.resolve_backup_dir(args.backup_dir, args.ledger)

    profile = CARRIER_PROFILES[args.carrier]
    sheet = profile["sheet"]

    new_records = read_preview(args.preview)
    if not new_records:
        print("预览表里没有「新增」记录，无需写入。")
        return
    print("预览表里有 %d 条新增记录待写入 %s 表。" % (len(new_records), sheet))

    # 确定要写的目标列（列号）
    write_plan = []  # [(列号, rec字段, 是否文本日期, target列名), ...]
    for col_name, use in profile["target"]:
        key = (col_name, use)
        if key not in COL_MAP:
            raise RuntimeError("未知的目标列 %s / %s，请更新 COL_MAP" % (col_name, use))
        col_idx = COL_MAP[key]
        if use == "date":
            # 用户台账里的日期是文本格式（如 09/23/2026），COM 写日期会偏时区，故按文本写
            write_plan.append((col_idx, "date", True, col_name))
        elif use == "amount":
            write_plan.append((col_idx, "amount", False, col_name))
        elif use == "key":
            write_plan.append((col_idx, "pro", False, col_name))

    # ---- 1. 备份 ----
    os.makedirs(args.backup_dir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = os.path.join(args.backup_dir, "%s_%s_台账_%s.xlsx" % (args.carrier, sheet, stamp))
    shutil.copy2(args.ledger, bak)
    print("备份完成：%s" % bak)

    # ---- 2. COM 写入 ----
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        sys.exit("需要 pywin32（仅 Windows）：python -m pip install pywin32\n"
                 "本台账含 WPS 嵌入图片与 XLOOKUP，必须用 WPS/Excel COM 写入，"
                 "openpyxl 会破坏这些内容。")
    pythoncom.CoInitialize()
    app = None
    used_progid = None
    for progid in ("KET.Application", "Ket.Application",
                   "Excel.Application", "ET.Application"):
        try:
            app = win32com.client.DispatchEx(progid)
            used_progid = progid
            break
        except Exception:
            continue
    if app is None:
        print("ERROR: 找不到可用的表格 COM 程序（试过 KET/Excel/ET）")
        sys.exit(1)
    print("COM 程序：%s" % used_progid)

    try:
        app.Visible = False
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        wb = app.Workbooks.Open(args.ledger, UpdateLinks=0, ReadOnly=False)
        ws = wb.Worksheets(sheet)

        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        print("台账 %s 表当前最后一条数据在第 %d 行" % (sheet, last))

        existing = set()
        for r in range(2, last + 1):
            v = ws.Cells(r, 1).Value
            if v is not None and str(v).strip():
                try:
                    existing.add(str(int(float(v))).strip())
                except (ValueError, TypeError):
                    existing.add(str(v).strip())
        print("已有 PRO# %d 个" % len(existing))

        clash = [r for r in new_records if r["pro"] in existing]
        if clash:
            print("ERROR: 以下 PRO# 已在台账中，取消写入：%s"
                  % ", ".join(r["pro"] for r in clash))
            wb.Close(SaveChanges=False)
            sys.exit(1)

        for rec in new_records:
            row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
            for col_idx, field, as_text, col_name in write_plan:
                if field == "pro":
                    # 用字符串写入，保留可能的前导零（如 ABF 的 070951951）
                    ws.Cells(row, col_idx).Value = rec["pro"]
                elif field == "amount":
                    ws.Cells(row, col_idx).Value = rec["amount"]
                elif field == "date":
                    if col_name == "pick up date":
                        # 用户台账里 pick up date 列是真日期，写成 datetime 避免格式不一致
                        ws.Cells(row, col_idx).NumberFormat = "[$-409]yyyy-mm-dd;@"
                        ws.Cells(row, col_idx).Value = datetime.datetime(
                            rec["date"].year, rec["date"].month, rec["date"].day, 12, 0, 0)
                    else:
                        # XPO 的 Due date 列是文本格式
                        ws.Cells(row, col_idx).NumberFormat = "@"
                        ws.Cells(row, col_idx).Value = rec["date"].strftime("%m/%d/%Y")
            print("已写入第 %d 行：PRO#=%s  %s" % (
                row, rec["pro"],
                " ".join("%s=%s" % (cn, rec.get(fn, ""))
                         for _, fn, _, cn in write_plan)))

        newlast = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        wb.Save()
        wb.Close(SaveChanges=False)
        print("保存完成，%s 表数据行数 %d -> %d" % (sheet, last, newlast))
    finally:
        app.Quit()
        pythoncom.CoUninitialize()

    # ---- 3. 用 openpyxl 只读复核 ----
    from openpyxl import load_workbook
    wb2 = load_workbook(args.ledger, read_only=True, data_only=True)
    ws2 = wb2[sheet]
    rows = list(ws2.iter_rows(values_only=True))
    wb2.close()
    ok = 0
    for rec in new_records:
        for r in rows[1:]:
            if r and r[0] is not None and str(r[0]).strip() == rec["pro"]:
                # 按 write_plan 逐项复核
                all_ok = True
                details = []
                for col_idx, field, as_text, col_name in write_plan:
                    val = r[col_idx - 1]
                    if field == "amount":
                        amount_ok = (val is not None and abs(float(val) - rec["amount"]) < 0.005)
                        all_ok = all_ok and amount_ok
                        details.append("amount=%s" % val)
                    elif field == "date":
                        if rec.get("date") is None:
                            date_ok = True
                        else:
                            if isinstance(val, datetime.datetime):
                                val = val.date()
                            date_ok = (val is not None and (
                                str(val)[:10] == rec["date"].isoformat()
                                or str(val) == rec["date"].strftime("%m/%d/%Y")
                                or str(val) == rec["date"].strftime("%Y-%m-%d")))
                        all_ok = all_ok and date_ok
                        details.append("date=%s" % val)
                    elif field == "pro":
                        pro_ok = (str(val).strip() == rec["pro"])
                        all_ok = all_ok and pro_ok
                        details.append("pro=%s" % val)
                print("复核 %s：%s (%s)" % (rec["pro"], "OK" if all_ok else "FAIL", " ".join(details)))
                if all_ok:
                    ok += 1
    print("复核通过 %d/%d" % (ok, len(new_records)))


if __name__ == "__main__":
    main()
