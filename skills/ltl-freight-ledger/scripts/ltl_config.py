# -*- coding: utf-8 -*-
"""
运输台账 · 路径与配置解析（跨 Agent 可移植）
============================================
任何支持 Agent Skills 标准的 agent（WorkBuddy / Claude Code / Cursor /
Codex / Windsurf ...）运行本 skill 时，都通过这里拿台账路径、备份目录、
输出目录。脚本里**不再出现任何硬编码的绝对路径**，换台机器只改一个
config.json 即可。

查找顺序（先命中者胜出）
------------------------
台账路径 ledger
    1. 命令行 `--ledger`
    2. 环境变量 `LTL_LEDGER`
    3. skill 目录下 `config.json` 的 `ledger`
    4. `~/.workbuddy/secrets/ltl-sites.json` 的 `ledger`（老结构兼容）
    5. 都没有 → 报错并给出建 config.json 的提示

备份目录 backup_dir
    1. `--backup-dir` / `LTL_BACKUP_DIR`
    2. `config.json` 的 `backup_dir`
    3. 默认：`<台账所在目录>/backups`

输出目录 outdir
    1. `--outdir` / `LTL_OUTDIR`
    2. `config.json` 的 `outdir`
    3. 默认：`<当前工作目录>/ltl_output`
    （若解析结果落在 skill 目录内，会自动改到工作目录，避免污染 skill）

时间窗 default_since
    1. `--since`
    2. `config.json` 的 `default_since`
    3. 无（不过滤）
"""
import json
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# skill 根目录（scripts/ 的上一级）
SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 只有真正的 skill 目录（含 SKILL.md）才启用"输出不得落在 skill 内"的保护；
# 工作区里的脚本副本没有 SKILL.md，不应被误判。
_HAS_SKILL_MD = os.path.isfile(os.path.join(SKILL_ROOT, "SKILL.md"))

CONFIG_CANDIDATES = [
    os.path.join(SKILL_ROOT, "config.json"),
    os.path.join(os.path.expanduser("~"), ".workbuddy", "secrets", "ltl-sites.json"),
]

CONFIG_TEMPLATE_HINT = (
    '在本 skill 目录下建 config.json，例如：\n'
    '    { "ledger": "D:/路径/运费台账.xlsx", "default_since": "2026-08-01" }\n'
    '  模板见 assets/config.example.json，skill 目录：%s' % SKILL_ROOT
)


def _load_json(path):
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError) as exc:
        print("提示：配置文件无法读取，已忽略：%s（%s）" % (path, exc))
        return {}
    return data if isinstance(data, dict) else {}


def load_config():
    """合并多个配置文件，靠后的先加载、靠前的覆盖（skill 内 config.json 优先）。"""
    merged = {}
    for p in reversed(CONFIG_CANDIDATES):
        merged.update(_load_json(p))
    return merged


def _pick(cfg, *keys):
    """按顺序取第一个非空值；ledger 兼容 {"path": "..."} 的老结构。"""
    for k in keys:
        if k not in cfg:
            continue
        v = cfg[k]
        if isinstance(v, dict):
            v = v.get("path") or v.get("ledger") or ""
        if v:
            return str(v).strip()
    return ""


def _abs(p):
    return os.path.abspath(os.path.expanduser(str(p).strip()))


def resolve_ledger(cli_value=None, must_exist=True):
    """定位台账 Excel。找不到或文件不存在时直接退出并给出修复指引。"""
    cfg = load_config()
    path = (
        (cli_value or "").strip()
        or os.environ.get("LTL_LEDGER", "").strip()
        or _pick(cfg, "ledger", "ledger_path")
    )
    if not path:
        sys.exit("找不到台账路径。任选一种方式配置：\n"
                 "  1) 命令行：--ledger \"D:/路径/运费台账.xlsx\"\n"
                 "  2) 环境变量：LTL_LEDGER=...\n"
                 "  3) " + CONFIG_TEMPLATE_HINT)
    path = _abs(path)
    if must_exist and not os.path.isfile(path):
        sys.exit("台账文件不存在：%s\n"
                 "请检查 config.json 里的 ledger 路径是否正确。" % path)
    return path


def resolve_backup_dir(cli_value=None, ledger_path=None):
    """定位备份目录；默认放在台账同级 backups/，跟着数据走，最可移植。"""
    cfg = load_config()
    d = (
        (cli_value or "").strip()
        or os.environ.get("LTL_BACKUP_DIR", "").strip()
        or _pick(cfg, "backup_dir")
    )
    if not d:
        if ledger_path:
            d = os.path.join(os.path.dirname(_abs(ledger_path)), "backups")
        else:
            d = os.path.join(os.getcwd(), "ltl_backups")
    return _abs(d)


def resolve_outdir(cli_value=None):
    """定位预览表输出目录；落在 skill 目录内时自动改道，避免污染 skill。"""
    cfg = load_config()
    d = (
        (cli_value or "").strip()
        or os.environ.get("LTL_OUTDIR", "").strip()
        or _pick(cfg, "outdir", "output_dir")
    )
    if not d:
        d = os.path.join(os.getcwd(), "ltl_output")
    d = _abs(d)
    root = os.path.normcase(SKILL_ROOT)
    if _HAS_SKILL_MD and (os.path.normcase(d) == root
                          or os.path.normcase(d).startswith(root + os.sep)):
        d = os.path.join(os.getcwd(), "ltl_output")
        print("提示：输出目录不能放在 skill 目录内，已自动改用 %s" % d)
    return d


def resolve_since(cli_value=None):
    """时间窗；返回 YYYY-MM-DD 字符串或 None（不过滤）。"""
    v = (cli_value or "").strip()
    if v:
        return v
    v = os.environ.get("LTL_SINCE", "").strip()
    if v:
        return v
    return _pick(load_config(), "default_since", "since") or None


if __name__ == "__main__":
    cfg = load_config()
    print("skill 目录        ：%s" % SKILL_ROOT)
    print("合并后的配置键    ：%s" % ", ".join(sorted(cfg.keys())))
    try:
        lg = resolve_ledger()
        print("台账路径          ：%s" % lg)
        print("备份目录          ：%s" % resolve_backup_dir(ledger_path=lg))
    except SystemExit as e:
        print("台账路径          ：未配置（%s）" % e)
    print("输出目录          ：%s" % resolve_outdir())
    print("默认时间窗        ：%s" % resolve_since())
