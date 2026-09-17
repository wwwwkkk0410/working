# -*- coding: utf-8 -*-
"""
运输台账 skill · 跨 Agent 安装器
================================
把本 skill 安装到其他支持 Agent Skills 标准的 AI 工具里，
安装后那些工具也能识别并调用本 skill。

两类目标
--------
A. 有本地技能目录的工具（自动安装）
    claude     Claude Code        ~/.claude/skills/     .claude/skills/
    cursor     Cursor             ~/.cursor/skills/     .cursor/skills/
    codex      Codex (OpenAI)     ~/.codex/skills/      .codex/skills/
    windsurf   Windsurf           ~/.windsurf/skills/   .windsurf/skills/
    agents     通用标准(推荐)      ~/.agents/skills/     .agents/skills/
    workbuddy  WorkBuddy          ~/.workbuddy/skills/  .workbuddy/skills/

B. 只能手动上传的工具（打包一份干净副本，由用户在界面里拖进去）
    doubao     豆包 / 豆包工作
               豆包没有本地技能目录，技能存在云端，必须在
               「技能·连接器·伙伴 → 我的技能 → 新建 → 上传技能」里
               手动上传目录或 zip。本安装器负责生成这份上传包。

用法：
    # 看看能装到哪些地方
    python install_skill.py --list

    # 装到 Claude Code + Cursor + Codex（用户级，所有项目可用）
    python install_skill.py --target claude,cursor,codex

    # 一次装到全部「自动安装」目标
    python install_skill.py --target all

    # 打包一份可直接拖进豆包的副本（默认输出到桌面）
    python install_skill.py --target doubao
    python install_skill.py --target doubao --dest "D:/我的目录"

    # 装到某个项目里（随项目 / git 走，团队共享）
    python install_skill.py --target agents --project "D:/我的项目"

    # Windows 建软链要管理员或开发者模式；默认是复制，最稳
    python install_skill.py --target claude --mode symlink

注意
    自动安装时会排除 config.json / output/ / backups/ 等本机私有内容，
    装过去的是一份干净 skill；在那边首次运行需要按
    assets/config.example.json 建一份 config.json 填台账路径。

    手动上传包（豆包）默认 *包含* config.json —— 因为上传后很难再改嵌套文件，
    而台账路径不含任何密码。如不想带，加 --no-config。
"""
import argparse
import os
import shutil
import sys
import zipfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_NAME = os.path.basename(SKILL_ROOT)

# ---- A 类：有本地技能目录，自动安装 ----
AGENTS = {
    "claude":    {"name": "Claude Code",              "user": "~/.claude/skills",    "project": ".claude/skills"},
    "cursor":    {"name": "Cursor",                   "user": "~/.cursor/skills",    "project": ".cursor/skills"},
    "codex":     {"name": "Codex (OpenAI)",           "user": "~/.codex/skills",     "project": ".codex/skills"},
    "windsurf":  {"name": "Windsurf",                 "user": "~/.windsurf/skills",  "project": ".windsurf/skills"},
    "agents":    {"name": "通用 Agent Skills 标准",   "user": "~/.agents/skills",    "project": ".agents/skills"},
    "workbuddy": {"name": "WorkBuddy",                "user": "~/.workbuddy/skills", "project": ".workbuddy/skills"},
}

# ---- B 类：只能手动上传 ----
MANUAL = {
    "doubao": {
        "name": "豆包 / 豆包工作",
        "dest_default": "~/Desktop",
        "how": "打开豆包 → 左栏「技能·连接器·伙伴」→「我的技能」→ 右上角 + →「上传技能」→ "
               "把生成的目录或 zip 拖进去。上传后新技能需重开会话才生效。",
        "works": "只在「本地电脑模式」下才有机会真正操作台账文件；「云电脑」是独立的 Linux 沙箱，"
                 "读写不到你本机的台账，脚本也跑不了。",
    },
}

# 自动安装时排除的本机私有内容
IGNORE = shutil.ignore_patterns(
    "config.json", "output", "ltl_output", "backups", "ltl_backups",
    "__pycache__", "*.pyc", ".git",
)
# 手动打包时排除的内容（config.json 默认保留，见 --no-config）
IGNORE_NO_CONFIG = shutil.ignore_patterns(
    "config.json", "output", "ltl_output", "backups", "ltl_backups",
    "__pycache__", "*.pyc", ".git",
)
IGNORE_KEEP_CONFIG = shutil.ignore_patterns(
    "output", "ltl_output", "backups", "ltl_backups",
    "__pycache__", "*.pyc", ".git",
)


def _expand(p):
    return os.path.abspath(os.path.expanduser(p))


def list_targets():
    print("skill 源目录：%s" % SKILL_ROOT)
    print("skill 名称  ：%s\n" % SKILL_NAME)
    print("A. 自动安装到本地技能目录")
    print("%-10s %-24s %-28s %s" % ("key", "工具", "用户级目录", "项目级目录"))
    print("-" * 96)
    for key, cfg in AGENTS.items():
        print("%-10s %-24s %-28s %s" % (key, cfg["name"], cfg["user"], cfg["project"]))
    print("\nB. 只能手动上传（打包一份干净副本）")
    print("%-10s %-24s %s" % ("key", "工具", "默认输出目录"))
    print("-" * 96)
    for key, cfg in MANUAL.items():
        print("%-10s %-24s %s" % (key, cfg["name"], cfg["dest_default"]))
    print("\n说明：Cursor 为了兼容，也会读取 ~/.claude/skills 与 ~/.codex/skills，"
          "所以装了 claude/codex 后 Cursor 通常也能看到。")


def _resolve_targets(target_arg):
    keys = [k.strip().lower() for k in (target_arg or "").split(",") if k.strip()]
    if not keys or keys == ["all"]:
        return list(AGENTS.keys())          # all 只针对 A 类
    bad = [k for k in keys if k not in AGENTS and k not in MANUAL]
    if bad:
        sys.exit("未知目标：%s\n可选：%s" % ("、".join(bad), "、".join(list(AGENTS) + list(MANUAL))))
    return keys


def _remove_existing(dst):
    if os.path.islink(dst):
        os.unlink(dst)
    elif os.path.isdir(dst):
        shutil.rmtree(dst)
    elif os.path.exists(dst):
        os.remove(dst)


def install_one(key, dest_dir, mode):
    dst = os.path.join(dest_dir, SKILL_NAME)
    os.makedirs(dest_dir, exist_ok=True)
    if os.path.islink(dst) and _expand(os.readlink(dst)) == _expand(SKILL_ROOT):
        print("  [跳过] 软链已存在：%s -> %s" % (dst, SKILL_ROOT))
        return True

    _remove_existing(dst)

    if mode == "symlink":
        try:
            os.symlink(SKILL_ROOT, dst, target_is_directory=True)
            print("  [OK] 已建软链：%s -> %s" % (dst, SKILL_ROOT))
            return True
        except (OSError, NotImplementedError) as exc:
            print("  [降级] 无法建软链（%s），改为复制。" % exc)
            print("         想用软链请以管理员运行，或开启 Windows 开发者模式。")

    shutil.copytree(SKILL_ROOT, dst, ignore=IGNORE)
    print("  [OK] 已复制到：%s" % dst)
    return True


def _zip_dir(src_dir, zip_path):
    """把 src_dir 打进 zip，zip 内以 src_dir 的名字为根目录。"""
    root_name = os.path.basename(src_dir)
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, _dirs, files in os.walk(src_dir):
            for fn in files:
                full = os.path.join(folder, fn)
                rel = os.path.relpath(full, os.path.dirname(src_dir))
                zf.write(full, rel)


def package_manual(key, dest_root, keep_config=True):
    cfg = MANUAL[key]
    dest_root = _expand(dest_root or cfg["dest_default"])
    os.makedirs(dest_root, exist_ok=True)
    stage = os.path.join(dest_root, SKILL_NAME)
    _remove_existing(stage)
    ignore = IGNORE_KEEP_CONFIG if keep_config else IGNORE_NO_CONFIG
    shutil.copytree(SKILL_ROOT, stage, ignore=ignore)
    zip_path = os.path.join(dest_root, SKILL_NAME + ".zip")
    _zip_dir(stage, zip_path)

    print("  [OK] 上传包已生成：")
    print("       目录：%s" % stage)
    print("       压缩包：%s" % zip_path)
    print("       config.json：%s" % ("已包含（含台账路径，不含任何密码）" if keep_config else "已排除"))
    print("\n  安装步骤")
    print("    %s" % cfg["how"])
    print("  运行环境提醒")
    print("    %s" % cfg["works"])
    return True


def main():
    ap = argparse.ArgumentParser(description="把本 skill 安装到其他 AI 工具")
    ap.add_argument("--target", default="all",
                    help="目标，逗号分隔：%s，或 all" % ",".join(list(AGENTS) + list(MANUAL)))
    ap.add_argument("--project", default=None,
                    help="装成项目级：传项目根目录。不传则装成用户级（所有项目可用）")
    ap.add_argument("--mode", choices=["copy", "symlink"], default="copy",
                    help="copy=复制一份（默认，最稳）；symlink=建软链（改一处到处生效）")
    ap.add_argument("--dest", default=None,
                    help="手动上传包（如豆包）的输出目录，默认见 --list")
    ap.add_argument("--no-config", action="store_true",
                    help="手动上传包里不包含本机 config.json")
    ap.add_argument("--list", action="store_true", help="只列出可选目标")
    args = ap.parse_args()

    if args.list:
        list_targets()
        return

    keys = _resolve_targets(args.target)
    print("skill 源目录：%s" % SKILL_ROOT)
    print("-" * 60)

    ok = 0
    for k in keys:
        if k in MANUAL:
            print("%s（%s）· 手动上传包" % (MANUAL[k]["name"], k))
            try:
                package_manual(k, args.dest, keep_config=not args.no_config)
                ok += 1
            except Exception as exc:      # noqa: BLE001
                print("  [失败] %s：%s" % (type(exc).__name__, exc))
            continue

        cfg = AGENTS[k]
        rel = cfg["project"] if args.project else cfg["user"]
        dest = os.path.join(_expand(args.project), *rel.split("/")) if args.project else _expand(rel)
        print("%s（%s）%s" % (cfg["name"], k, "· 项目级" if args.project else "· 用户级"))
        try:
            install_one(k, dest, args.mode)
            ok += 1
        except Exception as exc:          # noqa: BLE001 - 单个目标失败不影响其他
            print("  [失败] %s：%s" % (type(exc).__name__, exc))

    print("-" * 60)
    print("完成：%d/%d 个目标。" % (ok, len(keys)))
    if any(k in AGENTS for k in keys):
        print("\n首次在别的工具里使用前，记得在该 skill 目录建一份 config.json"
              "（模板：assets/config.example.json），填上台账路径。")
        print("运行环境：Windows + WPS 表格/Excel + openpyxl + pywin32，缺一不可。")


if __name__ == "__main__":
    main()
