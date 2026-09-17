# agent-skills 共享技能仓库

本仓库是 4 个 agent（豆包 / Claude Code / Codex / WorkBuddy）共享技能的**唯一存放处**。
所有 agent 通过 Windows 目录联接（junction）指向本仓库的 `skills\` 目录，因此**改一处，处处生效**。

## 目录结构

```
agent-skills\
├── skills\                      # 所有共享技能，一个技能 = 一个目录
│   ├── ltl-freight-ledger\      # 运费台账登记（含 config.json 本机私有配置）
│   ├── design-taste-frontend\   # 即 WorkBuddy 中的 "taste-skill"（前端设计）
│   ├── impeccable\              # 前端界面设计/评审
│   ├── ponytail\                # 以下 6 个来自 ponytail 插件内置技能
│   ├── ponytail-audit\
│   ├── ponytail-debt\
│   ├── ponytail-gain\
│   ├── ponytail-help\
│   ├── ponytail-review\
│   └── <今后新增技能>
├── add-shared-skill.ps1         # 新增技能时，一条命令链接到各 agent
└── README.md
```

## 本机（当前电脑）已配置的 junction

| Agent | 技能目录 | 方式 |
| --- | --- | --- |
| 豆包 | `...\workspace\.user_skills` | 整个目录 → `skills\`（新增技能自动可见） |
| Claude Code | `~\.claude\skills\<name>` | 每个技能一条 junction |
| Codex | `~\.codex\skills\<name>` | 每个技能一条 junction |
| WorkBuddy | `~\.workbuddy\skills\<name>` | 每个技能一条 junction |

已共享技能及各自 agent 下的链接名：

| 中央目录名 | 豆包 | Claude Code | Codex | WorkBuddy |
| --- | --- | --- | --- | --- |
| ltl-freight-ledger | 自动 | ltl-freight-ledger | ltl-freight-ledger | ltl-freight-ledger |
| design-taste-frontend | 自动 | design-taste-frontend | design-taste-frontend | taste-skill（保留原名） |
| impeccable | 自动 | impeccable | impeccable | impeccable |
| ponytail / ponytail-audit / -debt / -gain / -help / -review | 自动 | 走插件（未链接） | 走插件（未链接） | 同名链接 |

> ponytail 是一个多 agent 插件（Claude Code 4.9.0 / Codex 4.10.0），这里共享的是它内置的 6 个技能；Claude/Codex 已通过插件加载，故不重复链接。插件的命令/钩子等能力不在本仓库范围内。

## 日常使用

**新增/修改技能**：直接在 `skills\<name>\` 里编辑即可，4 个 agent 立即读到最新内容。

**把中央目录里的一个新技能链接到各 agent**：

```
powershell -ExecutionPolicy Bypass -File add-shared-skill.ps1 <技能名>
```

**同步到其他电脑**（本机改动 → 远程 → 其他电脑）：

```
# 本机推送
git add -A
git commit -m "描述改动"
git push

# 其他电脑拉取
git pull
```

## 在新电脑上初始化（一次性）

```
# 1. 克隆本仓库
git clone <远程仓库地址> C:\Users\<你>\Documents\agent-skills

# 2. 豆包（如使用）：把 .user_skills 换成 junction
Rename-Item "...\.doubao\agent_mode\workspace\.user_skills" ".user_skills.bak"
New-Item -ItemType Junction -Path "...\.doubao\agent_mode\workspace\.user_skills" -Target "C:\Users\<你>\Documents\agent-skills\skills"

# 3. 其他 agent：把技能逐个链接进去（可复用仓库里的 add-shared-skill.ps1）
```

## 重要约定

- **`config.json` 是本机私有配置，已被 `.gitignore` 排除，不会入库**（仅 ltl-freight-ledger 含此文件）。每台电脑上都需要自行创建或修改（模板见 `skills\ltl-freight-ledger\assets\config.example.json`）。
- 不要动各 agent 自己原有的、未纳入本仓库的技能目录内容，只新增 junction。
- 检查 junction 是否正常：`cmd /c dir <技能目录>`，看到 `<JUNCTION>` 即正常；若技能不加载，先确认 `git pull` 后中央目录内容完整。
- 备份遗留（确认一切正常后可删）：`.user_skills.bak`（空）、`~\.workbuddy\skills\ltl-freight-ledger.bak`、`~\.claude\skills\design-taste-frontend.bak`（与中央目录内容一致）。
