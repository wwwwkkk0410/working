# add-shared-skill.ps1
# 用法： powershell -ExecutionPolicy Bypass -File add-shared-skill.ps1 <技能名>
# 作用：把中央目录里的一个技能，链接到本机 Claude Code / Codex / WorkBuddy 的技能目录。
#       豆包无需处理：它的 .user_skills 已整体指向中央 skills 目录，天然可见。
# 注意：config.json 是本机私有配置（不入库），在新机器上需要单独创建/修改。

param(
    [Parameter(Mandatory = $true)]
    [string]$SkillName
)

$ErrorActionPreference = "Stop"

$central = "C:\Users\weiwo\Documents\agent-skills\skills"
$targets = @(
    "C:\Users\weiwo\.claude\skills",
    "C:\Users\weiwo\.codex\skills",
    "C:\Users\weiwo\.workbuddy\skills"
)

$skill = Join-Path $central $SkillName
if (-not (Test-Path (Join-Path $skill "SKILL.md"))) {
    Write-Host "错误：中央目录中不存在技能 '$SkillName'（$skill）" -ForegroundColor Red
    exit 1
}

foreach ($t in $targets) {
    if (-not (Test-Path $t)) {
        Write-Host "跳过：目标目录不存在（$t）" -ForegroundColor Yellow
        continue
    }
    $link = Join-Path $t $SkillName
    if (Test-Path $link) {
        Write-Host "跳过：$link 已存在" -ForegroundColor Yellow
        continue
    }
    New-Item -ItemType Junction -Path $link -Target $skill | Out-Null
    Write-Host "已链接：$link" -ForegroundColor Green
}

Write-Host "完成。豆包已通过 .user_skills junction 自动可见。" -ForegroundColor Green
