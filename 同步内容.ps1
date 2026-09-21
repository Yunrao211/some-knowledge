# 同步内容.ps1 — 把 vault 里选定的笔记同步到花园的 content/
#
# 设计原则：
#   1. 白名单式 —— 只有 $Manifest 里列出的文件才进花园（默认拒绝）
#   2. 不修改原笔记 —— 只读 vault、只写 content/，vault 一侧零改动
#   3. 清洗显式化 —— 每条规则在 $Rules 里，运行后打印实际改了什么
#   4. 链接自洽 —— 指向未发布笔记的 [[wiki 链接]] 转纯文本，不留死链；
#                  指向已发布笔记的链接归一为裸文件名（Quartz 按文件名解析）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File 同步内容.ps1
#   powershell -ExecutionPolicy Bypass -File 同步内容.ps1 -WhatIf   # 只预览不写入

[CmdletBinding()]
param(
    # vault 根目录（留空则自动推断：本脚本位于 <vault>/share-knowledge/quartz/）
    [string]$VaultRoot,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 下 $PSScriptRoot 在 param 默认值里可能为空，故在正文解析
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $VaultRoot) { $VaultRoot = (Resolve-Path (Join-Path $ScriptDir '..\..')).Path }

# ── 清单：vault 源文件 → content/ 目标相对路径 ────────────────────
$Manifest = @(
    @{ Src = '数学基础\张宇1000题\概念\不等式积累.md';         Dst = '张宇1000题\概念\不等式积累.md' }
    @{ Src = '数学基础\张宇1000题\概念\三角函数公式积累.md';   Dst = '张宇1000题\概念\三角函数公式积累.md' }
    @{ Src = '数学基础\张宇1000题\概念\数列通项与求和.md';     Dst = '张宇1000题\概念\数列通项与求和.md' }
    @{ Src = '数学基础\张宇1000题\概念\数学归纳法.md';         Dst = '张宇1000题\概念\数学归纳法.md' }
    @{ Src = '数学基础\张宇1000题\概念\极坐标.md';             Dst = '张宇1000题\概念\极坐标.md' }
    @{ Src = '数学基础\张宇1000题\概念\柯西-施瓦茨不等式.md';   Dst = '张宇1000题\概念\柯西-施瓦茨不等式.md' }
)

# ── 文本清洗规则（有序执行）────────────────────────────────────
$Rules = @(
    @{ P = '（陈先生\s*';                     R = '（';                  Note = '去署名（保留括号）' }
    @{ P = '陈先生\s*(?=\d{4}-\d{2}-\d{2})';  R = '';                    Note = '去署名（保留日期）' }
    @{ P = '陈先生';                           R = '';                    Note = '去署名（兜底）' }
    @{ P = '^> (\d{4}-\d{2}-\d{2}) 说明：';    R = '> **说明**（$1）：';  Note = '行首裸日期改为「说明」' }
)

# ── 已发布笔记名集合（链接判定基准）──────────────────────────────
$PublishedNames = @($Manifest | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Dst) })

# ── 执行 ─────────────────────────────────────────────────────────
if (-not (Test-Path -LiteralPath $VaultRoot)) { throw "vault 根目录不存在：$VaultRoot" }
$ContentDir = Join-Path $ScriptDir 'content'
$report = @()
$warn   = @()
$droppedAll = @{}

Write-Host "vault  : $VaultRoot"
Write-Host "content: $ContentDir"
Write-Host "已发布笔记 $($PublishedNames.Count) 篇：$($PublishedNames -join ' / ')"
Write-Host ""

# 链接正则：[[目标]] / [[目标|别名]] / [[目标\|别名]]（后者用于表格内转义）
$LinkPattern = '\[\[([^\]\|\\]+?)(\\?\|([^\]]*))?\]\]'

foreach ($item in $Manifest) {
    $srcPath = Join-Path $VaultRoot $item.Src
    $dstPath = Join-Path $ContentDir $item.Dst

    if (-not (Test-Path -LiteralPath $srcPath)) { $warn += "源文件缺失，跳过：$($item.Src)"; continue }

    $text = [System.IO.File]::ReadAllText($srcPath, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrEmpty($text)) { $warn += "源文件为空，跳过（防覆盖）：$($item.Src)"; continue }

    $orig = $text
    $hits = @()

    # 1) 文本清洗
    foreach ($rule in $Rules) {
        $n = ([regex]::Matches($text, $rule.P, 'Multiline')).Count
        if ($n -gt 0) { $hits += "$($rule.Note) ×$n"; $text = [regex]::Replace($text, $rule.P, $rule.R, 'Multiline') }
    }

    # 2) 链接处理（逐条判定 + 手工拼接，避免委托作用域问题）
    $matches = [regex]::Matches($text, $LinkPattern)
    if ($matches.Count -gt 0) {
        $sb = New-Object System.Text.StringBuilder
        $pos = 0; $keptN = 0; $normN = 0; $dropN = 0
        foreach ($m in $matches) {
            [void]$sb.Append($text.Substring($pos, $m.Index - $pos))
            $target = $m.Groups[1].Value
            $sep    = $(if ($m.Groups[2].Success) { $m.Groups[2].Value } else { '' })   # '|别名' 或 '\|别名'
            $alias  = $m.Groups[3].Value
            $base   = ($target -split '/')[-1].Trim()

            if ($PublishedNames -contains $base) {
                if ($target -ne $base) { $normN++ } else { $keptN++ }
                [void]$sb.Append($(if ($sep) { "[[$base$sep]]" } else { "[[$base]]" }))
            } else {
                $dropN++
                $droppedAll[$base] = 1 + [int]$droppedAll[$base]
                [void]$sb.Append($(if ($alias) { $alias } else { $base }))
            }
            $pos = $m.Index + $m.Length
        }
        [void]$sb.Append($text.Substring($pos))
        $text = $sb.ToString()
        if ($keptN -gt 0) { $hits += "保留内链 ×$keptN" }
        if ($normN -gt 0) { $hits += "内链归一为裸文件名 ×$normN" }
        if ($dropN -gt 0) { $hits += "外链转纯文本 ×$dropN" }
    }

    if (-not $WhatIf) {
        $dstDir = Split-Path $dstPath -Parent
        if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
        [System.IO.File]::WriteAllText($dstPath, $text, (New-Object System.Text.UTF8Encoding($false)))
    }

    $report += [PSCustomObject]@{
        目标 = $item.Dst
        字符 = $text.Length
        处理 = $(if ($orig -ne $text) { '已清洗' } else { '原样' })
        改动 = ($hits -join '；')
    }
}

Write-Host "=== 同步结果 ==="
$report | Format-Table -AutoSize -Wrap

if ($droppedAll.Count -gt 0) {
    Write-Host "=== 转为纯文本的链接目标（这些笔记未发布）==="
    $droppedAll.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { Write-Host ("  {0,3} 次  {1}" -f $_.Value, $_.Key) }
    Write-Host "  → 想恢复成可点链接，就把这些笔记加进 Manifest 一并发布。"
}

if ($warn.Count -gt 0) { Write-Host "=== 警告 ==="; $warn | ForEach-Object { Write-Host "  ! $_" } }

# ── 断链自检 ─────────────────────────────────────────────────────
Write-Host "=== 断链自检（content/ 里是否仍有指向未发布笔记的链接）==="
$pub = Get-ChildItem $ContentDir -Recurse -File -Filter '*.md' -EA SilentlyContinue
if ($pub) {
    $bad = Select-String -Path $pub.FullName -Pattern '\[\[([^\]\|\\]+?)(?:\\?\|[^\]]*)?\]\]' -Encoding utf8 -AllMatches -EA SilentlyContinue |
        ForEach-Object { foreach ($m in $_.Matches) { [PSCustomObject]@{ F = $_.Filename; T = ($m.Groups[1].Value -split '/')[-1] } } } |
        Where-Object { $PublishedNames -notcontains $_.T }
    if ($bad) { $bad | Group-Object T | ForEach-Object { Write-Host "  ! 悬空链接 → $($_.Name)  ($($_.Count) 处)" } }
    else { Write-Host "  无 ✅" }
} else { Write-Host "  content/ 为空" }

if ($WhatIf) { Write-Host "`n（-WhatIf 预览模式，未写入任何文件）" }
