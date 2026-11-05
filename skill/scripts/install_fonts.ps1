# 安装 skill 自带字体（OFL 授权，可随 skill 分发）
#
#   .\install_fonts.ps1            # 安装 assets/fonts 下所有字体
#   .\install_fonts.ps1 -Check     # 只检查缺哪些，不装
#   .\install_fonts.ps1 -List      # 列出主题里用到哪些字体族
#
# 管理员权限 → 装到 C:\Windows\Fonts（全机可用，立即生效）
# 普通权限  → 装到 %LOCALAPPDATA%\Microsoft\Windows\Fonts（仅本用户，需重启 PowerPoint 或重新登录）
#
# 为什么要装：pptx 里只存字体名，不存字形。目标机器没这个字体，PowerPoint 会静默替代，
# 字宽变化会让排好的版面走形。装不了就必须改用系统自带字体（见 themes.json 的 fonts_fallback）。

[CmdletBinding()]
param(
    [switch]$Check,
    [switch]$List
)

$ErrorActionPreference = 'Stop'
$fontDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'assets\fonts'
$themesJson = Join-Path (Split-Path -Parent $PSScriptRoot) 'assets\themes.json'

# 主题里引用的字体族 → 判断是否已安装时用的族名
$families = @(
    @{ file = 'InstrumentSerif-Regular.ttf';  family = 'Instrument Serif';   role = 'display 衬线'  },
    @{ file = 'InstrumentSerif-Italic.ttf';   family = 'Instrument Serif';   role = 'display 斜体'  },
    @{ file = 'BebasNeue-Regular.ttf';        family = 'Bebas Neue';         role = 'poster 窄体'   },
    @{ file = 'Barlow-Regular.ttf';           family = 'Barlow';             role = 'body 无衬线'   },
    @{ file = 'BarlowCondensed-Regular.ttf';  family = 'Barlow Condensed';   role = 'condensed 无衬线' },
    @{ file = 'IBMPlexMono-Regular.ttf';      family = 'IBM Plex Mono';      role = 'label 等宽'    }
)

function Get-InstalledFamilies {
    Add-Type -AssemblyName System.Drawing
    (New-Object System.Drawing.Text.InstalledFontCollection).Families |
        ForEach-Object { $_.Name }
}

function Show-Check {
    $have = Get-InstalledFamilies
    Write-Host "已安装字体族总数: $($have.Count)" -ForegroundColor DarkGray
    $missing = @()
    foreach ($f in $families) {
        $ok = $have -contains $f.family
        $tag = if ($ok) { 'OK  ' } else { 'MISS' }
        $color = if ($ok) { 'Green' } else { 'Yellow' }
        Write-Host ("  [{0}] {1,-22} {2}" -f $tag, $f.family, $f.role) -ForegroundColor $color
        if (-not $ok) { $missing += $f }
    }
    return $missing
}

if ($List) {
    Write-Host "主题字体族（$themesJson）:" -ForegroundColor Cyan
    if (Test-Path $themesJson) {
        $doc = Get-Content $themesJson -Raw -Encoding UTF8 | ConvertFrom-Json
        $doc.themes | Where-Object { $_.ppt_safe -eq $true } | ForEach-Object {
            $fn = $_.fonts
            Write-Host ("  {0,-20} display={1} / {2}   body={3} / {4}   mono={5}" -f `
                $_.slug, $fn.display_latin, $fn.display_ea, $fn.body_latin, $fn.body_ea, $fn.mono)
        }
    } else {
        Write-Host "  找不到 themes.json，先运行 scripts/build_themes.py" -ForegroundColor Yellow
    }
    return
}

if ($Check) {
    $missing = Show-Check
    if ($missing.Count -eq 0) {
        Write-Host "`n全部已安装。" -ForegroundColor Green
    } else {
        Write-Host "`n缺 $($missing.Count) 个字体文件。运行不带 -Check 即可安装。" -ForegroundColor Yellow
    }
    return
}

# ------------------------------------------------------------------ 安装
if (-not (Test-Path $fontDir)) { throw "找不到字体目录: $fontDir" }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    $target = Join-Path $env:SystemRoot 'Fonts'
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    Write-Host "管理员模式：装到 $target（全机可用）" -ForegroundColor Cyan
} else {
    $target = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
    $regPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
    Write-Host "普通权限：装到 $target（仅本用户）" -ForegroundColor Yellow
    Write-Host "  提示：装完需要重启 PowerPoint；仍不生效就注销再登录一次。" -ForegroundColor Yellow
}
New-Item -ItemType Directory -Force -Path $target | Out-Null
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }

# AddFontResourceW 让当前会话立即可用；WM_FONTCHANGE 通知其他进程
$sig = @'
[DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResourceW(string lpFileName);
[DllImport("user32.dll")] public static extern int SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@
Add-Type -MemberDefinition $sig -Name Native -Namespace FontInst -ErrorAction SilentlyContinue

foreach ($f in $families) {
    $src = Join-Path $fontDir $f.file
    if (-not (Test-Path $src)) {
        Write-Host "  跳过（skill 里没有这个文件）: $($f.file)" -ForegroundColor DarkGray
        continue
    }
    $dst = Join-Path $target $f.file
    if (-not (Test-Path $dst) -or (Get-FileHash $src).Hash -ne (Get-FileHash $dst).Hash) {
        Copy-Item $src $dst -Force
    }
    # 注册表值名必须形如 "Family (TrueType)"，值是文件名
    $suffix = if ($f.file -match 'Italic') { ' Italic (TrueType)' } else { ' (TrueType)' }
    New-ItemProperty -Path $regPath -Name ($f.family + $suffix) -Value $f.file `
        -PropertyType String -Force | Out-Null
    [void][FontInst.Native]::AddFontResourceW($dst)
    Write-Host "  装: $($f.family)"
}

$r = [UIntPtr]::Zero
[void][FontInst.Native]::SendMessageTimeout([IntPtr]0xffff, 0x001D, [UIntPtr]::Zero,
                                            [IntPtr]::Zero, 2, 1500, [ref]$r)

Write-Host "`n--- 复查 ---" -ForegroundColor Cyan
$missing = Show-Check
if ($missing.Count -eq 0) {
    Write-Host "`n安装完成，可直接用主题里的字体排版。" -ForegroundColor Green
} else {
    Write-Host "`n已写入磁盘和注册表，但本进程的字体缓存还没刷新（$($missing.Count) 个仍显示 MISS）。" -ForegroundColor Yellow
    Write-Host "处理办法任选其一：" -ForegroundColor Yellow
    Write-Host "  1) 关闭并重新打开 PowerPoint，再跑一次 -Check"
    Write-Host "  2) 注销后重新登录（按用户安装必须走这一步）"
    Write-Host "  3) 用管理员权限重跑本脚本（装到 C:\Windows\Fonts，立即生效）"
    Write-Host "在此期间生成 pptx 仍然正确——只是渲染预览会用替代字体。"
}
