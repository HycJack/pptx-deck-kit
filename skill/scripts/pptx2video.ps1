<#
.SYNOPSIS
    用 PowerPoint COM 把 pptx 导出成 mp4。

.DESCRIPTION
    LibreOffice 只能渲染静态 PNG，不播动画也不出视频，所以这条路只有 PowerPoint 能走。
    三个必须做对的地方：
      1. Presentations.Open 的第 2 个参数（ReadOnly）必须为 $false，否则 CreateVideo 失败；
      2. CreateVideo 是异步的，返回不代表编码完成，只能轮询输出文件大小；
      3. 必须在 finally 里 Close + Quit，否则会留下一个后台 PowerPoint 进程咬住文件。

.PARAMETER SlideSeconds
    每页默认停留秒数。实际时长会更长，因为动画本身要跑完。

.PARAMETER Height
    分辨率高度，720 得到 1280x720。

.EXAMPLE
    ./pptx2video.ps1 -Input deck.pptx -Output deck.mp4
    ./pptx2video.ps1 -Input deck.pptx -Output deck.mp4 -SlideSeconds 2 -Height 1080
#>
param(
    [Parameter(Mandatory = $true)][string]$Input,
    [string]$Output,
    [int]$SlideSeconds = 3,
    [int]$Height = 720,
    [int]$TimeoutSec = 300
)

$ErrorActionPreference = "Stop"
$src = (Resolve-Path $Input).Path
if (-not $Output) {
    $Output = [System.IO.Path]::ChangeExtension($src, ".mp4")
}
$out = [System.IO.Path]::GetFullPath($Output)
if (Test-Path $out) { Remove-Item $out -Force }

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$app = $null
$pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $app.DisplayAlerts = 1
    Start-Sleep -Milliseconds 400

    # ReadOnly 必须是 $false
    $pres = $app.Presentations.Open($src, $false, $false, $true)

    # 参数顺序：(文件名, 是否用旁白, 每页秒数, 分辨率, 帧率, 质量)
    $pres.CreateVideo($out, $true, $SlideSeconds, $Height)

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastSize = -1
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 1500
        if (Test-Path $out) {
            $size = (Get-Item $out).Length
            if ($size -gt 0 -and $size -eq $lastSize) { break }
            $lastSize = $size
        }
    }

    if (Test-Path $out) {
        $mb = (Get-Item $out).Length / 1MB
        "OK  {0}  {1:N1} MB  用时 {2:N1}s" -f $out, $mb, $sw.Elapsed.TotalSeconds
    }
    else {
        Write-Error "超时未生成视频：$out"
    }
}
finally {
    if ($pres) { try { $pres.Close() } catch { } }
    if ($app) { try { $app.Quit() } catch { } }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
