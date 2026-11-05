# 用 PowerPoint 做验证（COM 自动化）

PowerPoint 对非法 XML 是**静默丢弃**：不报错、不出「修复」提示，只是效果没了。
所以唯一可信的判据是「让 PowerPoint 自己读回来，和写进去的一致」。

LibreOffice 只能验静态版式——它不解析 `p:timing`，也不渲染文字渐变。

## 1. 打开文件的两种方式

```powershell
# 只读校验（不改文件）
$pres = $app.Presentations.Open($path, $true,  $false, $true)

# 需要 SaveCopyAs / CreateVideo 时必须可写
$pres = $app.Presentations.Open($path, $false, $false, $true)
```

四个参数是 `(路径, ReadOnly, 不显示窗口?, 不加载外部链接?)`。**第 2 个参数是最常踩的坑**——
只读打开时 `CreateVideo` 和 `SaveCopyAs` 会失败。

永远配 `finally`：

```powershell
finally { if ($pres) { try { $pres.Close() } catch {} }
          if ($app)  { try { $app.Quit() }  catch {} } }
```

否则会留一个后台 `POWERPNT.EXE` 咬住文件，下次打开报「文件被占用」。

## 2. 读回动画

```powershell
$seq = $pres.Slides.Item($i).TimeLine.MainSequence
"count=$($seq.Count)"
for ($k = 1; $k -le $seq.Count; $k++) {
    $e = $seq.Item($k)
    "  [$k] EffectType=$([int]$e.EffectType) dur=$($e.Timing.Duration)s delay=$($e.Timing.TriggerDelayTime)s"
}
```

`EffectType` 对应 presetID：`1=Appear`、`2=Fly`、`10=Fade`、`22=Wipe`。
`Duration` 是**秒**（我写进去的是毫秒，除以 1000 后应完全吻合）。

**数量必须等于形状数。** 如果你写了 8 个形状却读回 16 个，说明 `p:timing` 少了一层
`<p:par>` 容器——见 `references/ooxml-effects.md` §3。

## 3. 读回切换

```powershell
[int]$pres.Slides.Item($i).SlideShowTransition.EntryEffect
```

`0` 表示无切换。对照表见 `references/ooxml-effects.md` §2。
注意：`EntryEffect = 0` 也可能是「节点被丢弃」，此时节点名基本是错的
（这就是 `p14:morph` 的翻车现场）。

## 4. 导出图片做视觉验证

```powershell
$pres.Slides.Item(1).Export("out.png", "PNG", 1600, 900)
```

用于验文字渐变、字体替换这类 LibreOffice 渲染不准的东西。

## 5. 导出视频

见 `scripts/pptx2video.ps1`。要点复述：可写打开、异步轮询、finally 清理。
视频里跑的就是文件里的时间轴，所以「视频动起来了」等价于「动画写对了」。

## 6. 写坏了怎么定位

**让 PowerPoint 自己生成一份参照，再逐字对比。** 这是最有效的手段，我用它定位过两个坑：

```powershell
# 让 PowerPoint 自己设切换，保存后读它的 XML
$sl.SlideShowTransition.EntryEffect = 3849
$pres.SaveCopyAs("ref.pptx", 24)      # 24 = ppSaveAsOpenXMLPresentation
```

`SaveCopyAs` 的第二个参数必须是 **24**（OOXML）。传 `1` 会存成旧的 `.ppt` 二进制，
然后你会对着一堆读不出来的文件怀疑人生。

拿到参照后：

```python
import zipfile, re
x = zipfile.ZipFile("ref.pptx").read("ppt/slides/slide1.xml").decode()
print(re.search(r"<p:timing.*</p:timing>", x, re.S).group(0))
```

再和自己的产物 diff。**不要凭记忆猜 schema**——我猜错过两次，每次都是静默失败。
