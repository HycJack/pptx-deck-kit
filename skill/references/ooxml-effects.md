# OOXML 效果与动画：生成器写不出来的那一层

python-pptx 与 genppt 都只覆盖了 OOXML 的一小半。下面每个片段都是**实测通过**的
（PowerPoint 回读一致，或源码已内置），可以直接拼进 slideN.xml 或改写进生成器。

## 1. 形状的四个特效

### 渐变填充 `a:gradFill`

```xml
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="2E4A2A"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="E89CB1"><a:alpha val="60000"/></a:srgbClr></a:gs>
  </a:gsLst>
  <a:lin ang="2700000" scaled="0"/>            <!-- 45°；径向换成 <a:path path="circle"> -->
</a:gradFill>
```

`pos` 是 0–100000（千分之一百分比）；角度 `ang` 是 1/60000 度；透明度的 `alpha` 挂在 **每个 `a:gs`** 内的颜色上，不是渐变上。

**`spPr` 里的顺序是 `xfrm → fill → ln → effectLst`。** 形状级 `a:gradFill` 必须排在 `a:ln` 之前，
顺序反了不报错，只是渐变消失——和 §4 文字渐变的坑是同一个病。

python-pptx 两个附带问题：`add_shape` 会写 `<p:style>`（`lnRef` 指向主题色），
只加 `a:ln noFill` 还留着主题描边的壳，要连 `<p:style>` 一起删；
`font.name` 只写 `a:latin`，中文要自己补 `a:ea`，否则一行里两种字形。

### 阴影 `a:effectLst`

```xml
<a:effectLst>
  <a:outerShdw blurRad="254000" dist="127000" dir="5400000" algn="ctr" rotWithShape="0">
    <a:srgbClr val="1A1A17"><a:alpha val="18000"/></a:srgbClr>
  </a:outerShdw>
</a:effectLst>
```

`blurRad` / `dist` 单位 EMU（1 pt = 12700），`dir` 是 1/60000 度（90° = 5400000 = 正下方）。
`algn="ctr"` 才让阴影按形状中心偏移，`algn="tl"` 会偏出一截。

### 无边框

```xml
<a:ln><a:noFill/></a:ln>
```

**必须显式写。** 完全省略 `<a:ln>` 时 PowerPoint 会套主题默认描边（Office 蓝 `2F5496`）。
这是「细线莫名其妙变蓝」最常见的根因。

### 图层顺序

没有属性可以控制，**`p:spTree` 里的元素顺序就是叠压顺序**。
置底 = 把元素移动到 `spTree` 第 3 个位置（前两个是 `p:nvGrpSpPr`、`p:grpSpPr`）：

```python
spTree.remove(el)
spTree.insert(2, el)
```

## 2. 幻灯片切换 `p:transition`

位置固定：`p:cSld`、`p:clrMapOvr` 之后，`p:timing` 之前。

```xml
<p:transition spd="med"><p:fade/></p:transition>
```

PowerPoint 16.0 实测的编号↔标签对照（写作时可直接抄 `p:` 标签）：

| EntryEffect | 标签 | EntryEffect | 标签 |
|---|---|---|---|
| 0 | 无 | 3851 | `<p:plus/>` |
| 257 | `<p:cut/>` | 3855 | `<p:push dir="u"/>` |
| 2818 | `<p:wipe dir="u"/>` | 3856 | `<p:wedge/>` |
| 3845 | `<p:circle/>` | 3857 | `<p:wheel spokes="1"/>` |
| 3846 | `<p:diamond/>` | 3850 | `<p:newsflash/>` |
| 3849 | `<p:fade/>` | 3844 | `<p:cut/>`（同 257） |

**没有 morph。** `<p14:morph option="byObject"/>` 写进去后 PowerPoint 打开再另存会丢掉节点，
`EntryEffect` 扫到 3890 也没有对应值。不要写它。

## 3. 入场动画 `p:timing`

层级是四层，**少一层就会被数成两个效果**：

```
p:timing / p:tnLst / p:par
└ p:cTn nodeType="tmRoot"
  └ p:seq / p:cTn nodeType="mainSeq"
    └ p:par  ← 一次点击一组（click 组）
      │        stCondLst: delay=indefinite + evt=onBegin/tn=2
      └ p:par  ← 组内容器（这一层必须存在）
        └ p:par / p:cTn presetID=… presetClass="entr"  ← 一个效果
          ├ p:set        把 style.visibility 设成 visible
          └ p:animEffect 真正的过渡
```

实测：漏掉「组内容器」时，8 个形状被 PowerPoint 数成 16 个动画，且**不报错**。

已验证的 `presetID`：

| 效果 | presetID | subtype | 行为节点 |
|---|---|---|---|
| 出现 Appear | 1 | 0 | 只有 `p:set` |
| 飞入 Fly | 2 | 4 | `p:anim` on `ppt_x` 从 `#ppt_x-1/3` 回到原位 + fade |
| 淡入 Fade | 10 | 0 | `p:set` + `animEffect filter="fade"` |
| 擦除 Wipe | 22 | 4 | `p:set` + `animEffect filter="wipe(up)"` |

写法：`nodeType="clickEffect"`（组内第一个）与 `"withEffect"`（其余），
每个效果外层 `p:cond delay="毫秒"` 决定错峰。
最后补 `<p:bldLst><p:bldP spid="…" grpId="0"/></p:bldLst>`，
`spid` 必须等于形状的 `p:cNvPr id`。

## 4. 文字渐变

`a:gradFill` 可以直接放在 `a:rPr` 里，但**顺序有严格要求**：
`ln → fill → effectLst → highlight → uLnTx… → latin → ea → cs → sym`。

把 `a:gradFill` `append` 到 `a:latin` 后面，PowerPoint 会静默忽略它（标题仍是纯色）。

另一个反直觉的点：**LibreOffice 不渲染文字渐变**。同一份文件 LO 导出是纯黑标题，
PowerPoint 导出是渐变。所以文字效果只能用 PowerPoint 验证（`Slide.Export`）。

## 5. 验证顺序

| 要验什么 | 用什么 | 为什么不能用另一个 |
|---|---|---|
| 静态版式、配色、留白 | LibreOffice 渲染 PNG | —— |
| 文字渐变、字体替换 | **PowerPoint** `Slide.Export` | LO 不支持文字渐变 |
| 动画数量、类型、时长 | **PowerPoint** COM `MainSequence` | LO 根本不解析 timing |
| 切换效果 | **PowerPoint** COM `EntryEffect` | 同上 |
| 动起来的样子 | PowerPoint `CreateVideo` | 只能出静态帧的工具没意义 |

LaTeX 式的「能打开就算过」在这里不成立：PowerPoint 对非法节点是**静默丢弃**，
不报错、不修复提示，只是效果没了。所以判据是「读回来和写进去一致」。
