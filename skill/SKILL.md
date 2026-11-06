---
name: generating-pptx-decks
description: Use when asked to produce a .pptx deck (演示文稿 / 幻灯片 / PPT / deck) from scratch or from existing content, or when a generated deck "looks like a Word document", animations don't play, or a deck must be verified rather than merely opened. Covers choosing a generator (Go genppt, python-pptx, raw OOXML), design quality (type scale, palette, grid, px→pt), two orthogonal axes — 14 themes derived from 34 HTML templates (assets/themes.json + OFL fonts) and 13 structurally different page layouts (column ratios, full-bleed, huge numerals, tables, timelines), Chinese/CJK typography (character width = font size, line-height percentage conversion, 避头尾, CJK-Latin spacing, a:ea), effects (gradient, shadow, z-order, no-fill borders), entrance animations and slide transitions (p:timing), and MP4 export via PowerPoint COM.
---

# 生成演示文稿

生成器只负责把坐标写成 XML。**好看和难看的分界不在生成器，在字阶、栅格和验证**。
下面每一条都是实测踩出来的，不是推测。

## 0. 先把中间产物放对地方

一个 `demo/` 目录 + 每轮预览单独目录（`preview-v1/`、`preview-v2/`），
最后只留：生成器源码、成品 pptx、3–4 张证据图。中间扫描、对照文件全部删掉。

## 1. 选工具

| 你要的 | 用什么 | 代价 |
|---|---|---|
| 内容已经是 HTML / Markdown，先出个草稿 | genppt `FromHTMLFile` | 只有矩形、单字号、单列流式 → **出来像文档** |
| 要有设计感（自由坐标、色块、层次） | genppt `AddText/AddShape` 绝对坐标 API | 得自己算坐标 |
| 要读现成模板 / 母版 | **python-pptx** | genppt 是只写的，读不了 |
| 要渐变 / 阴影 / 透明度 / 图层 | genppt 打补丁（`scripts/genppt-effects.patch`）或 python-pptx + 手写 oxml | 见 `references/ooxml-effects.md` |
| 要动画 | **任何**生成器产物 + `scripts/add_anim.py` 后处理 | 不依赖 Office |
| 要 mp4 | `scripts/pptx2video.ps1`（PowerPoint COM） | 必须装 PowerPoint |

关键判断：**如果只有「HTML 转 pptx」这一条路，不要指望设计感。**
H 流式入口每块只有一种字号、只能画矩形，yPos 递增排布，天然是文档。要设计稿就得用绝对坐标。

**genppt 是 Go 库，不是 pip 包。** 环境里没有 Go、或者没 clone 下来，就别硬上——
用 python-pptx 走**完全相同的五步流程**（设计系统、栅格、字号阶一模一样，
只是写特效时多几行手写 oxml，见 `references/ooxml-effects.md`）。
要用 genppt：

```
git clone https://github.com/CoolBanHub/genppt
# go.mod:
#   require github.com/CoolBanHub/genppt v0.0.0
#   replace github.com/CoolBanHub/genppt => ../genppt
```

## 2. 五步流程

1. **定画布**：10 × 5.625 in（16:9，等价 1920×1080 px）或 13.333 × 7.5 in（宽屏）。
2. **定字阶**：先算 `Display ÷ Body`，**必须 ≥ 3.5**。54/15 = 3.6 才是设计稿，
   30/18 = 1.67 怎么调色都像 Word。这是第一变量，改字号前不用碰别的。
3. **建栅格**：边距 0.55，四栏 `w 2.14 / gap 0.16`，页码 y = H − M − 0.22。
4. **上色**：纸色打底 + 墨色写字 + **一种**高饱和点缀（面积 ≤ 5%）。
   三种以上高饱和色同时出现 = 廉价。
   **不想从零配色就直接取主题**：`assets/themes.json` 里有 14 套从 HTML 模板派生的色卡
   （角色、对比度、字号阶、字体都算好了），见 `references/themes.md`。
   **版心结构别每页都一样**：配色只是第一条轴，13 种结构不同的版式在
   `scripts/layouts.py`（`LAYOUTS` 注册表），用法见 `references/layouts.md`。
   只有一种版式 + 14 套配色 = 14 页「看起来一样，只是颜色不同」。
5. **验证**（见 §4）。跳过这步等于没做。

具体数值（字号阶、色板、封面公式、参考库怎么挖）→ `references/design-system.md`。
中文的字宽、行距换算、避头尾、中英混排 → `references/chinese-typography.md`。

## 3. 常见坑

1. **细线莫名其妙变蓝。** `AddShape` 在 `LineColor` 为空时会注入 `2F5496`（Office 蓝），
   于是每根细线都被涂蓝。必写 `<a:ln><a:noFill/></a:ln>`；用库的话传 `LineStyle: BorderNone`
   并确认该版本不会注入默认值。
2. **整页像文档。** 先量字号对比度，不要先换配色。
3. **文字撑破框。** 中文按 1.0em、ASCII 按 0.55em 估宽度，溢出会折行压住下面的元素。
   用 `scripts/lint_pptx.py` 抓，或渲染出来看。
4. **动画数量翻倍。** `p:timing` 少一层 `<p:par>` 容器，8 个形状会被数成 16 个，且**不报错**。
5. **文字渐变静默失效。** `a:gradFill` 必须排在 `a:latin` **之前**，`append` 到末尾会被忽略。
6. **拿 LibreOffice 验动画。** 它不解析 `p:timing`，也不渲染文字渐变。动画/文字效果只能用 PowerPoint 验。
7. **`p14:morph` 写进去没用。** PowerPoint 会丢弃该节点，`EntryEffect` 也没有对应值。别写。
8. **单位记错。** genppt：`CharSpacing` 是磅，`LineSpacing` 是百分比——
   但 **100% = 1.2 倍字号**，想要 1.5 倍行高要写 `125`。写 `150` 会得到 1.8 倍，
   整块文字松垮。换算见 `references/chinese-typography.md` §2。
9. **`SaveCopyAs` 存成 .ppt。** 第二个参数必须传 `24`（OOXML），传 `1` 会得到二进制 ppt。
10. **python-pptx 的形状自带 `<p:style>`。** 里面 `lnRef` 指向主题色、`effectRef` 带主题阴影，
    只加 `<a:ln><a:noFill/></a:ln>` 仍留着主题描边和阴影的壳 → 把整个 `<p:style>` 删掉。
    （实测：不删的话 LibreOffice 渲染出色块阴影，PowerPoint 里也看得见。）
11. **中文要写 `a:ea`。** python-pptx 的 `font.name` 只设 `a:latin`，结果是英文用你的字体、
    中文掉回默认，一行里两种字形。
12. **形状级 `a:gradFill` 必须排在 `a:ln` 之前**（`spPr` schema 顺序）。
    顺序错了不报错，只是渐变没了——和文字渐变排在 `a:latin` 后面是同一种病。
13. **LibreOffice cli 出图成功也可能 `exit 1`** 并往 stderr 打 `Image batch cleanup failed.`。
    不要用退出码判成败，去查 PNG 文件是否存在、尺寸对不对。
14. **lint 的溢出是启发式**：报告会写「约 N 行、需 X in、框高 Y in」，
    用这三个数字自己判断，不要只看红不红。
15. **pptx 不自带字体。** 文件里只存字体名和字号，目标机器缺这个字体时 PowerPoint **静默替代**，
    字宽一变，算好的断行和栏宽全废。所以：正文用系统一定有的字体，
    display 才用自带字体（`assets/fonts/` + `scripts/install_fonts.ps1`）。
    别指望脚本嵌字体——`EmbedTrueTypeFonts` 是 Word 的属性，**PowerPoint COM 没有这个开关**，
    它只是应用级「保存时嵌入」的 UI 选项。
16. **中文 `.ps1` 必须存 UTF-8 带 BOM。** 不带 BOM 时 Windows PowerShell 按 ANSI 解码，
    中文变乱码并连带语法错误（典型报错是 `[CmdletBinding()]` 报「意外的属性」——
    其实是上一行的注释被啃掉了首字符）。
17. **按主题算出来的坐标不等号。** 72pt 的 display 主题，标题框高度比 54pt 的大 0.27in，
    下面所有元素得跟着下移。样张脚本 `preview_themes.py` 全部由 `display_pt` 推导 y 坐标，
    加主题不用手调；写死坐标的模板换主题就会重叠。
18. **亮强调色当字会消失。** `#E6FF3D` 在纸色 `#ECECE8` 上对比度 **1.06**——不是「不够醒目」，
    是字没了。14 套主题里 5 套的强调色在白纸上是 1.0–1.8。规则：文字走 `accent_text`
    （对比不足自动退回墨色），细线走 `accent_mark`，反色页走 `on_ink`，
    色块上的字用 `tint_pair()` 连块带字一起定。`lint_pptx.py` 的 R8 会按**局部底色**查这一条。
    它是读回 pptx 真实 XML 查的，所以也查得了别人的产物（genppt 官方 demo 就被抓出白字压
    `#ED7D31`/`#70AD47` 只有 2.77/2.71）；表格按单元格算几何和底色。
19. **强调色换到底色上要现算，别复用主题级结论。** `accent_text` 只对纸色验过；表格的斑马行
    比纸色暗一档，grove 的砖红压上去只剩 2.85（R8 报 error）。凡是非纸色底一律走
    `th.accent_readable(bg)`。同理，一套主题干净不代表十四套都干净——验收要跑全量
    （`python tasks.py verify`）。
20. **页码不能手写。** 版式调一次顺序，手写的「07 / 13」就全错位（真实发生过：第 9 页印着 07）。
    用 `num / total` 推，或者在生成时统一注入。
21. **卡片高度跟内容走。** 3.4in 的卡片装两行字，底下空一截，读起来像没做完；
    2.6in 刚好。同理，正文只占半页时把标题锚到底部，空白才像构图而不是缺内容。
22. **同一套版式换主题要能过。** 至少用一套深底亮字（`studio`）和一套高饱和强调（`neo-grid-bold`）
    各渲一遍——纯色底上的版式问题，只有在换色后才暴露。

## 4. 验证（不可跳过）

```powershell
# ① 版面自检（路径按 skill 目录改；相对路径以 pptx 所在目录为基准最省事）
python <skill>/scripts/lint_pptx.py deck.pptx        # 越界/重叠/溢出/字号阶/文字底色对比(R8)/主题蓝

# ② 渲染出来看（输出目录必须不存在，否则 mkdir EEXIST）
node <libreoffice-kit>/cli.js render --input deck.pptx --output-dir preview-v1 --pages 1,2 --dpi 130
#    DSH 环境里的 cli.js：
#    ...\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js
#    页面号会被重排成 page-0001..；拿到的 PNG 再 read_image 逐页看，不要凭坐标脑补。

# ③ 动画回读（必须 PowerPoint）
#    MainSequence.Count 必须等于形状数；EffectType 必须和 presetID 对上
```

判据是「**读回来和写进去一致**」，不是「能打开」。PowerPoint 对非法 XML 静默丢弃，
打开成功什么都说明不了。

COM 脚本、编号对照表、出错后怎么让 PowerPoint 自己生成参照 → `references/verify-with-powerpoint.md`。
XML 片段（渐变/阴影/无边框/时间轴/切换） → `references/ooxml-effects.md`。
中文实测常数（字宽=字号、行距换算、避头尾、中英混排） → `references/chinese-typography.md`。

## 5. 交付前清单

- [ ] `lint_pptx.py` 无 error
- [ ] 中文 run 有 `a:ea`（R6 不报），段落没用 `just`（R7 不报）
- [ ] 文字色在**局部底色**上对比度 ≥ 3.0（R8 不报）；强调色没有直接刷在亮色块上
- [ ] 版心结构不是每页同一种：结构性版式（封面/章节/数字/对照/时间轴/表格）都出现过
- [ ] 中文标题每行不超过「框宽 ÷ 字号」个字（54pt 标题一行 ≤ 11 字）
- [ ] **字体能在目标机器上渲染**：display 用的自带字体装过（`python scripts/install_fonts.py --check`），
      正文只用系统字体；换机器演示前先在自己机器上渲一遍看替代效果
- [ ] 每一页都用 `read_image` 看过，不是抽样
- [ ] 有动画的话：`MainSequence.Count` 等于形状数，时长和设定一致
- [ ] 有切换的话：`EntryEffect` 非 0
- [ ] 中间预览目录已清掉，只留成品 + 证据图

## 6. 脚本

| 文件 | 作用 |
|---|---|
| `scripts/lint_pptx.py` | pptx 版面检查：越界、重叠、文字溢出、字号阶、文字/局部底色对比度(R8)、主题蓝残留 |
| `scripts/pptx_kit.py` | 共享排版层：`Grid`(12 栏)、`textbox`(对齐/锚点/行距/字距/项目符号)、`rect`(删 `<p:style>`)、`mix/contrast` |
| `scripts/layouts.py` | **13 种结构不同的版式**（`LAYOUTS` 注册表）+ `Theme` 包装（`accent_text/accent_mark/on_ink/tint_pair`） |
| `scripts/preview_layouts.py` | 用同一份内容排一份完整 deck：`--theme studio --only kpi-grid,table` |
| `scripts/build_themes.py` | 从 refs 模板派生 `assets/themes.json`（含对比度校验、字体预设） |
| `scripts/preview_themes.py` | 按注册表排主题样张（每页一个大标题 + 字号阶梯 + 色板），用来验收配色 |
| `scripts/install_fonts.py` | 装 `assets/fonts/` 的 OFL 字体（纯 Python：拷文件+注册表+`AddFontResourceW`）；`--check` / `--list` / `--dry-run` |
| `scripts/install_fonts.ps1` | 同上，PowerShell 版；**执行策略是 AllSigned 的机器跑不了 `.ps1`**（`-ExecutionPolicy Bypass` 也覆盖不了组策略），用 .py 那个 |
| `scripts/add_anim.py` | 给任意 pptx 批量加切换 + 错峰入场（`--effect fade --stagger 160`） |
| `scripts/pptx_anim.py` | 上面的库，`Animator(slide).entrance(shape, "fly", delay=200).build()` |
| `scripts/pptx2video.ps1` | PowerPoint COM 导出 mp4（可写打开 + 异步轮询 + finally 清理） |
| `scripts/genppt-effects.patch` | 给上游 genppt 加渐变/阴影/图层/动画的补丁（`git apply`） |

`add_anim.py` 会**替换**页面上已有的时间轴。手工调过的动画不要再过它。

## 7. 文件地图

| 路径 | 内容 |
|---|---|
| `assets/themes.json` | 14 套可用主题 + 14 套标记不可用；角色色、对比度、字号阶、字体降级链 |
| `assets/fonts/` | 6 个 OFL 字体 + 许可证（分发时必须带 `*-OFL.txt`） |
| `references/design-system.md` | 字号阶、栅格、封面公式、怎么从参考库挖设计 |
| `references/themes.md` | 主题注册表的字段、角色纪律、字体规矩、怎么加主题 |
| `references/layouts.md` | 布局这条轴：13 个版式的来源模板/栅格/内容契约、主题给布局的 5 个颜色 API、怎么加布局 |
| `references/chinese-typography.md` | 中文字宽、行距换算表、避头尾、中英混排、实测常数 |
| `references/ooxml-effects.md` | 渐变/阴影/无边框/时间轴/切换的 XML 片段 |
| `references/verify-with-powerpoint.md` | COM 回读脚本、编号对照表、让 PowerPoint 自己生成参照 |
