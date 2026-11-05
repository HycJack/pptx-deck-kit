# 布局：第二条轴

配色不是设计。14 套主题只换纸色、墨色、字体和字号阶——版心结构完全一样，
所以「看起来都一样，只是颜色不同」。布局是**另一条独立的轴**：

```
主题（配色/字体/字号阶） 14 种   ×   布局（版心结构） 13 种   =   182 种页面
```

两轴相乘才是「模板数量」。同内容换布局 = 结构变了；同布局换主题 = 只有颜色和字体变。

- 代码：`scripts/layouts.py`（注册表 `LAYOUTS` + `build_slide()`）、`scripts/preview_layouts.py`（CLI）
- 几何全部由 `pptx_kit.Grid`（12 栏，页边 0.8in）和主题的字号阶推导，**布局里没有写死的坐标常量**
- 画布默认 13.333×7.5（宽屏）。10×5.625 会把多栏挤扁，多栏布局别用 10×5.625

## 1. 13 个布局

| slug | 名称 | 来源（refs 单页模板） | 栅格特征 | 什么时候用 |
|---|---|---|---|---|
| `cover-type` | 封面·纯文字 | `cover.html` | 无栅格，全幅字 + 大留白 | 没有图，靠字撑封面；字要 ≥96pt |
| `cover-split` | 封面·图文对分 | `image-text-split.html` | 左 7 栏字 / 右 5 栏整高图位 | 有主图或想放色块 |
| `section-divider` | 章节页·反色 | `section-divider.html` | 整页铺墨色，180pt 章节号 | 分章；一页只放一个数字 |
| `two-column-text` | 左标题 + 右双栏 | `two-column.html` | 标题 4 栏 / 正文 7 栏再切两栏 | 讲原理、讲背景，正文最多 2×6 行 |
| `three-cards` | 三栏卡片 | `three-column.html` | 3 等分，卡片高 2.6in | 三类并列；每卡 ≤40 字 |
| `kpi-grid` | 四宫格数字 | `kpi-grid.html` | 4 等分，列间竖发丝线 | 4 个指标并列；数字 ≥72pt |
| `stat-highlight` | 单数字主打 | `stat-highlight.html` | 数字压左 8 栏，右侧 3 栏解释 | 一页只讲一个数字（230pt） |
| `big-quote` | 大引言 | `big-quote.html` | 引号 150pt + 引言 40pt | 用户原话、结论引用 |
| `comparison` | 左右对照 | `comparison.html`（`1fr 90px 1fr`） | 左 5 栏 / 中缝 / 右 5 栏 | 前后对比、方案 A/B；每边 ≤4 条 |
| `timeline` | 横向时间轴 | `timeline.html` | 轴线 + 5 等分节点 | 时间线，节点 3–6 个 |
| `process-steps` | 四步流程 | `process-steps.html` | 4 等分 | 方法/步骤，每步 ≤24 字 |
| `table` | 数据表 | `table.html` | 真表格，列宽按内容给 | 5 列 × 5 行以内；数字列右对齐 |
| `closing` | 结尾·反色 | `cta.html` / `thanks.html` | 居中大字 + 联系方式 | 结尾；和章节页同族 |

refs 里另有 `arch-diagram`、`gantt`、`image-gallery`、`image-grid`、
`pros-cons`、`roadmap`、`mindmap`、`diff`、`chart-*`、`code`、`terminal` 等，
没做成布局是因为它们要么依赖真实图片，要么依赖数据源——需要时按第 5 节自己加。

## 2. 内容契约

`build_slide(prs, theme_json, slug, content, num, total)` 里 `content` 是纯 dict，
键由布局决定（`preview_layouts.py` 的 `CONTENT` 是一份完整样例）：

| 布局 | 必填键 |
|---|---|
| `cover-type` | `kicker, title, lead, meta` |
| `cover-split` | `kicker, title, lead` |
| `section-divider` | `kicker, num, title, lead` |
| `two-column-text` | `kicker, title, paras`（2 段） |
| `three-cards` | `kicker, title, cards`（3 个 `(标题, 正文)`） |
| `kpi-grid` | `kicker, title, kpis`（4 个 `(数字, 单位, 说明)`） |
| `stat-highlight` | `kicker, num, unit, desc` |
| `big-quote` | `kicker, quote, by` |
| `comparison` | `kicker, title, left, right`（各 `(小标题, [条目])`） |
| `timeline` | `kicker, title, points`（`(时间, 说明)`） |
| `process-steps` | `kicker, title, steps`（4 个 `(标题, 正文)`） |
| `table` | `kicker, title, cols, widths, rows`，可选 `aligns, note` |
| `closing` | `kicker, title, contact` |

页码不要写进内容：`footer()` 用 `_num / _total` 推出来。手写页码必然错位
（这张 deck 里曾出现「07 / 13」出现在第 9 页）。

## 3. 主题给布局的 5 个 API

布局只准用这 5 个颜色入口，不许直接用 `accent` 画字——这是配色的耦合点：

| 属性 | 用途 | 为什么 |
|---|---|---|
| `accent_text` | 文字（数字、标题） | `accent_role=block` 或与纸色对比 <3.0 时自动退回墨色。14 套里 5 套的强调色在白纸上是 1.0–1.8 对比度，直接当字等于消失 |
| `accent_mark` | 细线、小圆点、顶部色条 | 太亮的强调色（#E6FF3D）在纸色上 1.06，压成细线就看不见了，往墨色混 50% |
| `accent2_mark` | 装饰引号、次要记号 | 同上，`accent2` 多半是浅色（editorial 1.74 / signal 1.14） |
| `on_ink` | 反色页上的强调元素 | 整页铺 `ink` 后，深强调色同样会消失（14 套里 7 套 accent/ink < 3.0） |
| `tint_pair(color)` | 图位/色块 | 返回 `(块色, 字色)`，保证对比度 ≥4.5。深色主题里中灰块上墨色纸色都不够看，得换块的深浅 |

`accent` 本身只用于面积 ≥0.5in 的色块（`stat-highlight` 的强调色块、表格表头、
整页底色）——那时候亮色才是效果，不是缺陷。

斑马纹用 `mix(paper, rule, 0.5)` 预混：DrawingML 没有透明度，
`<a:alpha>` 在 PowerPoint 2016 以前会被忽略，混色最稳。

## 4. 版式硬规则（踩过的）

1. **一页一个主角。** 三栏等宽卡片里没有主角；`stat-highlight` 用 230pt 数字占掉半个版面才成立。
2. **字号阶 ≥3.5 倍。** 36pt 标题配 12pt 说明是 3.0，刚好过线；24pt 配 12pt（2.0）会被
   `lint_pptx.py` 的 R5 报警——数据页的标题也要 36pt，别退到 24pt。
3. **卡片高度按内容定。** 3.4in 高的卡片装 2 行字，底下会空出一截，看着像没做完；
   2.6in 配 2–3 行刚好。
4. **正文只占半页时，标题贴底。** 左标题右正文那种版式，标题顶对齐会让下半页全空；
   标题锚到底部，空白就变成「构图」而不是「内容被砍了」。
5. **反色页要么整页铺色，要么不铺。** 半幅色块 + 反色字最容易出现对比度事故。
6. **列间竖发丝线比分割线好。** `kpi-grid` 用 0.008in 竖线，比灰色竖条安静。

## 5. 加一个布局

```python
# scripts/layouts.py 末尾
def my_layout(slide, th, g, c):
    bg(slide, g, th.paper)                       # 底色永远显式铺，别指望母版
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 8), 0.7, text=c["title"],
            size=th.h1, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    ...                                          # 用 g.x()/g.span()/g.cols_x()，不写死英寸
    footer(slide, th, g, c["_num"], c["_total"])

LAYOUTS["my-layout"] = (my_layout, "中文名", "来源模板.html", "一句话说明")
```

再在 `preview_layouts.py` 的 `CONTENT` 加一份内容，然后：

```bash
python scripts/preview_layouts.py out.pptx --theme editorial-forest --only my-layout
python scripts/lint_pptx.py out.pptx     # 必须 error 0 / warn 0
# 渲染看图，逐页看，不要抽样
```

验收三条：`lint` 干净、渲染图上文字不压形状（R4 只比文本框，不管压色块）、
换到 `studio`（深底亮字）和 `neo-grid-bold`（高饱和强调色）还能看。

## 6. 为什么不是「14 套主题各做一套版式」

那会变成 182 份要分别维护的代码。布局公用一套函数，主题只提供 5 个颜色 API +
字号阶，两边各自升级互不影响：加主题不用动布局，加布局不用动主题。
