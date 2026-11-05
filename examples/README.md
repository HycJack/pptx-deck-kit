# examples

四个成品 pptx 和五张预览图，都是本仓库脚本生成的，不是手改的。
用途是**验收证据**：`lint_pptx.py` 全绿 + 渲染逐页看过，才算这套版式能用。

| 文件 | 页数 | 画布 | 内容 |
|---|---|---|---|
| `layouts-editorial.pptx` | 13 | 13.333×7.5 | 13 种版式各一页，主题 `editorial-forest`（纸与墨） |
| `layouts-studio.pptx` | 13 | 13.333×7.5 | 同一套 13 页，主题 `studio`（黑黄车间，深底亮字） |
| `layouts-neo-grid-bold.pptx` | 13 | 13.333×7.5 | 同一套 13 页，主题 `neo-grid-bold`（酸绿网格，强调色只能当块） |
| `theme-sheet.pptx` | 14 | 10×5.625 | 14 套可用主题各一页样张：大标题 + 字号阶梯 + 色板 + 对比度 |

复现：

```bash
python tasks.py preview --theme editorial-forest --out examples/layouts-editorial.pptx
python tasks.py preview --theme studio           --out examples/layouts-studio.pptx
python tasks.py preview --theme neo-grid-bold    --out examples/layouts-neo-grid-bold.pptx
python tasks.py sheet --out examples/theme-sheet.pptx
python tasks.py lint examples/layouts-editorial.pptx
```

| 预览图 | 是什么 |
|---|---|
| `previews/layouts-editorial-overview.png` | 13 种版式总览（4×4 拼图） |
| `previews/layouts-studio-overview.png` | 同 13 页换深底主题 |
| `previews/layouts-neo-grid-bold-overview.png` | 同 13 页换高饱和主题 |
| `previews/layouts-matrix.png` | **4 个版式 × 3 套主题**的对照矩阵：结构逐像素一致，配色完全独立 |
| `previews/theme-sheet-overview.png` | 14 套主题总览（配色这条轴） |

`layouts-matrix.png` 是读懂这个仓库最快的图：横着看是同一套配色下的不同结构，
竖着看是同一结构下的不同配色。两条轴互不影响，才有 14 × 13 = 182 种页面。

示例内容是一份虚构的「把演示文稿做成设计稿」复盘，13 页里刻意覆盖了所有结构陷阱：
一页只讲一个数字、三栏但没有主角、表格 5 列、时间轴 5 个节点、反色章节页。
换成你自己的内容时，改 `skill/scripts/preview_layouts.py` 的 `CONTENT` 即可。
