# 要不要把 genppt 改成 python-pptx 那样的库

> 背景：本仓库曾经内嵌过 [genppt](https://github.com/CoolBanHub/genppt)（MIT，Go，写了 5,029 行）
> 并用一份 6 文件的补丁给它补上渐变/阴影/图层/动画。2026-11 起改成纯 Python，`vendor/genppt/`
> 已移除。这份文档留下当时的结论：**genppt 能不能照着 python-pptx 的模型重做，值不值得。**

## 1. 结论

能，但那不是"给 genppt 加个 read 函数"的量级，而是**按 OOXML schema 重写一层**。

- python-pptx 的 27,830 行里，真正干活的不是 API 糖，是 `pptx/oxml/`（31 个文件）那层
  「XML 元素 ⇄ 类型化对象」的映射，加上 zip/rels/content-types 的包读取。
- genppt 是 5,029 行**字符串拼 XML**的写入器：没有读、没有对象图、没有 schema 层。
- 要"一样的库模式"，缺的正是那两层，而它们占了 python-pptx 的绝大部分工作量。

所以：**现在不做**。真有 Go 侧需求时的三条路在 §7。

## 2. python-pptx 的模型，五个设计决定

拆开看，它的"库感"来自五件互不相关的事：

| # | 设计决定 | 具体表现 | 在本仓库里的痕迹 |
|---|---|---|---|
| 1 | **对象图**（Presentation → Slides → Shapes → TextFrame → Paragraphs → Runs） | 一切都可以遍历、可以改，不是一次性 options struct | `layouts.py` 里 `slide.shapes.add_textbox(...)` 之后接着拿 `tf.paragraphs[0].runs` |
| 2 | **oxml 代理层**：每个 wrapper 持有一个 lxml 元素，改属性即改 XML | `shape._element`、`qn("a:ea")`、`find/insert` | `pptx_kit.py` 用 `_element` 插 `<a:ea>`；`lint_pptx.py` 用 `sh._element.xml` 读回 |
| 3 | **读写同一棵树** | `Presentation(path)` 打开已有文件 → 就地改 → `save()` | 这就是 linter 能"读回真实 XML"的前提 |
| 4 | **模板保真**：未知 part 原样带过 | 打开一个设计好的 pptx 当模板，只替换文字 | 我们没用，但这是"用 PowerPoint 设计、Python 填内容"的路 |
| 5 | **无渲染** | 只管 XML，出图是别人（LO/Office）的事 | 见 `tasks.py render` 的三条路 |

第 2 条是关键：**OOXML 的元素顺序是 schema 强制的**（`a:latin` 必须在 `a:ea` 前、
`a:gradFill` 必须在 `a:ln` 前），顺序错了 PowerPoint 静默丢弃，不报错。python-pptx 用一个
`insert_element_before` 的机制把顺序知识编码进属性 setter 里；这就是它 27k 行里最值钱的代码。

## 3. genppt 的模型，三个事实

1. **纯写入**：`writer.go` 直接 `zip.NewWriter` + 逐 part 写字符串；
   `xml_slide.go` 是 `sb.WriteString("<a:pPr")` 这种拼接。全仓库没有 pptx 读取器
   （`html.go` / `markdown.go` 里的 `parse` 解析的是 HTML / Markdown，不是 pptx）。
2. **平面 options struct**：`slide.AddText("标题", TextOptions{X:1, Y:2, Width:8,
   FontSize:44, Align:AlignCenter})`——一次调用写到底，写完拿不回来改。
3. **API 面很窄但够写**：`AddText / AddShape / AddTable / AddImage / SetBackground /
   SetNotes / SetTransition / AnimateLast / SetLayout / SendToBack`。
   我们打的那份补丁（渐变/阴影/透明度/图层/动画）就是在这个模型上"多加几个 options 字段"。

它写得对、也够用；缺的是"读"和"可改"。

## 4. 差距清单

要做到「`Presentation(path)` → 遍历 → 改 → `save()`」，需要的模块：

| 模块 | genppt 现状 | 工作量锚点 |
|---|---|---|
| zip/rels/`[Content_Types].xml` 包读取 | 无（只会写） | 小，`archive/zip` + `encoding/xml` 一两周 |
| OOXML schema 类型（`CT_*` 结构体） | 无 | **大**：PowerPoint 主 schema 数千个类型；unioffice 的做法是按 XSD 生成，二进制因此 33MB |
| 顺序安全的序列化器 | 无（字符串拼接，顺序靠手写对） | **大**：Go 的 `encoding/xml` 按字段声明顺序输出，可以硬编码，但要为每个 `CT_*` 维护字段顺序 + 可选元素 |
| wrapper 对象 + `X()` 逃逸口 | 无 | 中：每个 wrapper 就是一薄层 getter/setter |
| 枚举 / 单位（EMU、`1/100 pt`、`1/1000 %`） | 部分 | 小 |
| 模板保真（未知 part 透传） | 无 | 中：读到不认识的 part 也要原样写回 |
| 渲染 | 无 | **不可能**：这条路 python-pptx 也没走 |

一句话：**写入器是一条线，读取+对象图是一个面**。genppt 现在的 5,029 行大概等于 python-pptx 的
API 糖部分；要补齐到"同一个模型"，代码量会往 python-pptx 的 27,830 行看齐。

## 5. 如果真做，分四步（每步都有可验证的交付）

1. **P0 拆出 oxml 层**：把现在拼字符串的地方换成 `CT_*` 结构体 + 一个按 schema 顺序输出的
   serialiser。判据：现有 13 版式 deck 用新层重新生成，XML 与旧版**逐字节等价**（顺序、空元素、
   命名空间前缀都不许变）。
2. **P1 加包读取**：`Open(path)` 返回 part 表 + rels 图；能列出每个 slide 的 part 名。
   判据：读 python-pptx 生成的 4 份 deck，slide 数、part 数、每个 part 的 sha256 全部对得上。
3. **P2 加 proxy 对象**：`Shapes()/TextFrame()/Paragraphs()/Runs()` 直接映射到树上的元素。
   判据（最关键）：**round-trip 测试**——用 python-pptx 生成 deck → genppt 打开 → 只改一个
   文本框的字 → 存回 → 用 `lint_pptx.py` 和新旧 XML diff 检查：除那一个字，其他一个字节都没变。
4. **P3 模板保真**：把未知 part 原样透传。判据：拿一份带 SmartArt 的 pptx，读进来存回去，
   PowerPoint 打开仍显示 SmartArt。

P2 的 round-trip 是分水岭：做不到"只改该改的"，就没有"库"，只有"生成器"。

## 6. 先例

- **[unioffice](https://github.com/unidoc/unioffice)（原 gooxml）**：Go 里唯一同级别的 OOXML 库，
  做法就是 python-pptx 那套——按 XSD 生成 `schema/` 结构体，外面套 wrapper，用 `X()` 逃逸到原生类型。
  但要注意两点（截至本次核对）：
  - 许可证是 **UniDoc EULA，商用要 license key**（免费额度需注册），不是 MIT/Apache；
  - 它的 **pptx 支持明显比 python-pptx 薄**：README 里 pptx 只列了「从模板创建 / 文本框与形状」，
    没有表格、图表、动画。也就是说：连"做对了模型"的那家，pptx 也只做到这个深度。
- 这也反过来说明 `pptx` 这个格式的重心不在"写"，在"改"和"保真"——而那正是我们不需要的部分：
  我们是从零排版，不需要读别人的文件。

## 7. 建议

| 你的情形 | 建议 |
|---|---|
| 就在本机/DSH 里生成 deck、要现场改版式 | **保持纯 Python**（现状）。改一行 Python 立刻重排，不需要工具链 |
| 要把生成逻辑放进 Go 服务（无 Python 运行时） | 先评估 **unioffice**（license 成本 vs 自研成本）；它的 pptx 只覆盖文本框/形状，够不够要按你的版式算 |
| 必须自研且有 3–4 周预算 | 按 §5 的 P0→P2 走，**别碰 P3/P4**。Lint/渲染/主题层继续留在 Python 侧，Go 只当写入器 |

无论哪条，都别指望 Go 侧替代 linter：审计必须由"不是写它的那个程序"来做
（见 README「验收」一节，以及 genppt 官方 demo 被我们 linter 抓出 2.77/2.71 对比度那次）。

## 8. 这次移除掉了什么

- `vendor/genppt/`（上游 61df963 原始树 + 补丁，41 个文件）
- `skill/scripts/genppt-effects.patch`
- `tasks.py genppt`（check/apply/revert）
- CI 的 `go writer` job 与补丁一致性守卫
- 文档里 genppt 专有的坑（单位记法、`SaveCopyAs` 参数）——**结论保留**，只是不再绑在某个 API 上

**保留的**：`skill/references/ooxml-effects.md`（渐变/阴影/切换/动画的 XML 片段与验证顺序，
这些是格式知识，跟用什么写无关）、`skill/scripts/pptx_anim.py`（`p:timing` 生成）、
以及「写错不报错、必须回读验证」这条纪律。
