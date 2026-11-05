# 主题注册表（themes.json）

`assets/themes.json` 是从 `refs/beautiful-html-templates` 的 34 套 HTML 模板**派生的机器可读色卡**。
脚本 `scripts/build_themes.py` 负责生成，样张 `scripts/preview_themes.py` 负责验收。

**为什么不是手写 markdown 色卡**：HTML 模板的 palette 键名是自由的（`neon_pink`、`bg_alt`、`wood`、
`acid`…），字号是 px，字体是 Google Fonts，还有 `mix-blend-mode` / 纹理 / 1px 网格这些 pptx 做不到
的东西。直接用会掉进三个坑：颜色对不上角色、字号换算错、字体在目标机器上根本不存在。
JSON 的价值就在这里——它把「上游的自由设计」翻译成「pptx 的固定槽位」，并且翻译规则是显式的。

**主题只是一条轴。** 换配色不会改版心结构——14 套主题排同一种版式，出来的就是「14 页看起来一样、只是颜色不同」。版式结构见 `layouts.md`（13 种），两者相乘才是模板数量。

## 1. 六条适配规则（`provenance.adaptation`）

| 上游 | pptx | 为什么必须改 |
|---|---|---|
| 自由颜色键名 | 固定角色 `paper/ink/accent/accent2/rule` | pptx 没有「语义色」；代码里只有 RGB。不定角色就会一页五个彩色 |
| px 字号 | `display_pt` 54–72 / `h1` 36 / `lead` 24 / `body` 15 / `small` 12 / `label` 8.5 | 1920px 宽的 HTML 换成 10in 版面，px 直接当 pt 用会大一倍 |
| web font | 本机/自带字体 + `fonts_fallback` 降级链 | pptx 只存字体名，不存字形；字体缺失时 PowerPoint 静默替代，字宽变了版面就崩 |
| `lineHeight: 0.92`（em 倍数） | `a:spcPct` 千分比：`em / 1.2 × 100000` | **实测 100% = 1.2 × 字号**，不是 1.0。0.92em 要写 76667，写 92000 就变 1.10 倍 |
| `letterSpacing: -0.02em` | 磅值（`spc = pt × 100`）；**中文强制 0** | 方块字加字距会散架；`tracking_display_cjk_pt` 单独记 0 |
| 凭感觉用强调色 | 按 WCAG 对比度定 `accent_role: text / block` | 亮黄绿当字只有 1.0–1.1 对比度，等于看不见，只能当色块 |

`rule`（发丝线色）和部分 `accent2` 上游没有，标在 `derived_roles` 里——**允许派生，但必须标出来**，
其余每个 hex 都由脚本校验必须真的出现在上游 `template.json` 的 palette 中，防止色值被手抄篡改。

## 2. 14 个可用主题

对比度列是 `ink/paper` 与 `accent/paper`，后者 < 3.0 就不能当字。

| slug | 中文名 | register | paper | ink | accent | 对比度 | accent_role | 字体预设 |
|---|---|---|---|---|---|---|---|---|
| `editorial-forest` | 纸与墨 | editorial | `#efe7d4` | `#1a1a17` | `#2e4a2a` | 14.2 / 8.0 | text | serif-editorial |
| `cobalt-grid` | 电光蓝 | technical | `#F0EBDE` | `#1F2BE0` | `#5560E5` | 7.1 / 4.2 | text | grotesque |
| `signal` | 夜航金 | formal-dark | `#1c2644` | `#e2dcd0` | `#c8a870` | 10.9 / 6.6 | text | serif-editorial |
| `neo-grid-bold` | 酸绿网格 | technical | `#ECECE8` | `#0A0A0A` | `#E6FF3D` | 16.7 / 1.1 | **block** | condensed |
| `bold-poster` | 黑红宣言 | poster | `#F5F2EF` | `#1C1410` | `#D8000F` | 16.3 / 4.8 | text | poster-condensed |
| `monochrome` | 素纸 | minimal | `#fafadf` | `#1a1a16` | `#5e5e54` | 16.5 / 6.2 | text | serif-warm |
| `biennale-yellow` | 双年展黄 | cultural | `#E9E5DB` | `#1B2566` | `#F1EE2E` | 11.1 / 1.0 | **block** | serif-warm |
| `mat` | 苔绿琥珀 | editorial-dark | `#232e26` | `#f0e8d2` | `#c07030` | 11.5 / 3.8 | text | grotesque |
| `long-table` | 长桌赭红 | cultural | `#FAF1E2` | `#8E2D1F` | `#B53D2A` | 7.4 / 5.1 | text | condensed |
| `soft-editorial` | 柔和三色 | editorial | `#F2EEDF` | `#2A241B` | `#E1A4C2` | 13.2 / 1.8 | **block** | serif-warm |
| `emerald-editorial` | 祖母绿 | editorial | `#F1E9D6` | `#0F1A5C` | `#3CD896` | 13.1 / 1.5 | **block** | serif-editorial |
| `studio` | 黑黄车间 | poster | `#1c1c1c` | `#f5d200` | `#f5d200` | 11.4 / 11.4 | text | poster-condensed |
| `grove` | 深林砖红 | editorial-dark | `#192b1b` | `#d4cfbf` | `#c8524a` | 9.6 / 3.4 | text | serif-editorial |
| `sakura-chroma` | 木色五色 | cultural | `#F1E6CB` | `#3A2516` | `#E5392A` | 11.6 / 3.4 | **block** | condensed |

选主题看 register：`editorial` 稳、`technical` 密、`poster` 狠、`cultural` 有情绪、
`*-dark` 用于年报/深度叙事、`minimal` 用于纯文字稿。

看效果：`python scripts/preview_themes.py out.pptx` 会按注册表排 14 页样张，
每页含大标题、正文、36/24/15/12 字号阶梯、五角色色板。**样张过 lint 才说明注册表的数值能用。**

## 3. 角色纪律

| 角色 | 独占用途 | 禁止 |
|---|---|---|
| `paper` | 满版底色 | 不要给正文框单独填 paper——会露出边界 |
| `ink` | 所有正文、标题、页脚 | 不要用纯黑 `#000000`：印刷/投影都会发死，上游给的都是带一点点色的近黑 |
| `accent` | 一处：要么一笔字，要么一块色 | 一页最多用一次。`accent_role=text` 时它替代 ink 做标题；`block` 时只出块，块上的字用 `accent_text_color` |
| `accent2` | 次要元素、分隔、次强调 | 不要和 accent 等量出现，否则没有主次 |
| `rule` | 0.75pt 发丝线、表格线、边框 | 不要大于 1pt；不要用来画装饰条（那是 accent 的活） |

`accent_text_color` 是脚本算出来的：块上文字取 `ink` 或 `#FFFFFF` 里对比度更高的那个。
`sakura-chroma` 的 `accent_block_because` 说明了为什么它 3.4 的对比度仍然只做块：
五色块是主题本体，红只当块。

## 4. 字体

`assets/fonts/` 里带 5 个 OFL 字体（`*-OFL.txt` 是许可证，**分发时必须一起带**）：

| 族 | 用途 | 文件 |
|---|---|---|
| Instrument Serif | 高对比衬线 display（只有 Regular/Italic，**没有粗体**） | `InstrumentSerif-*.ttf` |
| Bebas Neue | 全大写窄体 poster display | `BebasNeue-Regular.ttf` |
| Barlow / Barlow Condensed | 无衬线 body / display | `Barlow*.ttf` |
| IBM Plex Mono | label、色值、页码 | `IBMPlexMono-Regular.ttf` |

安装：`scripts\install_fonts.ps1`（管理员装全机；普通权限装本用户，需重启 PowerPoint 或重新登录）。
检查：`install_fonts.ps1 -Check`；看主题用到哪些族：`install_fonts.ps1 -List`。

**中文一律用系统字体**，不打包分发（微软雅黑/宋体/等线的授权不允许随文件分发）：
宋体系（`华文宋体`/`宋体`）配衬线主题，黑体系（`微软雅黑`/`等线 Light`）配无衬线主题。
macOS 上会掉到 `PingFang SC`，字宽接近，版面不会崩，但要在那边重新渲一遍看断行。

每个主题都有 `fonts`（目标）和 `fonts_fallback`（缺字体时）：

| 目标 | 降级 |
|---|---|
| Instrument Serif | Georgia |
| Bebas Neue / Barlow Condensed | Arial Narrow |
| Barlow | Segoe UI |
| IBM Plex Mono | Consolas |

两条硬规矩：

1. **高对比衬线只做 display。** Instrument Serif 没有粗体、笔画极细，15pt 正文会糊；
   正文回退到 Georgia/Constantia 这类有粗体的系统衬线。
2. **排版要在降级字体上也能看。** 目标机器没装 OFL 字体时 PowerPoint 会静默替代，
   所以 display 大标题可以依赖自带字体（短、替代后最多变宽），正文和表格一律用系统字体。

## 5. 加一个主题

1. 在 `refs/.../templates/<slug>/template.json` 找到上游模板，决定四个角色色 + `rule`。
2. 在 `build_themes.py` 的 `C` 列表加一条：`roles`（`rule` 等派生色写进 `derived`）、
   `accent_role`、`font` 预设、`type`（`display_pt` 与 `body_em`）、`note`（一句话说明适用场合）。
3. 跑 `python scripts/build_themes.py`。它会拦住五类错误：
   - hex 不在上游 palette 里
   - `accent_role=text` 但对比度 < 3.0
   - `accent_role=block` 但对比度已够当字（需显式写 `accent_block_because`）
   - `display_pt / 15 < 3.5`（标题不够大，读起来像公文）
   - 上游没有这个 slug
4. 跑 `python scripts/preview_themes.py --only <slug>` 看单页，再 `lint_pptx.py` 过一遍。

## 6. 不能用的 14 套（`ppt_safe=false`）

标记在 JSON 里，带 `ppt_limit` 原因：`block-frame`（叠印+负间距）、`8-bit-orbit`（像素字+发光）、
`retro-windows`（窗口控件）、`retro-zine`（网点纹理）、`scatterbrain`（随机旋转）、`cartesian`
（极细网格）、`capsule`（手绘 SVG 胶囊）、`pin-and-paper`（手写体+图钉位图）、`creative-mode`
（多色叠印）、`peoples-platform`（粗描边字）、`raw-grid`（1px 全铺网格）、`stencil-tablet`
（喷漆断字）、`daisy-days`（贴纸装饰）、`blue-professional`（`#1E2BFA` 与 Office 默认蓝太接近）。

要用它们的**配色**可以，别照抄版式——把颜色填进一个可用主题的角色槽，或者直接用色块近似纹理。
