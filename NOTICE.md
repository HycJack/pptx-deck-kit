# 第三方组件与来源

本仓库自己的代码是 MIT（见 [LICENSE](LICENSE)）。下面这些不是我的代码，
按各自许可证收录，一并说明来源与改写范围。

> 曾经内嵌过 Go 的 genppt 写入器（上游 `61df963` + 一份 6 文件补丁），2026-11 起改成纯 Python，
> 已整体移除。为什么不做成 python-pptx 那样的 Go 库，见 [docs/go-pptx-library.md](docs/go-pptx-library.md)。

## 1. beautiful-html-templates — 主题注册表的上游

- 上游：<https://github.com/zarazhangrui/beautiful-html-templates>
- 许可证：**MIT**，Copyright (c) 2026 Zara Zhang
- 派生基线：`e5e204fb1f3b06290846e7dcd7aceddabeceec8c`
- 派生范围：`skill/assets/themes.json` 里的**色值、字号比例、字体搭配**是从这套模板的 CSS 变量
  里提取并重新计算的（对比度、派生角色 `accent_text`/`accent_mark`/`on_ink` 都是这边新增的算法）。
  模板本身的 HTML/CSS/JS **没有**收进本仓库。
- 重新生成：`THEMES_REFS=/path/to/beautiful-html-templates/templates python skill/scripts/build_themes.py`
  （平时不需要跑，注册表已提交）

## 2. html-ppt-skill — 版式原型的参考

- 上游：<https://github.com/lewislulu/html-ppt-skill>
- 许可证：**MIT**，Copyright (c) 2026 lewis
- 基线：`fd1629067909ff55b36b905476de4e5f26062a1f`
- 参考范围：13 个版式的**栅格比例与信息层次**（例如 `1fr 90px 1fr` 的三段式对照、`repeat(4,1fr)`
  的四宫格、章节页的 180pt 数字）参照了这套模板的写法，落地时全部重算成 pptx 的英寸坐标，
  并按中文字宽（1 字 = 1 em）重新定了字号阶。没有复制其代码或文案。

## 3. 字体 — `skill/assets/fonts/`

6 个字体文件，全部 **SIL Open Font License 1.1**，许可证原文随字体一起分发（`*-OFL.txt`）：

| 字体 | 用途 |
|---|---|
| Instrument Serif（Regular / Italic） | 高对比 serif display |
| Bebas Neue | 压缩大写 display |
| Barlow / Barlow Condensed | 拉丁正文与标签 |
| IBM Plex Mono | 等宽：页码、kicker、元信息 |

分发要求：OFL 允许随产品打包与再分发，但**必须**带着 `*-OFL.txt`，且不得单独售卖字体本身。
中文不打包，走系统字体（`华文宋体` / `微软雅黑` / `等线 Light` / `Songti SC` / `PingFang SC`）。

## 4. LibreOffice（仅开发期）

**可选**：只在把 pptx 渲成 PNG 做肉眼验收时才用到，不随本仓库分发，也不是依赖。
`tasks.py render` 有三条路：libreoffice-kit 的 cli.js（自带 soffice，机器上不用装 LO）/
PowerPoint COM（Windows + Office，`--engine com`）/ 裸 `soffice`（只出第一页）。
三者都没有时，`lint`、`preview`、`verify` 照常工作。
