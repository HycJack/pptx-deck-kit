# 设计系统：把 HTML 排版搬进 10×5.625 英寸

生成器只负责把坐标写成 XML，**好不好看全在这几个数字上**。这套数字是从两个 HTML 模板库
（`lewislulu/html-ppt-skill`、`zarazhangrui/beautiful-html-templates`）里量出来再换算的。

## 1. 画布与单位换算

| HTML 参考 | PPT 成品 |
|---|---|
| 1920 × 1080 px | 10 × 5.625 in = 720 × 405 pt |
| 96 px 边距 | 36 pt = 0.5 in（实际取 0.55 更稳） |
| 72 px 上下留白 | 27 pt = 0.375 in |

**1 px = 0.375 pt**。看到模板里的 px 值，乘 0.375 就是 pt。

**画布怎么选**：模板是 16:9 就用 `10 × 5.625`（本文所有栅格数字都只对这一个成立）；
要 4:3 用 `10 × 7.5`；要电影感宽幅用 `13.333 × 7.5`（此时边距仍取 0.55，
但四栏宽度要重算：内容宽 12.23 in，`w = (12.23 - 3×0.16) / 4 = 2.94`）。

## 1.5 字体

字号阶只是骨架，字族决定气质。Windows 上确定能用的：

| 用途 | 字族 |
|---|---|
| 标题（细、清爽） | `DengXian Light` |
| 标题（常规） | `微软雅黑` |
| 正文（黑体） | `SimHei` |
| 正文（宋体） | `SimSun` |
| 代码 / 等宽 | `Consolas`、`Cascadia Mono` |

**中文必须同时设 `a:latin` 和 `a:ea`。** python-pptx 的 `font.name` 只写 `a:latin`，
结果是中文掉回默认字体、英文却是你要的那个，一行里两种字形。所以要同时写 `a:latin` 和 `a:ea`。

字距：`letter-spacing: 0.12em` 在 8.5pt 上等于 **1.02 pt**——
OOXML 里写 `rPr/@spc="130"`（单位 1/100 磅，即 1.3 磅）。

## 2. 字号阶

| 角色 | 模板里的 px | 成品 pt |
|---|---|---|
| Display | 220 / 140 / 96 | 54 |
| H1 | 68 / 56 | 36 |
| Lead | 44 | 24 |
| H2 | 32 | 19 |
| Body | 26 | 15 |
| Small | 24 | 12 |
| Label | 24 | 8.5 + 字距 |

**唯一最重要的数字：Display ÷ Body ≥ 3.5。**

我第一版用 30/18 = 1.67，出来像 Word 文档；改成 54/15 = 3.6 之后立刻像设计稿。
不是配色问题，不是留白问题，是**对比度**问题。改字号前先算这个比值。

## 3. 字距与行距

- Display 的 `letter-spacing: -0.02em`：小字号设负值会糊，54pt 以上可以不设。
- Label 的 `+0.12 ~ 0.18em`：8.5pt 字配 `spc="130"`（1.3 磅）。
- 行高：**`a:spcPct/@val` 是 1/1000 百分比**，`100000` = 100%、`95000` = 95%；
  但 100% 等于 **1.2 倍字号**（不是 1.0 倍）。想要 1.5 倍行高要写 `125`。
  完整换算表见 `chinese-typography.md` §2。
  手写 OOXML 时对应 `a:lnSpc/a:spcPct val="95000"`（千分之一百分比）。
  Display 要压紧到 0.95× 字号 → `LineSpacing: 79`；Body 要 1.5× → `LineSpacing: 125`。
- 用**英寸**直接摆多行文本时，行距 = `字号 × 行高倍数 / 72`：
  15pt × 1.5 = **0.31 in**。这个 0.31 不是另一种行高规范，它就是 1.5 倍换算来的。

## 4. 配色

```
paper   F6F2E9   纸色，打底
paper2  EFE8D8   次级底，用于整页区分
ink     1A1A17   正文墨色（纯黑 000000 太硬）
forest  2E4A2A   主色
forestDk 1B2E18  深色页底
pink    E89CB1   点缀 1
pinkPale F7E3E9  点缀 1 的浅版
rust    A6432A   点缀 2
muted   6F6B60   注释、页脚
rule    D8D0BE   分隔线
```

规则：**高饱和色只做点缀，总面积不超过 5%**。一片纸上只有一条粉带、一个深色面板，
其余全是纸色和墨色——这就是「杂志感」的来源。三种以上高饱和色同时出现 = 廉价。

## 5. 栅格

```
左右边距 M = 0.55          → 内容宽 8.9 in
四栏卡: w 2.14 / gap 0.16  → x = 0.55, 2.85, 5.15, 7.45
页眉 label:  y = 0.62
主标题:      y = 0.92
正文起点:    y = 1.7 ~ 1.75
页脚页码:    y = H - M - 0.22 = 4.855
```

分隔线用 `h = 0.008` 的极细矩形，**不要用 `LineStyle: BorderNone` 的空心矩形凑**（见 SKILL.md 的坑 1）。

## 6. 一个可以直接抄的封面公式

```
0 ~ 0.10 in      竖向色脊（粉）
6.55 ~ 10 in     深色面板满高（forest 渐变），带阴影
0.55, 0.72       label 8.5pt + 字距
0.55, 1.42       第一行 display 46–54pt
0.55, 2.12       第二行 display，换主色
0.55, 3.22       hairline
0.55, 3.40 起    三行 body 15pt，行距 0.32 **英寸**（= 1.5 倍行高，见 §3），逐行 y 递增
```

## 7. 参考库怎么用

| 库 | 有什么 | 怎么用 |
|---|---|---|
| [html-ppt-skill](https://github.com/lewislulu/html-ppt-skill) | 36 个主题 CSS、24 个 JS 特效 | 抄 CSS 变量里的色板和字号 |
| [beautiful-html-templates](https://github.com/zarazhangrui/beautiful-html-templates) | 34 套，每套 `template.html` + `design.md` + `template.json` | 读 `design.md` 拿 tokens 和构图说明 |

**不要试图把这些 HTML 直接喂给生成器。** HTML→pptx 的解析器普遍只认 inline-style 白名单、
单列 yPos 流式布局、每框一个字号，不吃 CSS grid、渐变、绝对定位。它们是**设计资料**，不是输入。

正确的用法：挖 tokens（字号阶、色板、字距、行高）→ 换算成 pt → 用自己的绝对坐标 API 重画。
