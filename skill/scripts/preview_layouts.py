#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用同一套内容，把 13 种布局排成一份完整 deck。同一个 --theme 换配色即可对比。

用法:
    python preview_layouts.py deck.pptx --theme editorial-forest
    python preview_layouts.py deck.pptx --theme studio --only cover-type,kpi-grid
    python preview_layouts.py a.pptx --theme signal --list   # 列出布局

设计意图：布局和配色是两条独立的轴。
同一份内容换布局 = 结构变了；同一套布局换主题 = 只有颜色和字体变。
两者相乘才是「模板数量」，不是 14 套主题各自硬编码一套版式。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from layouts import LAYOUTS, build_slide          # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
REGISTRY = os.path.join(os.path.dirname(HERE), "assets", "themes.json")

CANVAS = (13.333, 7.5)      # 宽屏：布局需要横向空间；10×5.625 会把多栏挤扁


# ---------------------------------------------------------------- 内容
CONTENT = {
    "cover-type": dict(
        kicker="2025 · 年度复盘 · 内部",
        title="把演示文稿\n做成设计稿",
        lead="同一个内容，13 种版式；同一套版式，14 套配色。",
        meta="产品设计组\n2025.11.20\n仅内部传阅",
    ),
    "cover-split": dict(
        kicker="案例 01 / C 端产品",
        title="移动端\n重设计",
        lead="12 周，4 个版本，NPS 从 31 提到 49。",
        meta="",
    ),
    "section-divider": dict(
        num="02", title="内容决定版式",
        lead="不是先挑模板，是先看这一页要干什么。",
        kicker="第二章",
    ),
    "two-column-text": dict(
        kicker="02.1 / 原理",
        title="先定字阶，\n再谈布局",
        paras=[
            "字号是唯一的层级手段。54pt 标题配 15pt 正文，比值 3.6，"
            "眼睛一眼就知道从哪读起。这个比值低于 3.0，无论怎么调色都像 Word 文档。",
            "布局只是在字阶之上分配空间：一页只允许一个视觉主角，其余全部降一级。"
            "三栏卡片里三块等宽，就没有主角；把其中一块放大到 2 倍，主角立刻出现。",
        ],
    ),
    "three-cards": dict(
        kicker="02.2 / 页型",
        title="三种页型覆盖 90% 场景",
        cards=[
            ("陈述", "标题 + 一段话。用于背景、动机、结论。字少，留白多，最容易被低估。"),
            ("举证", "数字、表格、对照。用于说服。数字必须大到自己站得住，旁边只放单位。"),
            ("推进", "时间轴、四步流程。用于讲过程。节点等距，说明不超过两行。"),
        ],
    ),
    "kpi-grid": dict(
        kicker="03.1 / 结果",
        title="上线后的四个变化",
        kpis=[
            ("3.6", "× 首页停留", "从 42 秒到 152 秒，主要来自首屏信息密度下降。"),
            ("49", "NPS", "比上一版 +18，负面反馈集中在设置页。"),
            ("−38", "% 跳出率", "首屏加载从 2.4s 降到 1.1s。"),
            ("12", "周交付", "含 4 轮可用性测试，每轮 8 人。"),
        ],
    ),
    "stat-highlight": dict(
        kicker="03.2 / 单一指标",
        num="3.6",
        unit="× 首页停留时长",
        desc="所有改动的收益集中在一个指标上。\n一页只讲一个数字时，数字要占到半个版面，其余全部让位。",
    ),
    "big-quote": dict(
        kicker="04 / 用户原话",
        quote="以前要先读三行字才知道这页在说什么，\n现在第一眼看到的就是结论。",
        by="— 受访者 07，第 3 轮可用性测试",
    ),
    "comparison": dict(
        kicker="05.1 / 前后对照",
        title="旧版 vs 新版：同一份内容",
        left=("旧版", ["标题 30pt，正文 18pt，比值 1.7",
                       "四栏等宽塞满，没有主角",
                       "三种高饱和色同时出现",
                       "关键数字藏在第三段"]),
        right=("新版", ["标题 54pt，正文 15pt，比值 3.6",
                        "一页一个主角，其余降级",
                        "纸色 + 墨色 + 一处强调",
                        "数字独占一页，230pt"]),
    ),
    "timeline": dict(
        kicker="05.2 / 过程",
        title="12 周里发生了什么",
        points=[("W1", "审计旧版\n17 个页面"),
                ("W3", "重定字阶\n与栅格"),
                ("W6", "第一版\n可用性测试"),
                ("W9", "两轮迭代\n收敛版式"),
                ("W12", "上线\nNPS 49")],
    ),
    "process-steps": dict(
        kicker="05.3 / 方法",
        title="我现在的四步",
        steps=[("定画布", "10×5.625 还是 13.333×7.5，先定，别中途改。"),
               ("定字阶", "Display ÷ Body ≥ 3.5，这一条不过就别往下走。"),
               ("选页型", "陈述 / 举证 / 推进，一页只挑一种。"),
               ("验证", "lint 过一遍，再渲染逐页看，不抽样。")],
    ),
    "table": dict(
        kicker="06 / 数据",
        title="三种页型的成本对比",
        cols=["页型", "单页耗时", "返工率", "适用", "评分"],
        widths=[3.4, 2.3, 2.2, 2.4, 1.43],
        aligns=["left", "right", "right", "left", "left"],
        note="耗时与返工率取 12 周内的实际记录；评分 A/B/C 为设计组内部口径，不是行业标准。",
        rows=[["陈述", "18 min", "12%", "背景/结论", "A"],
              ["举证", "45 min", "31%", "数字/对照", "A"],
              ["推进", "38 min", "24%", "过程/时间", "B+"],
              ["图墙", "62 min", "44%", "案例/作品", "B"],
              ["纯文字", "9 min", "6%", "讲稿/备忘", "C"]],
    ),
    "closing": dict(
        kicker="产品设计组 · 2025",
        title="谢谢",
        contact="deck@example.com  ·  +86 138 0000 0000",
    ),
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default="layout-deck.pptx")
    ap.add_argument("--theme", default="editorial-forest",
                    help="themes.json 里的 slug")
    ap.add_argument("--only", default="", help="只排这些布局，逗号分隔")
    ap.add_argument("--list", action="store_true", help="列出布局后退出")
    ap.add_argument("--size", default="13.333x7.5")
    args = ap.parse_args()

    if args.list:
        for slug, (_, name, src, note) in LAYOUTS.items():
            print(f"{slug:18} {name:12} {src:34} {note}")
        return 0

    with open(REGISTRY, encoding="utf-8") as fh:
        reg = json.load(fh)
    themes = {t["slug"]: t for t in reg["themes"] if t.get("ppt_safe")}
    if args.theme not in themes:
        print(f"没有这个主题: {args.theme}\n可用: {', '.join(sorted(themes))}")
        return 1
    theme = themes[args.theme]

    slugs = list(LAYOUTS)
    if args.only:
        want = [s.strip() for s in args.only.split(",") if s.strip()]
        unknown = [s for s in want if s not in LAYOUTS]
        if unknown:
            print(f"未知布局: {unknown}")
            return 1
        slugs = want

    w, h = (float(v) for v in args.size.lower().split("x"))
    prs = Presentation()
    prs.slide_width = Emu(int(w * 914400))
    prs.slide_height = Emu(int(h * 914400))

    for i, slug in enumerate(slugs, 1):
        content = CONTENT.get(slug)
        if content is None:
            print(f"缺内容: {slug}")
            return 1
        build_slide(prs, theme, slug, content, num=i, total=len(slugs), canvas=(w, h))

    prs.save(args.out)
    print(f"{args.out}  {len(slugs)} 页  {w}×{h} in  主题 {args.theme}"
          f"（{theme['font_preset']} / {theme['name_zh']}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
