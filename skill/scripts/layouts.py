#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""布局原型库：13 种结构性不同的单页版式，与配色主题正交。

来源是 refs/html-ppt-skill/templates/single-page/ 的栅格签名，我把它翻成 pptx 的英寸坐标：
  image-text-split 1fr/1fr · comparison 1fr/90px/1fr · timeline 5 等分
  kpi-grid 4 等分 · process-steps 4 等分 · table/code 无栅格

每个布局函数签名统一 (slide, th, g, c)：
  th = 主题（色 + 字体 + 字号阶），g = 12 栏栅格，c = 该页内容 dict。
**几何全部由 g 和 th 推导**，所以同一套布局换主题不用改坐标——
72pt 的 poster 主题和 54pt 的 editorial 主题都能直接跑。

用法：from layouts import LAYOUTS, build_slide
"""
from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from pptx_kit import Grid, contrast, mix, rgb, rect, textbox

L, C, R = PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT
TOP, MID, BOT = "top", "middle", "bottom"


class Theme:
    """把 themes.json 的一条包装成排版可用的对象。"""

    def __init__(self, t: dict):
        p = t["palette"]
        self.slug = t["slug"]
        self.paper, self.ink = p["paper"], p["ink"]
        self.accent, self.accent2, self.rule = p["accent"], p["accent2"], p["rule"]
        self.on_accent = t.get("accent_text_color", self.ink)
        self.accent_role = t.get("accent_role", "text")
        f = t["fonts"]
        self.d_lat, self.d_ea = f["display_latin"], f["display_ea"]
        self.b_lat, self.b_ea = f["body_latin"], f["body_ea"]
        self.mono = f["mono"]
        ty = t["type"]
        self.display = ty["display_pt"]
        self.h1, self.lead, self.body = ty["h1_pt"], ty["lead_pt"], ty["body_pt"]
        self.small, self.label = ty["small_pt"], ty["label_pt"]
        self.d_lh, self.b_lh = ty["display_lh_pct"], ty["body_lh_pct"]
        self.track_label = ty["tracking_label_pt"]
        # 反色页整页铺 ink，深底上再用深强调色字会消失——实测 14 套里 7 套 accent/ink < 3.0。
        # 依次试 accent → 与 paper 混到 0.35 / 0.6 → paper，取第一个 ≥ 3.0 的。
        self.on_ink = self.paper
        for cand in (self.accent,
                     mix(self.accent, self.paper, 0.35),
                     mix(self.accent, self.paper, 0.6),
                     self.paper):
            if contrast(cand, self.ink) >= 3.0:
                self.on_ink = cand
                break
        # 强调色有两种用法，别混：
        #   accent_text —— 当**文字**（数字、标题）。role=block 或与纸色对比不足时退回墨色。
        #   accent_mark —— 当**细线/小圆点**。太亮的强调色在白纸上看不见，往墨色压一点。
        # accent 本身只用于面积 ≥ 0.5in 的色块，那时候亮色反而正是效果。
        self.accent_text = (self.accent
                            if self.accent_role == "text" and contrast(self.accent, self.paper) >= 3.0
                            else self.ink)
        self.accent_mark = (self.accent if contrast(self.accent, self.paper) >= 3.0
                            else mix(self.accent, self.ink, 0.5))
        self.accent2_mark = (self.accent2 if contrast(self.accent2, self.paper) >= 3.0
                             else mix(self.accent2, self.ink, 0.45))

    # 常用组合
    def label_style(self):
        """等宽小标签的公共参数（不含 color，由调用方决定用墨色还是纸色）。"""
        return dict(size=self.label, latin=self.mono, ea=self.b_ea,
                    track_pt=self.track_label)

    def best_on(self, bg_color: str) -> str:
        """压在某个色块上的字用墨色还是纸色，按对比度选。"""
        return (self.ink if contrast(self.ink, bg_color) >= contrast(self.paper, bg_color)
                else self.paper)

    def tint_pair(self, color: str, need: float = 4.5):
        """给「图位/色块」挑色调 + 压在上面的字色，保证对比度 ≥ need。

        只调字不够：深色主题里 mix 出来的中灰上，墨色（亮黄）和纸色（近黑）都不够看，
        得反过来调块的深浅。所以这里返回 (块色, 字色) 两个值。
        """
        for f in (0.62, 0.35, 0.2, 0.8, 0.5):
            block = mix(color, self.paper, f)
            tc = self.best_on(block)
            if contrast(tc, block) >= need:
                return block, tc
        return mix(color, self.paper, 0.62), self.ink

    def zebra(self):
        return mix(self.paper, self.rule, 0.5)


def bg(slide, g, color):
    rect(slide, 0, 0, g.W, g.H, fill=color)


def label(slide, th, x, y, text, w=6.0, color=None):
    textbox(slide, x, y, w, 0.2, text=text, color=color or th.ink, **th.label_style())


def hairline(slide, th, x, y, w, color=None, h=0.008):
    rect(slide, x, y, w, h, fill=color or th.rule)


def footer(slide, th, g, num, total, note=""):
    label(slide, th, g.M, g.H - 0.5, note or f"{num:02d} / {total:02d}", w=g.content_w * 0.7)
    textbox(slide, g.W - g.M - 0.9, g.H - 0.5, 0.9, 0.2, text=f"{num:02d}",
            color=th.ink, size=th.label, latin=th.mono, ea=th.b_ea, align=R)


# ---------------------------------------------------------------- 1 封面·纯文字
def cover_type(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    size = min(104, th.display * 1.6)
    textbox(slide, g.M, 2.25, g.span(0, 9), size * 0.92 / 72 * 2 + 0.4,
            text=c["title"], size=size, color=th.ink, latin=th.d_lat, ea=th.d_ea,
            line_pct=th.d_lh, anchor=TOP)
    hairline(slide, th, g.M, 5.62, g.span(0, 4))
    textbox(slide, g.M, 5.86, g.span(0, 6), 0.9, text=c["lead"],
            size=th.body, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh)
    textbox(slide, g.x(8), 5.62, g.span(8, 4), 1.2, text=c["meta"],
            size=th.small, color=th.ink, latin=th.b_lat, ea=th.b_ea,
            align=R, line_pct=th.b_lh)


# ---------------------------------------------------------------- 2 封面·图文对分
def cover_split(slide, th, g, c):
    bg(slide, g, th.paper)
    bx = g.x(7)
    bw = g.W - g.M - bx
    block, block_text = th.tint_pair(th.accent2)   # 图位：块色与字色成对定，别各挑各的
    rect(slide, bx, 0, bw, g.H, fill=block)
    rect(slide, bx + bw / 2 - 0.5, g.H / 2 - 0.5, 1.0, 0.012, fill=block_text)
    label(slide, th, bx + 0.4, g.H - 0.9, "IMG · 1600 × 1200", w=bw - 0.8,
          color=block_text)
    label(slide, th, g.M, g.M * 0.75, c["kicker"], w=g.span(0, 6))
    textbox(slide, g.M, 1.9, g.span(0, 6), 3.2, text=c["title"],
            size=min(72, th.display), color=th.ink, latin=th.d_lat, ea=th.d_ea,
            line_pct=th.d_lh)
    hairline(slide, th, g.M, 5.35, g.span(0, 3))
    textbox(slide, g.M, 5.6, g.span(0, 5), 1.0, text=c["lead"],
            size=th.body, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh)


# ---------------------------------------------------------------- 3 章节页（反色）
def section_divider(slide, th, g, c):
    bg(slide, g, th.ink)                       # 整页反色：底=墨色，字=纸色
    fg = th.paper
    textbox(slide, g.M, 1.0, 5.2, 2.35, text=c["num"], size=180, color=th.on_ink,
            latin=th.d_lat, ea=th.d_ea, line_pct=85000)
    textbox(slide, g.M + 0.06, 3.45, g.span(2, 10), 0.9, text=c["title"],
            size=th.display * 0.8, color=fg, latin=th.d_lat, ea=th.d_ea,
            line_pct=th.d_lh)
    hairline(slide, th, g.M + 0.06, 4.75, g.span(2, 3), color=th.on_ink)
    textbox(slide, g.M + 0.06, 4.98, g.span(2, 7), 0.7, text=c["lead"],
            size=th.body, color=fg, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh)
    label(slide, th, g.M, g.H - 0.62, c["kicker"], color=fg)


# ---------------------------------------------------------------- 4 左标题 + 右双栏正文
def two_column_text(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    # 标题贴到版心底部：正文只占上半页时，上面留白会被读成「设计」而不是「内容被砍了」
    textbox(slide, g.M, 3.05, g.span(0, 4), 2.55, text=c["title"],
            size=th.h1, color=th.ink, latin=th.d_lat, ea=th.d_ea,
            line_pct=max(th.d_lh, 100000), anchor=BOT)
    hairline(slide, th, g.M, 5.92, g.span(0, 4))
    xs, cw = g.cols_x(5, 7, 2, gap=0.4)
    for i, para in enumerate(c["paras"]):
        textbox(slide, xs[i], 1.62, cw, 4.5, text=para,
                size=th.body + 1, color=th.ink, latin=th.b_lat, ea=th.b_ea,
                line_pct=th.b_lh, space_after_pt=10)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 5 三栏卡片
def three_cards(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.2, g.span(0, 8), 0.9, text=c["title"],
            size=th.h1 * 0.86, color=th.ink, latin=th.d_lat, ea=th.d_ea,
            line_pct=th.d_lh)
    xs, cw = g.cols_x(0, 12, 3, gap=0.34)
    top, ch = 2.7, 2.6                              # 卡片高度按内容定，别拉成 3.4 留一截空
    for i, (head, body) in enumerate(c["cards"]):
        x = xs[i]
        rect(slide, x, top, cw, ch, fill=mix(th.paper, th.rule, 0.35))
        rect(slide, x, top, cw, 0.055, fill=th.accent_mark)
        label(slide, th, x + 0.3, top + 0.28, f"0{i+1}", w=1.0)
        textbox(slide, x + 0.3, top + 0.62, cw - 0.6, 0.5, text=head,
                size=20, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=110000)
        textbox(slide, x + 0.3, top + 1.22, cw - 0.6, 1.1, text=body,
                size=th.body - 1, color=th.ink, latin=th.b_lat, ea=th.b_ea,
                line_pct=th.b_lh)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 6 四个数字
def kpi_grid(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 8), 0.7, text=c["title"],
            size=th.h1, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    xs, cw = g.cols_x(0, 12, 4, gap=0.0)
    top = 2.5
    for i, (num, unit, desc) in enumerate(c["kpis"]):
        x = xs[i]
        if i:
            rect(slide, x - 0.02, top, 0.008, 2.7, fill=th.rule)
        textbox(slide, x + 0.36, top, cw - 0.5, 1.25, text=num,
                size=min(76, th.display * 1.2), color=th.accent_text,
                latin=th.d_lat, ea=th.d_ea, line_pct=95000)
        textbox(slide, x + 0.36, top + 1.28, cw - 0.5, 0.34, text=unit,
                size=th.small, color=th.ink, latin=th.mono, ea=th.b_ea)
        hairline(slide, th, x + 0.36, top + 1.74, cw - 0.7)
        textbox(slide, x + 0.36, top + 1.92, cw - 0.56, 0.9, text=desc,
                size=th.small, color=th.ink, latin=th.b_lat, ea=th.b_ea,
                line_pct=th.b_lh)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 7 单数字主打
def stat_highlight(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    if th.accent_role == "block":              # 亮强调色不能当字：它当底，字用墨色
        rect(slide, g.M - 0.05, 1.95, g.span(0, 8) + 0.1, 4.0, fill=th.accent)
        num_color = th.best_on(th.accent)
    else:
        num_color = th.accent_text
    textbox(slide, g.M - 0.1, 1.95, g.span(0, 8), 4.0, text=c["num"],
            size=230, color=num_color, latin=th.d_lat, ea=th.d_ea,
            line_pct=80000, anchor=MID)          # 和色块同高同锚，数字才居中
    textbox(slide, g.x(9), 2.45, g.span(9, 3), 0.4, text=c["unit"],
            size=th.lead, color=th.ink, latin=th.mono, ea=th.b_ea)
    hairline(slide, th, g.x(9), 2.95, g.span(9, 3))
    textbox(slide, g.x(9), 3.15, g.span(9, 3), 2.2, text=c["desc"],
            size=th.body, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 8 引言
def big_quote(slide, th, g, c):
    bg(slide, g, th.paper)
    textbox(slide, g.M - 0.15, 0.72, 1.6, 2.3, text="“", size=150,
            color=th.accent2_mark, latin=th.d_lat, ea=th.d_ea)
    label(slide, th, g.x(9), g.M * 0.8, c["kicker"], w=g.span(9, 3))
    textbox(slide, g.M + 0.5, 2.15, g.span(1, 11), 2.9, text=c["quote"],
            size=min(40, th.h1 + 4), color=th.ink, latin=th.d_lat, ea=th.d_ea,
            line_pct=125000)
    hairline(slide, th, g.M + 0.5, 5.5, g.span(1, 3), color=th.accent_mark)
    textbox(slide, g.M + 0.5, 5.72, g.span(1, 8), 0.3, text=c["by"],
            size=th.small, color=th.ink, latin=th.mono, ea=th.b_ea)
    footer(slide, th, g, c["_num"], c["_total"], "")


# ---------------------------------------------------------------- 9 左右对照
def comparison(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 9), 0.7, text=c["title"],
            size=th.h1, color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    left_w, right_x = g.span(0, 5), g.x(7)
    panels = ((g.M, left_w, c["left"], th.ink),
              (right_x, g.W - g.M - right_x, c["right"], th.accent_mark))
    top = 2.3
    for x, w, (head, items), bar in panels:
        rect(slide, x, top, w, 0.07, fill=bar)
        textbox(slide, x, top + 0.3, w, 0.42, text=head, size=20, color=th.ink,
                latin=th.b_lat, ea=th.b_ea, line_pct=110000)
        textbox(slide, x, top + 1.0, w, 2.6,
                text="\n".join(items), size=th.body, color=th.ink,
                latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh, space_after_pt=8,
                bullet=(0.22, "·"))
    textbox(slide, g.x(5), 3.9, g.span(5, 2), 0.3, text="VS", size=th.small,
            color=th.ink, latin=th.mono, ea=th.b_ea, align=C)
    rect(slide, g.x(5) + g.span(5, 2) / 2, top, 0.008, 2.9, fill=th.rule)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 10 时间轴
def timeline(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 8), 0.7, text=c["title"], size=th.h1,
            color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    axis_y = 3.75
    rect(slide, g.M, axis_y, g.content_w, 0.01, fill=th.rule)
    xs, cw = g.cols_x(0, 12, len(c["points"]), gap=0.24)
    for i, (when, what) in enumerate(c["points"]):
        cx = xs[i] + cw / 2
        filled = i == len(c["points"]) - 1
        rect(slide, cx - 0.085, axis_y - 0.075, 0.17, 0.17,
             fill=th.accent_mark if filled else th.paper, line=th.accent_mark, line_w_pt=1.5,
             shape=MSO_SHAPE.OVAL)
        textbox(slide, xs[i], axis_y - 0.72, cw, 0.3, text=when, size=18,
                color=th.ink, latin=th.d_lat, ea=th.d_ea, align=C)
        textbox(slide, xs[i], axis_y + 0.34, cw, 1.5, text=what, size=th.small,
                color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh, align=C)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 11 四步流程
def process_steps(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 8), 0.7, text=c["title"], size=th.h1,
            color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    xs, cw = g.cols_x(0, 12, 4, gap=0.3)
    top = 2.55
    for i, (head, body) in enumerate(c["steps"]):
        x = xs[i]
        textbox(slide, x, top, 0.8, 0.3, text=f"0{i+1}", size=th.label,
                color=th.ink, latin=th.mono, ea=th.b_ea, track_pt=th.track_label)
        textbox(slide, x, top + 0.42, cw, 0.75, text=head, size=22, color=th.ink,
                latin=th.b_lat, ea=th.b_ea, line_pct=110000)
        hairline(slide, th, x, top + 1.28, cw - 0.3, color=th.accent_mark)
        textbox(slide, x, top + 1.46, cw - 0.15, 1.9, text=body, size=th.small,
                color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.b_lh)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 12 数据表
def table(slide, th, g, c):
    bg(slide, g, th.paper)
    label(slide, th, g.M, g.M * 0.75, c["kicker"])
    textbox(slide, g.M, 1.15, g.span(0, 8), 0.7, text=c["title"], size=th.h1,
            color=th.ink, latin=th.b_lat, ea=th.b_ea, line_pct=th.d_lh)
    cols, rows = c["cols"], c["rows"]
    widths = c["widths"]
    aligns = [{"left": L, "center": C, "right": R}.get(a, L)
              for a in c.get("aligns", [])] or [L] * len(cols)   # 数字列右对齐才立得住
    pad = 0.18
    top, height = 2.35, 0.5
    # 表头
    x = g.M
    rect(slide, g.M, top, sum(widths), height, fill=th.ink)
    for i, name in enumerate(cols):
        textbox(slide, x + pad, top + 0.14, widths[i] - 2 * pad, 0.26, text=name,
                size=th.small, color=th.paper, latin=th.mono, ea=th.b_ea,
                align=aligns[i])
        x += widths[i]
    # 数据行：斑马纹用 paper 与 rule 各半混出来的浅色，不用透明度
    for r, row in enumerate(rows):
        y = top + height * (r + 1)
        if r % 2 == 0:
            rect(slide, g.M, y, sum(widths), height, fill=th.zebra())
        hairline(slide, th, g.M, y + height - 0.006, sum(widths), h=0.006)
        x = g.M
        for i, cell in enumerate(row):
            last = i == len(row) - 1
            textbox(slide, x + pad, y + 0.13, widths[i] - 2 * pad, 0.3, text=cell,
                    size=th.small,
                    color=th.accent if last and th.accent_role == "text" else th.ink,
                    latin=th.mono if i == 0 or last else th.b_lat, ea=th.b_ea,
                    align=aligns[i])
            x += widths[i]
    if c.get("note"):
        textbox(slide, g.M, top + height * (len(rows) + 1) + 0.22, sum(widths), 0.3,
                text=c["note"], size=th.label, color=th.ink, latin=th.mono,
                ea=th.b_ea, track_pt=th.track_label)
    footer(slide, th, g, c["_num"], c["_total"], c.get("foot", ""))


# ---------------------------------------------------------------- 13 结尾
def closing(slide, th, g, c):
    bg(slide, g, th.ink)
    fg = th.paper
    textbox(slide, g.M, 2.35, g.content_w, 1.6, text=c["title"],
            size=min(88, th.display * 1.35), color=fg, latin=th.d_lat, ea=th.d_ea,
            align=C, line_pct=th.d_lh)
    hairline(slide, th, g.W / 2 - 1.1, 4.35, 2.2, color=th.on_ink)
    textbox(slide, g.M, 4.6, g.content_w, 0.4, text=c["contact"], size=th.small,
            color=fg, latin=th.mono, ea=th.b_ea, align=C, track_pt=0.4)
    label(slide, th, g.M, g.H - 0.62, c["kicker"], color=fg)


# ---------------------------------------------------------------- 注册表
LAYOUTS = {
    "cover-type": (cover_type, "封面·纯文字", "cover.html",
                   "全幅字 + 大留白；没有图也能撑住封面"),
    "cover-split": (cover_split, "封面·图文对分", "image-text-split.html 1fr/1fr",
                    "右 5 栏整高图位，左文字块"),
    "section-divider": (section_divider, "章节页·反色", "section-divider.html",
                        "整页墨底反色 + 180pt 章节号"),
    "two-column-text": (two_column_text, "左标题 + 右双栏", "two-column.html",
                        "4 栏标题 / 8 栏正文再切两栏"),
    "three-cards": (three_cards, "三栏卡片", "three-column.html",
                    "等宽三卡，顶部 0.055in 强调条"),
    "kpi-grid": (kpi_grid, "四宫格数字", "kpi-grid.html 4 等分",
                 "76pt 数字 + 单位 + 说明，列间竖发丝线"),
    "stat-highlight": (stat_highlight, "单数字主打", "stat-highlight.html",
                       "230pt 数字压左，右侧三行解释"),
    "big-quote": (big_quote, "大引言", "big-quote.html",
                  "150pt 引号 + 40pt 引言，出处等宽右下"),
    "comparison": (comparison, "左右对照", "comparison.html 1fr/90px/1fr",
                   "中缝 VS + 两条 0.07in 色条区分两栏"),
    "timeline": (timeline, "横向时间轴", "timeline.html 5 等分",
                 "轴线 + 圆点（实心=当前），时间在上说明在下"),
    "process-steps": (process_steps, "四步流程", "process-steps.html 4 等分",
                      "编号 + 标题 + 强调发丝线 + 说明"),
    "table": (table, "数据表", "table.html",
              "真 pptx 表格：墨色表头 + 混色斑马纹（不用透明度）"),
    "closing": (closing, "结尾·反色", "cta.html / thanks.html",
                "居中大字 + 联系方式，与章节页同族"),
}


def build_slide(prs, theme_json: dict, layout_slug: str, content: dict,
                num: int = 1, total: int = 1, canvas=(13.333, 7.5),
                margin: float = 0.8):
    th = Theme(theme_json)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    g = Grid(canvas[0], canvas[1], margin=margin)
    c = dict(content)
    c["_num"], c["_total"] = num, total
    LAYOUTS[layout_slug][0](slide, th, g, c)
    return slide
