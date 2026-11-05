#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""python-pptx 排版小工具：颜色、字体（含 a:ea）、文本框、色块、栅格。

两个预览脚本（preview_themes / preview_layouts）共用这里的实现，
避免「一处修了 p:style 另一处没修」这类分歧。

要点：
- set_fonts() 同时写 a:latin / a:ea / a:cs：python-pptx 的 font.name 只写 a:latin，
  中文会掉回默认字体，一行里出现两种字形。
- rect() 删掉 <p:style>：AddShape 自带主题描边（lnRef）和主题阴影（effectRef），
  只加 a:noFill 仍会渲染出阴影。
- textbox() 的 line_pct 是 a:spcPct 千分比，实测 100000 = 1.2 倍字号。
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

EMU = 914400


def rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def lum(hexstr: str) -> float:
    h = hexstr.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a: str, b: str) -> float:
    """WCAG 对比度。反色版式前必须查：深底上再用深强调色，字会消失。"""
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return round((hi + 0.05) / (lo + 0.05), 2)


def mix(a: str, b: str, t: float) -> str:
    """两色线性混合，t=0 得 a。pptx 没有透明度，浅底/斑马纹只能预混。"""
    a, b = a.lstrip("#"), b.lstrip("#")
    out = []
    for i in (0, 2, 4):
        va, vb = int(a[i:i + 2], 16), int(b[i:i + 2], 16)
        out.append(f"{round(va + (vb - va) * t):02X}")
    return "#" + "".join(out)


def set_fonts(run, latin: str, ea: str, mono: str | None = None):
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", ea), ("a:cs", mono or ea)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", face)


def textbox(slide, x, y, w, h, *, text, size, color, latin, ea,
            bold=False, italic=False, align=PP_ALIGN.LEFT, anchor="top",
            line_pct=None, track_pt=None, wrap=True, space_after_pt=None,
            bullet=None):
    """多行文本。text 用 \\n 分段；每段的 line_pct/track 一致，够排版用。"""
    anchor_map = {"top": 1, "middle": 3, "bottom": 4}
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor_map[anchor]
    for idx, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        pPr = p._p.get_or_add_pPr()
        if line_pct is not None:
            ln = pPr.makeelement(qn("a:lnSpc"), {})
            pPr.insert(0, ln)
            ln.append(ln.makeelement(qn("a:spcPct"), {"val": str(int(line_pct))}))
        if space_after_pt:
            sa = pPr.makeelement(qn("a:spcAft"), {})
            pPr.append(sa)
            sa.append(sa.makeelement(qn("a:spcPts"), {"val": str(int(space_after_pt * 100))}))
        if bullet:
            pPr.set("marL", str(int(bullet[0] * EMU)))
            pPr.set("indent", str(int(-bullet[0] * EMU)))
            bu = pPr.makeelement(qn("a:buChar"), {"char": bullet[1]})
            pPr.append(bu)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = rgb(color)
        set_fonts(r, latin, ea)
        if track_pt:
            rPr = r._r.get_or_add_rPr()
            rPr.set("spc", str(int(track_pt * 100)))
    return tb


def rect(slide, x, y, w, h, *, fill=None, line=None, line_w_pt=None,
         shape=MSO_SHAPE.RECTANGLE, dash=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp = sh._element
    style = sp.find(qn("p:style"))
    if style is not None:
        sp.remove(style)                      # 去掉主题描边 + 主题阴影
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(line_w_pt or 0.75)
        if dash:
            ln = sh.line._get_or_add_ln()
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    else:
        ln = sh.line._get_or_add_ln()
        for child in list(ln):
            ln.remove(child)
        ln.append(ln.makeelement(qn("a:noFill"), {}))
    sh.shadow.inherit = False
    if sh.has_text_frame:
        sh.text_frame.word_wrap = True
    return sh


# ---------------------------------------------------------------- 栅格
class Grid:
    """12 栏栅格。M 边距、G 间隙都可调，列宽由画布宽推导。"""

    def __init__(self, width_in: float, height_in: float, margin: float = 0.8,
                 gutter: float = 0.2, cols: int = 12):
        self.W, self.H = width_in, height_in
        self.M, self.G, self.cols = margin, gutter, cols
        self.content_w = width_in - 2 * margin
        self.col_w = (self.content_w - (cols - 1) * gutter) / cols

    def x(self, col: int) -> float:
        """第 col 栏左边缘（0 起）。"""
        return self.M + col * (self.col_w + self.G)

    def span(self, col: int, n: int) -> float:
        return n * self.col_w + (n - 1) * self.G

    def cols_x(self, start: int, n: int, count: int, *, gap: float | None = None):
        """把 n 栏切成 count 段等宽并返回左边缘列表（用于 2/3/4 等分）。"""
        gap = self.G if gap is None else gap
        total = self.span(start, n)
        unit = (total - (count - 1) * gap) / count
        return [self.x(start) + i * (unit + gap) for i in range(count)], unit
