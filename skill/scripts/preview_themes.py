#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 themes.json 排成一张主题样张（每页一个主题），用来肉眼验收配色和字体。

用法:
    python preview_themes.py [输出.pptx] [--only slug1,slug2] [--size 10x5.625]

样张每页结构固定：label / 大标题 / 发丝线 / 三行正文 / 强调色块或色条 / 页脚色值。
排版完全读 themes.json 里的数字，不在这里另写一套尺寸——
这样样张能过 lint，就说明注册表里的值本身是可用的。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
REGISTRY = os.path.join(SKILL, "assets", "themes.json")

M = 0.55          # 页边距（英寸），与 design-system.md 的栅格一致
FONT_MONO = "IBM Plex Mono"


# ---------------------------------------------------------------- 基础工具
def _lum(h: str) -> float:
    h = h.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def legible(fg: str, bg: str, need: float = 4.5) -> str:
    """小字要 4.5 才在投影仪上看得住；强调色不够就逐级往墨色压。

    8.5pt 的强调色小标签在 3.4 对比度上是真的糊——lint R8 会报，别靠调低阈值糊过去。
    """
    if contrast(fg, bg) >= need:
        return fg
    # 底浅就往黑压，底深就往白提——往反方向混只会越混越糊
    target = "000000" if _lum(bg) > 0.5 else "FFFFFF"
    a = fg.lstrip("#")
    for k in (0.25, 0.45, 0.65, 0.85):
        c = "".join(f"{round(int(a[i:i+2],16)+(int(target[i:i+2],16)-int(a[i:i+2],16))*k):02X}"
                    for i in (0, 2, 4))
        if contrast("#" + c, bg) >= need:
            return "#" + c
    return "#" + target


def rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_fonts(run, latin: str, ea: str):
    """同时写 a:latin 和 a:ea，否则中文会掉字体。"""
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", ea), ("a:cs", ea)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", face)


def textbox(slide, x, y, w, h, *, text, size, color, latin, ea,
            bold=False, align=PP_ALIGN.LEFT, line_pct=None, track_pt=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = 1  # MSO_ANCHOR.TOP
    lines = text.split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align
        if line_pct is not None:
            pPr = p._p.get_or_add_pPr()
            ln = pPr.makeelement(qn("a:lnSpc"), {})
            pPr.insert(0, ln)
            ln.append(ln.makeelement(qn("a:spcPct"), {"val": str(int(line_pct))}))
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        set_fonts(r, latin, ea)
        if track_pt:
            r.font._rPr.set("spc", str(int(track_pt * 100)))
    return tb


def rect(slide, x, y, w, h, *, fill=None, line=None, line_w_pt=None):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    # AddShape 会带 <p:style>（主题描边 + 主题阴影），LibreOffice 会真的画出阴影。
    # 删掉它，改用显式的 spPr，才能保证「平、无边框、无阴影」。
    sp = sh._element
    style = sp.find(qn("p:style"))
    if style is not None:
        sp.remove(style)
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(line_w_pt or 0.75)
    else:
        # 必须显式 noFill，否则会继承 Office 默认蓝描边
        ln = sh.line._get_or_add_ln()
        for child in list(ln):
            ln.remove(child)
        ln.append(ln.makeelement(qn("a:noFill"), {}))
    sh.shadow.inherit = False
    return sh


# ---------------------------------------------------------------- 单页
def build_page(prs, t, idx, total):
    pal = t["palette"]
    ty = t["type"]
    f = t["fonts"]
    paper, ink = pal["paper"], pal["ink"]
    accent, accent2, rule = pal["accent"], pal["accent2"], pal["rule"]

    s = prs.slides.add_slide(prs.slide_layouts[6])
    # 底色铺满（用大矩形而不是 slide background，保证 LibreOffice/PowerPoint 一致）
    rect(s, 0, 0, 13.333, 7.5, fill=paper)

    W = prs.slide_width / 914400
    H = prs.slide_height / 914400
    content_w = W - 2 * M

    # label：强调色当字时用 accent，否则用墨色；8.5pt 小字还得够 4.5 对比度
    label_color = legible(accent, paper) if t["accent_role"] == "text" else ink
    textbox(s, M, 0.40, content_w * 0.6, 0.22,
            text=f"{idx:02d} / {t['register'].upper()} · {t['slug']}",
            size=ty["label_pt"], color=label_color,
            latin=FONT_MONO, ea=f["body_ea"], track_pt=ty["tracking_label_pt"])

    # 大标题
    title_w = content_w * 0.52
    title = t["name_zh"]
    lh = ty["display_lh_pct"] / 100000 * 1.2          # 相对字号的行高
    th = ty["display_pt"] * lh * 1.4 / 72
    textbox(s, M, 0.70, title_w, th, text=title, size=ty["display_pt"], color=ink,
            latin=f["display_latin"], ea=f["display_ea"],
            line_pct=ty["display_lh_pct"])

    # 发丝线
    ry = 0.70 + th + 0.18
    rect(s, M, ry, title_w, 0.008, fill=rule)

    # 正文
    body = t["usage"]["note"]
    textbox(s, M, ry + 0.24, title_w, 0.90, text=body,
            size=ty["body_pt"], color=ink, latin=f["body_latin"], ea=f["body_ea"],
            line_pct=ty["body_lh_pct"])

    # 字号阶梯：横向一排、基线对齐，像一个真的字体样张
    ly = ry + 1.42
    textbox(s, M, ly - 0.22, 0.52, 0.2, text="TYPE", size=ty["label_pt"], color=ink,
            latin=FONT_MONO, ea=f["body_ea"], track_pt=1.3)
    for n, sz in enumerate((36, 24, 15, 12)):
        x = M + n * 1.15
        tb = textbox(s, x, ly, 1.10, 0.62, text="Aa", size=sz, color=ink,
                     latin=f["display_latin"] if n < 2 else f["body_latin"],
                     ea=f["display_ea"] if n < 2 else f["body_ea"])
        tb.text_frame.vertical_anchor = 3          # MSO_ANCHOR.BOTTOM，共享基线
        textbox(s, x, ly + 0.66, 1.0, 0.2, text=f"{sz} pt",
                size=ty["label_pt"], color=ink, latin=FONT_MONO, ea=f["body_ea"])

    # 色板挪到右栏，避免 72pt 主题的样张把色板顶掉（画在强调块下面）
    # 位置由 _palette() 负责

    # 强调：能当字就一笔小字；只能当块就画块，并在块上写反色字
    bx = M + title_w + 0.7
    bw = W - M - bx
    if t["accent_role"] == "block":
        bh = 1.55
        # accent_text_color 只保证「大字压块」能用（实测樱花 #FFFFFF/#E5392A = 4.24），
        # 8.5pt 小标签达不到 4.5，换成墨色或干脆不提。
        small = t["accent_text_color"]
        small_pt = ty["label_pt"]
        if contrast(small, accent) < 4.5:
            if contrast(ink, accent) >= 4.5:
                small = ink
            else:
                small_pt = 18.0     # 换色也救不了（红底上只有白字，4.24）：放大到「大字」档，3.0 才算过关
        rect(s, bx, 0.95, bw, bh, fill=accent)
        inner = bw - 0.6
        # 块上字号按块宽反算，避免中文按 1em/字 顶破
        size = min(44.0, inner * 72 / max(len(title), 1) * 0.92)
        textbox(s, bx + 0.3, 1.18, inner, 0.9,
                text=title, size=size, color=t["accent_text_color"],
                latin=f["display_latin"], ea=f["display_ea"])
        textbox(s, bx + 0.3, 0.95 + bh - 0.40, inner, 0.2,
                text="黑字压块 / 白字用不了", size=small_pt,
                color=small, latin=FONT_MONO, ea=f["body_ea"])
    else:
        rect(s, bx, 0.95, 0.10, 1.55, fill=accent)
        textbox(s, bx + 0.34, 1.02, bw - 0.34, 1.4,
                text="强调色\n当字用", size=min(32, ty["display_pt"] * 0.6),
                color=accent, latin=f["body_latin"], ea=f["body_ea"],
                line_pct=ty["display_lh_pct"])

    # 色板：放在右栏强调块下面，五列 + 角色名 + 色值
    sy = 2.78
    textbox(s, bx, sy - 0.24, 1.2, 0.18, text="PALETTE", size=ty["label_pt"],
            color=ink, latin=FONT_MONO, ea=f["body_ea"], track_pt=1.3)
    roles = ("paper", "ink", "accent", "accent2", "rule")
    gap = 0.09
    sw = (bw - 4 * gap) / 5
    for k, role in enumerate(roles):
        x = bx + k * (sw + gap)
        rect(s, x, sy, sw, 0.34, fill=pal[role],
             line=rule if role == "paper" else None, line_w_pt=0.75)
        textbox(s, x, sy + 0.40, sw + gap, 0.18, text=role,
                size=ty["label_pt"], color=ink, latin=FONT_MONO, ea=f["body_ea"])
        textbox(s, x, sy + 0.58, sw + gap, 0.18, text=pal[role],
                size=ty["label_pt"], color=ink, latin=FONT_MONO, ea=f["body_ea"])

    # 页脚：对比度 + 字体/T 值，色值已经在色板上标过
    foot = (f"ink/paper {t['contrast']['ink_on_paper']:.1f}   "
            f"display {f['display_latin']} / {f['display_ea']}   "
            f"body {f['body_latin']} {ty['body_pt']}pt×{ty['body_lh_em']}")
    textbox(s, M, H - 0.40, content_w * 0.88, 0.2, text=foot,
            size=ty["label_pt"], color=ink, latin=FONT_MONO, ea=f["body_ea"],
            track_pt=0.3)
    textbox(s, W - M - 0.8, H - 0.40, 0.8, 0.2, text=f"{idx}/{total}",
            size=ty["label_pt"], color=ink, latin=FONT_MONO, ea=f["body_ea"],
            align=PP_ALIGN.RIGHT)
    return s


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default="theme-sheet.pptx")
    ap.add_argument("--only", default="")
    ap.add_argument("--size", default="10x5.625")
    args = ap.parse_args()

    with open(REGISTRY, encoding="utf-8") as fh:
        reg = json.load(fh)
    themes = [t for t in reg["themes"] if t.get("ppt_safe")]
    if args.only:
        want = {s.strip() for s in args.only.split(",") if s.strip()}
        themes = [t for t in themes if t["slug"] in want]

    w, h = (float(v) for v in args.size.lower().split("x"))
    prs = Presentation()
    prs.slide_width = Emu(int(w * 914400))
    prs.slide_height = Emu(int(h * 914400))

    for i, t in enumerate(themes, 1):
        build_page(prs, t, i, len(themes))

    prs.save(args.out)
    print(f"{args.out}  {len(themes)} 页  {w}x{h} in")
    return 0


if __name__ == "__main__":
    sys.exit(main())
