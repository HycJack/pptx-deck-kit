# -*- coding: utf-8 -*-
"""pptx 版面检查：越界、重叠、文字溢出、字号对比度、主题蓝残留、文字/底色对比度。

    python lint_pptx.py deck.pptx [--json out.json]

退出码 1 表示有 error 级问题。它是「生成完先自检」这一步，替代肉眼逐页看。
"""
import argparse
import json
import os
import re
import sys
from collections import Counter

from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.oxml.ns import qn
from pptx.util import Emu

EMU_IN = 914400
THEME_BLUE = "2F5496"      # Office 默认描边，通常是漏设 noFill 的证据
TOL = 0.02                  # 英寸容差

CJK = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]")


def inch(v):
    return (v or 0) / EMU_IN


def frames(shape):
    """可读的文本容器：[(TextFrame, 局部底色 or None, 几何 or None)]。

    GraphicFrame（表格、图表）**没有** `.text_frame`——`has_text_frame` 是 False，直接取就
    AttributeError（genppt 的 AddTable 产物会踩到；本仓库的 `table` 版式自己画文本框因此没暴露）。
    表格按单元格展开，并且真的算单元格几何：否则几行 cell 拼成一段文字，R2 会误报成溢出。
    """
    try:
        if shape.has_text_frame:
            return [(shape.text_frame, None, None)]
        if getattr(shape, "has_table", False):
            t = shape.table
            colw = [inch(c.width) for c in t.columns]
            rowh = [inch(r.height) for r in t.rows]
            x0, y0 = inch(shape.left), inch(shape.top)
            out, y = [], y0
            for ri, row in enumerate(t.rows):
                x = x0
                for ci, cell in enumerate(row.cells):
                    bg = None
                    try:
                        if cell.fill.type == MSO_FILL.SOLID:
                            bg = str(cell.fill.fore_color.rgb)
                    except (TypeError, AttributeError, ValueError):
                        bg = None
                    out.append((cell.text_frame, bg, (x, y, colw[ci], rowh[ri])))
                    x += colw[ci]
                y += rowh[ri]
            return out
    except (AttributeError, ValueError, TypeError, IndexError):
        pass
    return []


def text_of(shape):
    """段落之间必须补换行：python-pptx 把它们存成独立的 <a:p>，空串拼接会把行数算少。"""
    paras = []
    for tf, _, _ in frames(shape):
        paras.extend(p.text for p in tf.paragraphs)
    return "\n".join(paras)


def char_width_em(ch):
    return 1.0 if CJK.match(ch) else 0.55


def line_factor(size_pt):
    """行高倍数。字号越大越紧——和设计稿一致的排法（display 0.95–1.05，body 1.35+）。

    固定用 1.35 会把两行 54pt 的标题误判成溢出，这是早期版本的真实误报。
    """
    if size_pt >= 40:
        return 1.0
    if size_pt >= 20:
        return 1.2
    return 1.35


def est_lines(text, width_in, size_pt):
    """按字符宽度粗估能排几行。显式换行按段落分别算，不要折叠掉。"""
    if not text:
        return 0
    per_line = max(1.0, width_in * 72.0 / size_pt)        # 每行可容纳的 em 数
    total = 0
    for para in text.split("\n"):
        used = sum(char_width_em(c) for c in para)
        total += max(1, int(used / per_line + 0.999))
    return total


def est_text_width(text, size_pt):
    """粗估这段文字实际占多宽（英寸），用于判断「框重叠但字没重叠」。"""
    per_para = []
    for para in text.split("\n"):
        per_para.append(sum(char_width_em(c) for c in para) * size_pt / 72.0)
    return max(per_para) if per_para else 0.0


def text_rect(tf, x, y, w, h, txt, size):
    """按对齐方式把文本框收缩到文字真正占据的区间。"""
    tw = min(est_text_width(txt, size), w) if txt else 0.0
    al = None
    try:
        al = tf.paragraphs[0].alignment if tf is not None else None
    except (IndexError, AttributeError):
        pass
    if al is not None and int(al) == 2:        # CENTER
        tx = x + (w - tw) / 2
    elif al is not None and int(al) == 3:      # RIGHT
        tx = x + w - tw
    else:                                      # LEFT / 未设
        tx = x
    return (tx, y, tw, h, txt)


def slide_name(idx):
    return f"slide{idx}"


def _lum(hexstr):
    h = hexstr.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def filled_rects(slide):
    """所有实色填充的形状，按 z 序（文档顺序 = 从下往上）。"""
    out = []
    for sh in slide.shapes:
        try:
            if sh.fill.type != MSO_FILL.SOLID:
                continue
            out.append((inch(sh.left), inch(sh.top), inch(sh.width), inch(sh.height),
                        str(sh.fill.fore_color.rgb)))
        except (TypeError, AttributeError, ValueError):
            continue
    return out


def slide_bg(slide):
    """幻灯片背景色（`<p:bg>`）。拿不到就当白底。

    自己这套版式永远是「满幅色块打底」，所以以前只在形状填充里找；换成别人的产物
    （比如 genppt 的 `SetBackground`）就会拿白底去比白字，误报 1.00。
    """
    try:
        f = slide.background.fill
        if f.type == MSO_FILL.SOLID:
            return str(f.fore_color.rgb)
    except (TypeError, AttributeError, ValueError):
        pass
    return "FFFFFF"


def bg_under(x, y, w, h, fills, default="FFFFFF"):
    """文字框中心点落在哪个色块上——倒序找即取最上层，找不到就用 default（通常整页底色）。

    表头「纸色字压墨色块」只有查局部底色才算得对；用整页底色会误报成 1.00。
    """
    cx, cy = x + w / 2, y + h / 2
    for fx, fy, fw, fh, color in reversed(fills):
        if fx <= cx <= fx + fw and fy <= cy <= fy + fh:
            return color
    return default


def run_colors(shape, page_bg):
    """每个 run 的显式颜色 + 字号 + 压着的底色。没显式设色（继承布局）的跳过，避免误报。

    底色优先用「最近的一层实色填充」：表格单元格自己的填充 → 否则回到调用方算的局部/整页底色。
    """
    out = []
    for tf, local_bg, _ in frames(shape):
        for p in tf.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                try:
                    if r.font.color is None or r.font.color.type is None:
                        continue
                    rgb = str(r.font.color.rgb)
                except (AttributeError, ValueError, TypeError):
                    continue
                out.append((rgb, r.font.size.pt if r.font.size else 18.0, r.text,
                            local_bg or page_bg))
    return out


def check(path):
    prs = Presentation(path)
    W, H = inch(prs.slide_width), inch(prs.slide_height)
    problems, warnings = [], []
    slide_count = len(prs.slides)

    for si, slide in enumerate(prs.slides, start=1):
        shapes = []
        sizes = []
        fills = filled_rects(slide)
        page_bg = slide_bg(slide)
        for sh in slide.shapes:
            try:
                x, y = inch(sh.left), inch(sh.top)
                w, h = inch(sh.width), inch(sh.height)
            except TypeError:
                continue

            # R1 越界
            if x < -TOL or y < -TOL or x + w > W + TOL or y + h > H + TOL:
                if not (w > W * 0.98 and h > H * 0.98):     # 满幅背景块不算
                    problems.append({
                        "rule": "R1", "slide": si, "level": "error",
                        "msg": f"越界 {w:.2f}x{h:.2f}in @ ({x:.2f},{y:.2f})，画布 {W:.2f}x{H:.2f}",
                    })

            txt = text_of(sh)
            # 逐个文本容器算：文本框中只有一个，表格按单元格（几何、底色各算各的）
            for tf, cell_bg, box in frames(sh):
                ftxt = "\n".join(p.text for p in tf.paragraphs)
                if not ftxt.strip():
                    continue
                fx, fy, fw, fh = box or (x, y, w, h)
                fsz = next((r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size),
                           18.0)
                sizes.append((fsz, ftxt))
                # R2 文字溢出（估）
                lines = est_lines(ftxt, fw, fsz)
                need = lines * fsz * line_factor(fsz) / 72.0            # 见 line_factor()
                if need > fh + 0.15:
                    problems.append({
                        "rule": "R2", "slide": si, "level": "error",
                        "msg": f"文字可能溢出：{fsz:.1f}pt × 约{lines}行 需 {need:.2f}in，框高 {fh:.2f}in｜{ftxt.split(chr(10))[0][:22]}",
                    })
                shapes.append(text_rect(tf, fx, fy, fw, fh, ftxt, fsz))

            # R3 主题蓝残留
            xml = sh._element.xml
            if THEME_BLUE in xml.upper():
                warnings.append({
                    "rule": "R3", "slide": si, "level": "warn",
                    "msg": "出现 Office 默认描边 2F5496（多半是漏写 <a:ln><a:noFill/></a:ln>）",
                })

            # R6 中文 run 缺 a:ea（中文会掉回默认字体，一行里两种字形）
            if CJK.search(txt) and not any(
                r._r.find(qn("a:rPr")) is not None and r._r.find(qn("a:rPr")).find(qn("a:ea")) is not None
                for tf, _, _ in frames(sh) for p in tf.paragraphs for r in p.runs
            ):
                warnings.append({
                    "rule": "R6", "slide": si, "level": "warn",
                    "msg": f"中文缺 a:ea，会掉字体｜{txt.split(chr(10))[0][:20]}",
                })

            # R7 中文段落用两端对齐：实测 just 不拉伸，纯属白设，建议 l 或 dist
            if CJK.search(txt) and 'algn="just"' in sh._element.xml:
                warnings.append({
                    "rule": "R7", "slide": si, "level": "warn",
                    "msg": f"中文段落 algn=\"just\" 实测不拉伸，改用 l（左对齐）或 dist（分散对齐）｜{txt.split(chr(10))[0][:16]}",
                })

            # R8 文字色 vs 底色：亮强调色（酸绿/柠檬黄）当正文会彻底消失
            local_bg = bg_under(x, y, w, h, fills, default=page_bg)
            for rgb, pt, run_text, run_bg in run_colors(sh, local_bg):
                bg = run_bg
                c = contrast(rgb, bg)
                if c < 3.0:
                    problems.append({
                        "rule": "R8", "slide": si, "level": "error",
                        "msg": f"文字 #{rgb} 在底色 #{bg} 上对比度 {c:.2f} < 3.0｜{run_text[:16]}",
                    })
                elif c < 4.5 and pt < 18:
                    warnings.append({
                        "rule": "R8", "slide": si, "level": "warn",
                        "msg": f"小字 #{rgb}/{bg} 对比度 {c:.2f} < 4.5，投影仪上会糊｜{run_text[:16]}",
                    })

        # R4 文本框重叠
        for i in range(len(shapes)):
            for j in range(i + 1, len(shapes)):
                ax, ay, aw, ah, at = shapes[i]
                bx, by, bw, bh, bt = shapes[j]
                ox = max(0, min(ax + aw, bx + bw) - max(ax, bx))
                oy = max(0, min(ay + ah, by + bh) - max(ay, by))
                if ox <= 0 or oy <= 0:
                    continue
                inter = ox * oy
                small = min(aw * ah, bw * bh)
                if small > 0 and inter / small > 0.40:
                    problems.append({
                        "rule": "R4", "slide": si, "level": "error",
                        "msg": f"文本框重叠 {inter/small:.0%}：{at[:16]} × {bt[:16]}",
                    })

        # R5 字号对比度
        if len(sizes) >= 2:
            big = max(s for s, _ in sizes)
            body = sorted(s for s, _ in sizes)[len(sizes) // 3]
            if body > 0 and big / body < 3.0:
                warnings.append({
                    "rule": "R5", "slide": si, "level": "warn",
                    "msg": f"字号对比度 {big:.0f}/{body:.0f} = {big/body:.2f} < 3.0，容易像文档而不是设计稿",
                })

    return slide_count, problems, warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--json")
    args = ap.parse_args()

    if not os.path.exists(args.path):
        print(f"找不到 {args.path}", file=sys.stderr)
        return 2

    n, problems, warnings = check(args.path)
    print(f"{os.path.basename(args.path)}  {n} 页")
    for p in problems:
        print(f"  [error] {p.get('slide','-')}  {p['rule']}  {p['msg']}")
    for w in warnings:
        print(f"  [warn ] {w.get('slide','-')}  {w['rule']}  {w['msg']}")
    if not problems and not warnings:
        print("  干净：没有越界、重叠、溢出或对比度问题")

    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump({"slides": n, "problems": problems, "warnings": warnings}, f, ensure_ascii=False, indent=2)

    print()
    print(f"合计 error {len(problems)} / warn {len(warnings)}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
