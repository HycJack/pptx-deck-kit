# -*- coding: utf-8 -*-
"""
通用后处理：给任意 pptx 的每一页加切换 + 入场动画。

它是「生成器做不到」的逃生口——不管用什么库生成，产物都是标准 pptx，
只要在 zip 里的 slideN.xml 追加 <p:transition> 与 <p:timing> 就能动起来。

  python tools/add_anim.py in.pptx out.pptx --effect fade --stagger 160

动画顺序按「从上到下、从左到右」排，符合阅读顺序；整页背景块（满幅色块）
默认跳过，否则一开场就是整页发光。
"""
import argparse
import os
import sys

from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_anim import Animator  # noqa: E402

EMU_IN = 914400
FULLBLEED_W = 9.4      # in，超过视为整页色块
FULLBLEED_H = 5.0


def candidates(slide, skip_fullbleed=True, max_n=8):
    out = []
    for sh in slide.shapes:
        try:
            w, h = sh.width / EMU_IN, sh.height / EMU_IN
            t, l = sh.top / EMU_IN, sh.left / EMU_IN
        except TypeError:
            continue
        if w <= 0 or h <= 0:
            continue
        if skip_fullbleed and (w >= FULLBLEED_W or h >= FULLBLEED_H):
            continue
        out.append((round(t, 2), round(l, 2), sh))
    out.sort(key=lambda x: (x[0], x[1]))
    return [sh for _, _, sh in out[:max_n]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--effect", default="fade", choices=["fade", "wipe", "fly", "appear"])
    ap.add_argument("--stagger", type=int, default=160, help="每个元素的延迟 ms")
    ap.add_argument("--duration", type=int, default=450)
    ap.add_argument("--transition", default="fade", choices=["fade", "wipe", "push", "morph", "none"])
    ap.add_argument("--max", type=int, default=8)
    ap.add_argument("--include-fullbleed", action="store_true")
    args = ap.parse_args()

    prs = Presentation(args.src)
    total = 0
    for idx, slide in enumerate(prs.slides, 1):
        a = Animator(slide)
        if args.transition != "none":
            a.transition(args.transition)
        shapes = candidates(slide, not args.include_fullbleed, args.max)
        for i, sh in enumerate(shapes):
            a.entrance(sh, args.effect, delay=i * args.stagger, duration=args.duration)
        a.build()
        total += len(shapes)
        print(f"  slide {idx}: {len(shapes)} 个元素")

    prs.save(args.dst)
    kind = args.transition if args.transition != "none" else "无切换"
    print(f"✔ {args.dst}  {len(prs.slides.__iter__.__self__._sldIdLst)} 页  "
          f"{total} 个动画  {os.path.getsize(args.dst)/1024:.0f} KB  切换={kind}")


if __name__ == "__main__":
    main()
