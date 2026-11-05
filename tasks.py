#!/usr/bin/env python3
"""统一入口。跨平台，不需要 make。

    python tasks.py list                                    # 13 个版式
    python tasks.py preview --theme studio --out out.pptx    # 排一份完整 deck
    python tasks.py preview --only kpi-grid,table            # 只排两页
    python tasks.py sheet                                    # 14 套主题样张
    python tasks.py lint out.pptx                            # 8 条规则自检
    python tasks.py render out.pptx --dpi 110                # 需要 RENDER_CLI
    python tasks.py fonts --check                            # 查自带字体装没装

`render` 依赖 LibreOffice 的 cli.js，路径因机器而异，所以从环境变量取：

    $env:RENDER_CLI = "...\\libreoffice-kit\\lib\\cli.js"
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPTS = ROOT / "skill" / "scripts"


def call(script: str, *args: str) -> int:
    return subprocess.call([sys.executable, str(SCRIPTS / script), *args])


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="tasks.py", description="pptx-deck-kit 统一入口")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list", help="列出 13 个版式及其来源")

    sp = sub.add_parser("preview", help="按版式排一份 deck")
    sp.add_argument("--theme", default="editorial-forest")
    sp.add_argument("--only", default="", help="逗号分隔的 slug 子集")
    sp.add_argument("--out", default="out.pptx")
    sp.add_argument("--size", default="", help="WxH，例如 13.333x7.5（默认即此）")

    sp = sub.add_parser("sheet", help="14 套主题排成样张")
    sp.add_argument("--out", default="theme-sheet.pptx")

    sp = sub.add_parser("lint", help="版面自检")
    sp.add_argument("pptx")

    sp = sub.add_parser("render", help="渲成 PNG 逐页看（需要 RENDER_CLI）")
    sp.add_argument("pptx")
    sp.add_argument("--dpi", default="110")
    sp.add_argument("--out-dir", default=None, help="默认 <pptx 同名>-render，必须不存在")

    sp = sub.add_parser("fonts", help="查/装自带 OFL 字体（Windows）")
    sp.add_argument("--check", action="store_true", help="只查缺哪些")
    sp.add_argument("--list", action="store_true", help="列出主题用到的字体族")

    args = p.parse_args(argv)

    if args.cmd == "list":
        return call("preview_layouts.py", "--list")

    if args.cmd == "preview":
        a = [args.out, "--theme", args.theme]
        if args.only:
            a += ["--only", args.only]
        if args.size:
            a += ["--size", args.size]
        return call("preview_layouts.py", *a)

    if args.cmd == "sheet":
        return call("preview_themes.py", args.out)

    if args.cmd == "lint":
        return call("lint_pptx.py", args.pptx)

    if args.cmd == "render":
        cli = os.environ.get("RENDER_CLI", "")
        node = shutil.which("node")
        if not cli or not Path(cli).exists():
            print("需要环境变量 RENDER_CLI 指向 libreoffice-kit 的 cli.js", file=sys.stderr)
            return 2
        if not node:
            print("PATH 里找不到 node", file=sys.stderr)
            return 2
        out_dir = args.out_dir or (str(Path(args.pptx).with_suffix("")) + "-render")
        if Path(out_dir).exists():
            print(f"输出目录已存在（cli.js 会报 mkdir EEXIST）：{out_dir}", file=sys.stderr)
            return 2
        rc = subprocess.call([node, cli, "render", "--input", args.pptx,
                              "--output-dir", out_dir, "--dpi", args.dpi])
        pngs = sorted(Path(out_dir).glob("page-*.png"))
        if pngs:
            # cli.js 收尾时会打一行 "Image batch cleanup failed" 并返回非 0，图其实是全的
            print(f"{len(pngs)} 张 PNG → {out_dir}")
            return 0
        return rc or 1

    if args.cmd == "fonts":
        a = []
        if args.check:
            a.append("--check")
        if args.list:
            a.append("--list")
        return call("install_fonts.py", *a)

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
