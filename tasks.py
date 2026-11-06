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


def _theme_slugs() -> list[str]:
    import json

    doc = json.loads((ROOT / "skill" / "assets" / "themes.json").read_text(encoding="utf-8"))
    return [t["slug"] for t in doc["themes"] if t.get("ppt_safe")]


def verify(only: str, out_dir: str, keep: bool) -> int:
    """14 套主题 × 13 个版式全排一遍，逐份 lint。

    一套主题干净不代表都干净：版式里用到强调色的地方（评级列、色条、发丝线）
    在浅纸和深纸上表现完全不同，grove 的砖红砸在斑马行上就掉到 2.85。
    """
    import tempfile

    slugs = [s for s in only.split(",") if s] or _theme_slugs()
    tmp = Path(out_dir) if out_dir else Path(tempfile.mkdtemp(prefix="deckkit-verify-"))
    tmp.mkdir(parents=True, exist_ok=True)
    bad: list[str] = []
    for slug in slugs:
        pptx = tmp / f"{slug}.pptx"
        rc = call("preview_layouts.py", str(pptx), "--theme", slug)
        if rc != 0:
            bad.append(slug)
            print(f"{slug:<20} 排版失败")
            continue
        log = subprocess.run([sys.executable, str(SCRIPTS / "lint_pptx.py"), str(pptx)],
                             capture_output=True, text=True, encoding="utf-8", errors="replace")
        last = [ln for ln in log.stdout.splitlines() if ln.strip()]
        summary = last[-1].strip() if last else "(无输出)"
        ok = log.returncode == 0 and "error 0" in summary
        print(f"{slug:<20} {summary}")
        if not ok:
            bad.append(slug)
            for ln in last[:-1]:
                print("   " + ln.strip())
    print()
    if bad:
        print(f"失败 {len(bad)} 套：{', '.join(bad)}")
        print(f"产物留在 {tmp}")
        return 1
    print(f"{len(slugs)} 套主题 × 13 个版式全部 lint 干净。")
    if not keep and not out_dir:
        import shutil as _sh

        _sh.rmtree(tmp, ignore_errors=True)
    return 0


def genppt(action: str, patch: str, verify: bool) -> int:
    """内嵌的 genppt = 上游 61df963 + skill/scripts/genppt-effects.patch。

    上游升级时：先换 vendor/genppt 成新的原始树，再 `python tasks.py genppt --apply`；
    补丁打不上（context 冲突）就得重新裁一遍，别手改 vendor——手改就再也传不回上游了。
    """
    patch_path = Path(patch) if patch else SCRIPTS / "genppt-effects.patch"
    if not patch_path.exists():
        print(f"找不到补丁: {patch_path}", file=sys.stderr)
        return 2

    files = [ln.split(" b/", 1)[1] for ln in patch_path.read_text(encoding="utf-8").splitlines()
             if ln.startswith("diff --git ")]
    base = ["git", "apply", "-p1", "--directory=vendor/genppt"]

    def gitapply(extra: list[str], quiet: bool = False) -> int:
        r = subprocess.run([*base, *extra, str(patch_path)], cwd=ROOT,
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode != 0 and r.stderr.strip() and not quiet:
            print(r.stderr.strip(), file=sys.stderr)
        return r.returncode

    applied = gitapply(["--check", "-R"], quiet=True) == 0
    can_apply = gitapply(["--check"], quiet=True) == 0

    if action == "check":
        print(f"补丁: {patch_path.name}  ({len(files)} 个文件)")
        for f in files:
            print(f"  {f}")
        state = "已应用" if applied else ("可应用" if can_apply else "打不上（上下文冲突，需要重裁）")
        print(f"vendor/genppt 状态: {state}")
        if verify:
            return _go_all()
        return 0 if applied else 1

    if action == "apply":
        if applied:
            print("已经是打过补丁的状态，无事可做。")
        elif not can_apply:
            print("补丁打不上：vendor/genppt 可能不是 61df963 的原始树。", file=sys.stderr)
            return 1
        else:
            rc = gitapply([])
            if rc != 0:
                return rc
            print(f"已应用 {patch_path.name}（{len(files)} 个文件）")
    elif action == "revert":
        if not applied:
            print("当前没有应用补丁，无事可做。")
        else:
            rc = gitapply(["-R"])
            if rc != 0:
                return rc
            print("已回退补丁（vendor/genppt 回到上游原始状态）")

    if verify:
        return _go_all()
    return 0


def _go_all() -> int:
    for target in ("build", "vet", "test"):
        rc = _go(target)
        if rc != 0:
            return rc
    return 0


def _go(target: str) -> int:
    """在 vendor/genppt 里跑 go build / vet / test。"""
    import shutil as _sh

    go = _sh.which("go")
    if not go:
        print("PATH 里没有 go，跳过 Go 校验", file=sys.stderr)
        return 0
    print(f"$ go {target} ./...  (vendor/genppt)")
    r = subprocess.run([go, target, "./..."], cwd=ROOT / "vendor" / "genppt")
    return r.returncode


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

    sp = sub.add_parser("verify", help="14 套主题各排一份 + lint，全绿才算过（CI 用）")
    sp.add_argument("--themes", default="", help="只验证这些 slug，逗号分隔")
    sp.add_argument("--out-dir", default="", help="产物目录，默认临时目录")
    sp.add_argument("--keep", action="store_true", help="保留产物")

    sp = sub.add_parser("genppt", help="内嵌 genppt 的补丁管理（check/apply/revert）")
    sp.add_argument("action", nargs="?", default="check",
                    choices=["check", "apply", "revert"], help="默认 check")
    sp.add_argument("--patch", default="", help="补丁路径，默认 skill/scripts/genppt-effects.patch")
    sp.add_argument("--go", action="store_true", help="顺带跑 go build/vet/test")

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
        # cli.js 的 mkdir 不递归：父目录不存在会 ENOENT
        Path(out_dir).parent.mkdir(parents=True, exist_ok=True)
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

    if args.cmd == "verify":
        return verify(args.themes, args.out_dir, args.keep)

    if args.cmd == "genppt":
        return genppt(args.action, args.patch, args.go)

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
