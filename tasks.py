#!/usr/bin/env python3
"""统一入口。跨平台，不需要 make。

    python tasks.py list                                    # 13 个版式
    python tasks.py preview --theme studio --out out.pptx    # 排一份完整 deck
    python tasks.py preview --only kpi-grid,table            # 只排两页
    python tasks.py sheet                                    # 14 套主题样张
    python tasks.py lint out.pptx                            # 8 条规则自检
    python tasks.py render out.pptx --dpi 110                # 可选：rkit → com → soffice
    python tasks.py fonts --check                            # 查自带字体装没装

`render` 只是「肉眼看一眼」的补充：lint / preview / verify 都不需要渲染器。
三条路：libreoffice-kit 的 cli.js（自带 soffice）、PowerPoint COM（Windows + Office）、
裸 soffice（只出第一页）。cli.js 路径因机器而异，所以从环境变量取：

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


def _render_rkit(pptx: str, out_dir: str, dpi: str) -> int:
    """LibreOffice 路线：libreoffice-kit 的 cli.js（自带 soffice，机器上不用装 LO）。"""
    cli = os.environ.get("RENDER_CLI", "")
    node = shutil.which("node")
    if not cli or not Path(cli).exists():
        return -1
    if not node:
        print("PATH 里找不到 node", file=sys.stderr)
        return 2
    rc = subprocess.call([node, cli, "render", "--input", pptx,
                          "--output-dir", out_dir, "--dpi", dpi])
    pngs = sorted(Path(out_dir).glob("page-*.png"))
    if pngs:
        # cli.js 收尾时会打一行 "Image batch cleanup failed" 并返回非 0，图其实是全的
        print(f"{len(pngs)} 张 PNG → {out_dir}（LibreOffice）")
        return 0
    return rc or 1


_COM_VBS = r'''Option Explicit
' Export one PNG per slide with PowerPoint COM.
' Keep this file pure ASCII: cscript reads .vbs as ANSI, UTF-8 Chinese comments
' make it fail with "invalid character" at (1,1).
Dim app, pres, i, outDir, w, h
If WScript.Arguments.Count < 4 Then WScript.Quit 2
Set app = CreateObject("PowerPoint.Application")
Set pres = app.Presentations.Open(WScript.Arguments(0), True, False, False)
outDir = WScript.Arguments(1)
w = CLng(WScript.Arguments(2))
h = CLng(WScript.Arguments(3))
For i = 1 To pres.Slides.Count
  pres.Slides(i).Export outDir & "\page-" & Right("0000" & i, 4) & ".png", "PNG", w, h
Next
pres.Close
app.Quit
'''


def _render_com(pptx: str, out_dir: str, dpi: str) -> int:
    """PowerPoint COM 路线（Windows + 装了 Office）。不依赖 PowerShell 执行策略。

    用 cscript + 临时 .vbs：`.ps1` 在组策略 AllSigned 的机器上跑不起来（本仓库踩过），
    而 VBScript 不受那条策略管。
    """
    if os.name != "nt":
        return -1
    cscript = str(Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "cscript.exe")
    if not Path(cscript).exists():
        return -1
    import tempfile

    from pptx import Presentation

    try:
        prs = Presentation(pptx)
        w_in = prs.slide_width / 914400
        h_in = prs.slide_height / 914400
    except Exception as e:                                    # noqa: BLE001
        print(f"读不了 {pptx}: {e}", file=sys.stderr)
        return 2
    w_px, h_px = round(w_in * float(dpi)), round(h_in * float(dpi))

    with tempfile.TemporaryDirectory(prefix="pptx-com-") as td:
        vbs = Path(td) / "export.vbs"
        vbs.write_text(_COM_VBS, encoding="ascii")
        r = subprocess.run([cscript, "//nologo", str(vbs), str(Path(pptx).resolve()),
                            str(Path(out_dir).resolve()), str(w_px), str(h_px)],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
    pngs = sorted(Path(out_dir).glob("page-*.png"))
    if pngs:
        print(f"{len(pngs)} 张 PNG → {out_dir}（PowerPoint COM，{w_px}×{h_px}）")
        return 0
    if r.stderr.strip():
        print(r.stderr.strip()[:400], file=sys.stderr)
    print("PowerPoint COM 导出失败：确认装了 Office、且 pptx 路径没有中文/空格问题。",
          file=sys.stderr)
    return r.returncode or 1


def _render_soffice(pptx: str, out_dir: str) -> int:
    """裸 soffice 路线：只出第一页（`--convert-to png` 不支持多页），当兜底用。"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return -1
    import tempfile

    with tempfile.TemporaryDirectory(prefix="pptx-lo-") as td:
        r = subprocess.run([soffice, "--headless", "--convert-to", "png",
                            "--outdir", td, str(Path(pptx).resolve())],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        src = next(Path(td).glob("*.png"), None)
        if not src:
            print(r.stderr.strip()[:300] or "soffice 没出图", file=sys.stderr)
            return 1
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, Path(out_dir) / "page-0001.png")
    print(f"1 张 PNG → {out_dir}（soffice 只出第一页；要看全篇请用 --engine rkit 或 com）")
    return 0


def render(pptx: str, out_dir: str | None, dpi: str, engine: str) -> int:
    """渲染验收：三条路都行，都没有也能干活（只跑 lint）。

    lint / preview / verify 全部不需要渲染器——渲染只是「肉眼看一眼」的补充手段。
    """
    if not Path(pptx).exists():
        print(f"文件不存在: {pptx}", file=sys.stderr)
        return 2
    out_dir = out_dir or (str(Path(pptx).with_suffix("")) + "-render")
    if Path(out_dir).exists():
        print(f"输出目录已存在：{out_dir}（先删掉，避免混进旧图）", file=sys.stderr)
        return 2
    # cli.js 的 mkdir 不递归，父目录不存在会 ENOENT
    Path(out_dir).parent.mkdir(parents=True, exist_ok=True)

    order = {"auto": ["rkit", "com", "soffice"],
             "rkit": ["rkit"], "com": ["com"], "soffice": ["soffice"]}[engine]
    for name in order:
        if name == "rkit":
            rc = _render_rkit(pptx, out_dir, dpi)
        elif name == "com":
            Path(out_dir).mkdir(parents=True, exist_ok=True)
            rc = _render_com(pptx, out_dir, dpi)
        else:
            rc = _render_soffice(pptx, out_dir)
        if rc != -1:
            return rc
    print("找不到可用的渲染器。三条路选一条：", file=sys.stderr)
    print("  1) LibreOffice：装 LO 后 `--engine soffice`；或用 libreoffice-kit 的 cli.js：", file=sys.stderr)
    print('     $env:RENDER_CLI = "...\\libreoffice-kit\\lib\\cli.js"', file=sys.stderr)
    print("  2) Windows + Office：`--engine com`（走 PowerPoint COM，不需要 LO）", file=sys.stderr)
    print("  3) 都不装：跳过渲染，`python tasks.py lint out.pptx` 也能查 8 条规则", file=sys.stderr)
    return 2


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

    sp = sub.add_parser("render", help="渲成 PNG 逐页看（rkit / com / soffice，都不是必须的）")
    sp.add_argument("pptx")
    sp.add_argument("--dpi", default="110")
    sp.add_argument("--out-dir", default=None, help="默认 <pptx 同名>-render，必须不存在")
    sp.add_argument("--engine", default="auto", choices=["auto", "rkit", "com", "soffice"],
                    help="auto=rkit→com→soffice，取第一个可用的")

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
        return render(args.pptx, args.out_dir, args.dpi, args.engine)

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
