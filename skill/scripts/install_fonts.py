#!/usr/bin/env python3
"""安装 / 检查 skill 自带的 6 个 OFL 字体（Windows）。

为什么有这个脚本：`install_fonts.ps1` 做的是同一件事，但不少机器的 PowerShell
执行策略是 AllSigned（组策略下发），`-ExecutionPolicy Bypass` 也覆盖不了，
`.ps1` 直接跑不起来。字体安装本质就是「拷文件 + 写注册表 + 通知 GDI」，
用 Python 做同样三件事，不受脚本策略影响。

    python install_fonts.py --check      # 只检查缺哪些（退出码 1 = 有缺）
    python install_fonts.py --list       # 列主题里用到的字体族
    python install_fonts.py              # 安装（管理员→全机；否则→仅当前用户）
    python install_fonts.py --dry-run    # 只说要做什么

装到哪：
    管理员 → %WINDIR%\\Fonts                       + HKLM\\...\\Fonts
    普通   → %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts + HKCU\\...\\Fonts
             （按用户安装要重开 PowerPoint 才认；实在不行注销重登）

注意：pptx 里只存字体名、不存字形。目标机器没这个字体，PowerPoint 会静默替换，
字宽一变版面就走形。所以正文一律用系统字体，只有 display 层可以用这些自带字体。
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
FONT_DIR = SKILL / "assets" / "fonts"
THEMES = SKILL / "assets" / "themes.json"

# 文件 → (族名, 用途, 注册表值名的后缀)
FACES = [
    ("InstrumentSerif-Regular.ttf", "Instrument Serif", "display 衬线", " (TrueType)"),
    ("InstrumentSerif-Italic.ttf", "Instrument Serif", "display 斜体", " Italic (TrueType)"),
    ("BebasNeue-Regular.ttf", "Bebas Neue", "poster 窄体", " (TrueType)"),
    ("Barlow-Regular.ttf", "Barlow", "body 无衬线", " (TrueType)"),
    ("BarlowCondensed-Regular.ttf", "Barlow Condensed", "condensed 无衬线", " (TrueType)"),
    ("IBMPlexMono-Regular.ttf", "IBM Plex Mono", "label 等宽", " (TrueType)"),
]

FONTS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def _is_windows() -> bool:
    return os.name == "nt"


def _registry_families() -> set[str]:
    """从 HKLM + HKCU 的 Fonts 键读出已注册的族名（值名形如 "Barlow (TrueType)"）。"""
    fams: set[str] = set()
    if not _is_windows():
        return fams
    import winreg

    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(root, FONTS_KEY) as k:
                n = winreg.QueryInfoKey(k)[1]
                for i in range(n):
                    name, value, _ = winreg.EnumValue(k, i)
                    fams.add(name.split("(")[0].strip())
                    # 值里也可能带文件名，兜底按文件名识别
                    fams.add(str(value))
        except OSError:
            continue
    return fams


def _installed_families() -> set[str]:
    """优先用 GDI 的字体列表（最准）；拿不到就退回注册表。"""
    if _is_windows():
        try:
            import ctypes.wintypes

            LF_FACESIZE = 32

            class LOGFONTW(ctypes.Structure):
                _fields_ = [("lfHeight", ctypes.c_long), ("lfWidth", ctypes.c_long),
                            ("lfEscapement", ctypes.c_long), ("lfOrientation", ctypes.c_long),
                            ("lfWeight", ctypes.c_long), ("lfItalic", ctypes.c_byte),
                            ("lfUnderline", ctypes.c_byte), ("lfStrikeOut", ctypes.c_byte),
                            ("lfCharSet", ctypes.c_byte), ("lfOutPrecision", ctypes.c_byte),
                            ("lfClipPrecision", ctypes.c_byte), ("lfQuality", ctypes.c_byte),
                            ("lfPitchAndFamily", ctypes.c_byte),
                            ("lfFaceName", ctypes.c_wchar * LF_FACESIZE)]

            gdi = ctypes.WinDLL("gdi32")
            user = ctypes.WinDLL("user32")
            hdc = user.GetDC(None)
            fams: set[str] = set()
            FONTENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW),
                                              ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p)

            def cb(lf, _tm, _type, _lparam):
                fams.add(lf.contents.lfFaceName)
                return 1

            lf = LOGFONTW()
            lf.lfCharSet = 1  # DEFAULT_CHARSET
            gdi.EnumFontFamiliesExW(hdc, ctypes.byref(lf), FONTENUMPROC(cb), 0, 0)
            user.ReleaseDC(None, hdc)
            if fams:
                return fams
        except Exception:
            pass
    return _registry_families()


def cmd_check() -> int:
    have = _installed_families()
    print(f"已安装字体族: {len(have)} 个")
    missing = []
    for fname, family, role, _ in FACES:
        ok = family in have
        print(f"  [{'OK  ' if ok else 'MISS'}] {family:<20} {role}")
        if not ok:
            missing.append(fname)
    if missing:
        print(f"\n缺 {len(missing)} 个。运行 `python tasks.py fonts`（或本脚本不带参数）安装。")
        return 1
    print("\n全部已安装。")
    return 0


def cmd_list() -> int:
    if not THEMES.exists():
        print(f"找不到 {THEMES}", file=sys.stderr)
        return 2
    doc = json.loads(THEMES.read_text(encoding="utf-8"))
    for t in doc["themes"]:
        if not t.get("ppt_safe"):
            continue
        f = t["fonts"]
        print(f"{t['slug']:<20} display={f['display_latin']} / {f['display_ea']}   "
              f"body={f['body_latin']} / {f['body_ea']}   mono={f['mono']}")
    return 0


def cmd_install(dry_run: bool) -> int:
    if not _is_windows():
        print("这个脚本只做 Windows 安装；macOS 把 ttf 拖进「字体册」即可。", file=sys.stderr)
        return 2
    if not FONT_DIR.is_dir():
        print(f"找不到字体目录: {FONT_DIR}", file=sys.stderr)
        return 2

    admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
    if admin:
        target = Path(os.environ["WINDIR"]) / "Fonts"
        import winreg
        key_root, key_label = winreg.HKEY_LOCAL_MACHINE, "HKLM"
    else:
        target = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
        import winreg
        key_root, key_label = winreg.HKEY_CURRENT_USER, "HKCU"

    print(f"{'管理员' if admin else '普通权限'} → {target}  ({key_label})")
    if not admin:
        print("  提示：按用户安装要重启 PowerPoint 才认；不行就注销重登。")
    if dry_run:
        print("  --dry-run：不写任何东西")
        for fname, family, _, _suffix in FACES:
            print(f"  会装 {family:<20} ← {fname}")
        return 0

    target.mkdir(parents=True, exist_ok=True)
    with winreg.CreateKeyEx(key_root, FONTS_KEY, 0, winreg.KEY_SET_VALUE) as k:
        for fname, family, role, suffix in FACES:
            src = FONT_DIR / fname
            if not src.exists():
                print(f"  跳过（skill 里没有）: {fname}")
                continue
            dst = target / fname
            if not dst.exists() or dst.stat().st_size != src.stat().st_size:
                shutil.copy2(src, dst)
            # 值名必须是 "Family (TrueType)"，值是文件名
            winreg.SetValueEx(k, family + suffix, 0, winreg.REG_SZ, fname)
            ctypes.windll.gdi32.AddFontResourceW(str(dst))
            print(f"  装: {family}  ({role})")

    # WM_FONTCHANGE 广播，让已开着的进程刷新字体表
    HWND_BROADCAST, WM_FONTCHANGE, SMTO_ABORTIFHUNG = 0xFFFF, 0x001D, 0x0002
    result = ctypes.c_size_t()
    ctypes.windll.user32.SendMessageTimeoutW(
        HWND_BROADCAST, WM_FONTCHANGE, 0, 0, SMTO_ABORTIFHUNG, 1500, ctypes.byref(result))

    print("\n--- 复查 ---")
    rc = cmd_check()
    if rc != 0:
        print("已写入磁盘和注册表，但当前进程的字体缓存还没刷新。")
        print("  1) 重开 PowerPoint，再跑 --check   2) 注销重登   3) 用管理员重跑（立即生效）")
        print("在此期间生成 pptx 仍然正确，只是预览会用替代字体渲染。")
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="安装/检查 skill 自带 OFL 字体")
    ap.add_argument("--check", action="store_true", help="只检查缺哪些")
    ap.add_argument("--list", action="store_true", help="列主题用到的字体族")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不写文件/注册表")
    a = ap.parse_args(argv)
    if a.check:
        return cmd_check()
    if a.list:
        return cmd_list()
    return cmd_install(a.dry_run)


if __name__ == "__main__":
    raise SystemExit(main())
