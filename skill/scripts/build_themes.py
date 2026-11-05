#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 refs 模板派生 pptx 主题注册表。

上游 template.json 是给 HTML 用的：颜色键名自由（neon_pink/bg_alt/wood）、
字号是 px、字体是 web font，还有 blend-mode / 纹理这类 pptx 做不到的玩法。
本脚本做四件事：
  1. 把自由键名映射成 pptx 的固定角色 paper/ink/accent/accent2/rule
  2. 校验：这里声称的每个 hex 必须真的出现在上游 palette 里（防止我编）
  3. 算 WCAG 对比度，据此决定 accent 是「当字」还是「当块」
  4. 配字体：本机字体 + skill 自带字体 + 缺失时的降级链

用法: python build_themes.py
输出: ../assets/themes.json
"""
from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
OUT = os.path.join(SKILL, "assets", "themes.json")
# 只在「重新生成注册表」时才需要：上游 HTML 模板库
#   git clone https://github.com/zarazhangrui/beautiful-html-templates
#   export THEMES_REFS=.../beautiful-html-templates/templates
# assets/themes.json 已经提交在仓库里，平时不用跑这个脚本。
# 派生基线：beautiful-html-templates @ e5e204f
REFS = os.environ.get("THEMES_REFS", "refs/beautiful-html-templates/templates")


def lum(h: str) -> float:
    h = h.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def cr(a: str, b: str) -> float:
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return round((hi + 0.05) / (lo + 0.05), 2)


# ---------------------------------------------------------------- 字体预设
# 每个预设给出 fonts（目标）+ fonts_fallback（目标机器缺字体时用）。
# bundled 列出哪些族是 skill 自带（assets/fonts，OFL 授权）。中文一律用系统字体：
# 宋体系（宋体/华文宋体）配衬线主题，黑体系（微软雅黑/等线 Light）配无衬线主题。
# 规则：高对比衬线只做 display —— Instrument Serif 没有粗体，15pt 正文会糊。
FONTS = {
    "serif-editorial": dict(
        fonts=dict(display_latin="Instrument Serif", display_ea="华文宋体",
                   body_latin="Georgia", body_ea="宋体", mono="IBM Plex Mono"),
        fallback=dict(display_latin="Georgia", display_ea="华文宋体",
                      body_latin="Georgia", body_ea="宋体", mono="Consolas"),
        bundled=["Instrument Serif", "IBM Plex Mono"],
    ),
    "serif-warm": dict(
        fonts=dict(display_latin="Instrument Serif", display_ea="华文宋体",
                   body_latin="Constantia", body_ea="宋体", mono="IBM Plex Mono"),
        fallback=dict(display_latin="Constantia", display_ea="华文宋体",
                      body_latin="Constantia", body_ea="宋体", mono="Consolas"),
        bundled=["Instrument Serif", "IBM Plex Mono"],
    ),
    "poster-condensed": dict(
        fonts=dict(display_latin="Bebas Neue", display_ea="微软雅黑",
                   body_latin="Barlow", body_ea="微软雅黑", mono="IBM Plex Mono"),
        fallback=dict(display_latin="Arial Narrow", display_ea="微软雅黑",
                      body_latin="Segoe UI", body_ea="微软雅黑", mono="Consolas"),
        bundled=["Bebas Neue", "Barlow", "IBM Plex Mono"],
    ),
    "grotesque": dict(
        fonts=dict(display_latin="Barlow", display_ea="微软雅黑",
                   body_latin="Barlow", body_ea="等线 Light", mono="IBM Plex Mono"),
        fallback=dict(display_latin="Segoe UI Semibold", display_ea="微软雅黑",
                      body_latin="Segoe UI", body_ea="等线 Light", mono="Consolas"),
        bundled=["Barlow", "IBM Plex Mono"],
    ),
    "condensed": dict(
        fonts=dict(display_latin="Barlow Condensed", display_ea="微软雅黑",
                   body_latin="Barlow", body_ea="等线 Light", mono="IBM Plex Mono"),
        fallback=dict(display_latin="Arial Narrow", display_ea="微软雅黑",
                      body_latin="Segoe UI", body_ea="等线 Light", mono="Consolas"),
        bundled=["Barlow Condensed", "Barlow", "IBM Plex Mono"],
    ),
}

# ---------------------------------------------------------------- 手工策展
# 角色分配是判断题，不能自动猜：paper 不一定是亮色（dark 主题）、
# accent 不一定是彩色（monochrome）、ink 不一定接近黑（cobalt 是电光蓝）。
# 每个 hex 都来自上游 palette；只有 rule（发丝线色，上游通常没有）允许派生。
C = [
    dict(
        slug="editorial-forest", name_zh="纸与墨", register="editorial",
        roles=dict(paper="#efe7d4", ink="#1a1a17", accent="#2e4a2a",
                   accent2="#e89cb1", rule="#d8d0be"),
        derived=["rule"], accent_role="text", font="serif-editorial",
        type=dict(display_pt=54, display_em=0.92, body_em=1.45,
                  tracking_display_em=-0.02, tracking_label_em=0.16),
        note="季度回顾、内部汇报、工作室更新。安静，不催人。",
    ),
    dict(
        slug="cobalt-grid", name_zh="电光蓝", register="technical",
        roles=dict(paper="#F0EBDE", ink="#1F2BE0", accent="#5560E5",
                   accent2="#1A1A17", rule="#C9C2B2"),
        derived=["accent2", "rule"], accent_role="text", font="grotesque",
        type=dict(display_pt=54, display_em=0.98, body_em=1.5,
                  tracking_display_em=-0.01, tracking_label_em=0.16),
        note="技术方案、架构评审。全身蓝色文字是这套的性格，别加第二个彩色。",
    ),
    dict(
        slug="signal", name_zh="夜航金", register="formal-dark",
        roles=dict(paper="#1c2644", ink="#e2dcd0", accent="#c8a870",
                   accent2="#232f55", rule="#3A4568"),
        derived=["rule"], accent_role="text", font="serif-editorial",
        type=dict(display_pt=54, display_em=0.95, body_em=1.45,
                  tracking_display_em=-0.02, tracking_label_em=0.18),
        note="年报、董事会、对外正式发布。深底 + 金，庄重但不冷漠。",
    ),
    dict(
        slug="neo-grid-bold", name_zh="酸绿网格", register="technical",
        roles=dict(paper="#ECECE8", ink="#0A0A0A", accent="#E6FF3D",
                   accent2="#8A8A85", rule="#CFCFC8"),
        derived=["rule"], accent_role="block", font="condensed",
        type=dict(display_pt=64, display_em=0.9, body_em=1.45,
                  tracking_display_em=-0.03, tracking_label_em=0.18),
        note="产品发布、建筑/工业题材。酸绿只能做色块和横条，做字会消失。",
    ),
    dict(
        slug="bold-poster", name_zh="黑红宣言", register="poster",
        roles=dict(paper="#F5F2EF", ink="#1C1410", accent="#D8000F",
                   accent2="#8A8580", rule="#D8D4D0"),
        derived=["accent2", "rule"], accent_role="text", font="poster-condensed",
        type=dict(display_pt=72, display_em=0.88, body_em=1.4,
                  tracking_display_em=-0.03, tracking_label_em=0.2),
        note="宣言、促销、对外主张。字号要大到有点过分才对。",
    ),
    dict(
        slug="monochrome", name_zh="素纸", register="minimal",
        roles=dict(paper="#fafadf", ink="#1a1a16", accent="#5e5e54",
                   accent2="#D8D8B8", rule="#C4C4A4"),
        derived=["accent2", "rule"], accent_role="text", font="serif-warm",
        type=dict(display_pt=54, display_em=0.95, body_em=1.55,
                  tracking_display_em=-0.01, tracking_label_em=0.18),
        note="文字为主的稿子、讲稿、内部备忘录。完全没有彩色，靠字号和留白。",
    ),
    dict(
        slug="biennale-yellow", name_zh="双年展黄", register="cultural",
        roles=dict(paper="#E9E5DB", ink="#1B2566", accent="#F1EE2E",
                   accent2="#E26B4A", rule="#C8C2B2"),
        derived=["rule"], accent_role="block", font="serif-warm",
        type=dict(display_pt=60, display_em=0.92, body_em=1.45,
                  tracking_display_em=-0.02, tracking_label_em=0.16),
        note="展览、文化活动、品牌提案。亮黄必须配深蓝字，白字会瞎。",
    ),
    dict(
        slug="mat", name_zh="苔绿琥珀", register="editorial-dark",
        roles=dict(paper="#232e26", ink="#f0e8d2", accent="#c07030",
                   accent2="#ede6d0", rule="#3A4A3C"),
        derived=["rule"], accent_role="text", font="grotesque",
        type=dict(display_pt=54, display_em=0.95, body_em=1.5,
                  tracking_display_em=-0.02, tracking_label_em=0.18),
        note="材质、工艺、食物、慢生活。暗绿底配琥珀，像纸垫上的样品。",
    ),
    dict(
        slug="long-table", name_zh="长桌赭红", register="cultural",
        roles=dict(paper="#FAF1E2", ink="#8E2D1F", accent="#B53D2A",
                   accent2="#6B5B4A", rule="#DCCDB8"),
        derived=["accent2", "rule"], accent_role="text", font="condensed",
        type=dict(display_pt=54, display_em=1.0, body_em=1.5,
                  tracking_display_em=0.0, tracking_label_em=0.16),
        note="餐饮、酒单、菜单式清单。全身赭红文字，像印章。",
    ),
    dict(
        slug="soft-editorial", name_zh="柔和三色", register="editorial",
        roles=dict(paper="#F2EEDF", ink="#2A241B", accent="#E1A4C2",
                   accent2="#B7C7A8", rule="#D5D0C0"),
        derived=["rule"], accent_role="block", font="serif-warm",
        type=dict(display_pt=54, display_em=1.0, body_em=1.55,
                  tracking_display_em=-0.01, tracking_label_em=0.16),
        note="女性向品牌、生活方式、课程介绍。粉+鼠尾草绿当色块，文字用墨色。",
    ),
    dict(
        slug="emerald-editorial", name_zh="祖母绿", register="editorial",
        roles=dict(paper="#F1E9D6", ink="#0F1A5C", accent="#3CD896",
                   accent2="#B8AE95", rule="#D0C8B2"),
        derived=["accent2", "rule"], accent_role="block", font="serif-editorial",
        type=dict(display_pt=56, display_em=0.94, body_em=1.45,
                  tracking_display_em=-0.02, tracking_label_em=0.18),
        note="金融科技、增长报告。深海军蓝字 + 薄荷绿块，数字密度高的页面很扛。",
    ),
    dict(
        slug="studio", name_zh="黑黄车间", register="poster",
        roles=dict(paper="#1c1c1c", ink="#f5d200", accent="#f5d200",
                   accent2="#F5F5F5", rule="#3A3A38"),
        derived=["accent2", "rule"], accent_role="text", font="poster-condensed",
        type=dict(display_pt=72, display_em=0.88, body_em=1.4,
                  tracking_display_em=-0.02, tracking_label_em=0.2),
        note="工作室介绍、硬核产品、演出海报。黑底黄字，全篇只有这一个颜色。",
    ),
    dict(
        slug="grove", name_zh="深林砖红", register="editorial-dark",
        roles=dict(paper="#192b1b", ink="#d4cfbf", accent="#c8524a",
                   accent2="#e8e4d6", rule="#2E4430"),
        derived=["rule"], accent_role="text", font="serif-editorial",
        type=dict(display_pt=54, display_em=0.95, body_em=1.5,
                  tracking_display_em=-0.02, tracking_label_em=0.18),
        note="深度研究、年度总结、纪录片式叙事。暗林绿几乎接近黑。",
    ),
    dict(
        slug="sakura-chroma", name_zh="木色五色", register="cultural",
        roles=dict(paper="#F1E6CB", ink="#3A2516", accent="#E5392A",
                   accent2="#3F8BC4", rule="#D6C7A8"),
        derived=["rule"], accent_role="block", font="condensed",
        accent_block_because="五色块是主题本体，红只做块不做字",
        type=dict(display_pt=54, display_em=0.95, body_em=1.45,
                  tracking_display_em=-0.02, tracking_label_em=0.16),
        notable=["#E54489", "#F09131", "#F0BC2A", "#3D9F47"],
        note="节庆、市集、儿童内容。上游给了 5 个彩色，一页最多挑 2 个。",
    ),
]

# HTML 里能做、pptx 里做不了的，逐个标出来，免得有人照抄
UNSAFE = {
    "block-frame": "粗色边框 + 叠印，需要负间距和 mix-blend-mode",
    "8-bit-orbit": "像素字体 + 发光，pptx 无外发光，需要位图",
    "retro-windows": "Win98 窗口控件，纯 CSS 边框模拟",
    "retro-zine": "网点/撕纸纹理，需位图底",
    "scatterbrain": "随机散落旋转，pptx 无法保证布局稳定",
    "cartesian": "极细网格与坐标轴，线宽低于 0.75pt 打印会消失",
    "capsule": "胶囊圆角 + 多色块，需手绘 SVG",
    "pin-and-paper": "手写体 + 图钉位图",
    "creative-mode": "多色叠印与旋转贴纸",
    "peoples-platform": "粗描边字 + 拼贴",
    "raw-grid": "1px 网格线全铺，pptx 会有对齐残影",
    "stencil-tablet": "模板喷漆断开效果，需要位图字",
    "daisy-days": "圆润贴纸风，依赖位图装饰",
    "blue-professional": "可用但 ERP 蓝 #1E2BFA 与 Office 默认蓝太近，易混",
}


def load_upstream() -> dict:
    out = {}
    if not os.path.isdir(REFS):
        return out
    for d in sorted(os.listdir(REFS)):
        j = os.path.join(REFS, d, "template.json")
        if os.path.exists(j):
            with open(j, encoding="utf-8") as f:
                out[d] = json.load(f)
    return out


def main() -> int:
    up = load_upstream()
    themes, problems = [], []

    for c in C:
        slug = c["slug"]
        upstream = up.get(slug)
        if upstream is None:
            problems.append(f"{slug}: 上游没有这个 slug")
            continue

        # ---- 校验：hex 必须来自上游（rule 允许派生）
        up_hex = {v.upper() for v in (upstream.get("palette") or {}).values()
                  if isinstance(v, str) and v.startswith("#")}
        derived = set(c.get("derived", []))
        for role, hexv in c["roles"].items():
            if role in derived:
                continue
            if hexv.upper() not in up_hex:
                problems.append(f"{slug}.{role} = {hexv} 不在上游 palette 里：{sorted(up_hex)}")

        p, i, a = c["roles"]["paper"], c["roles"]["ink"], c["roles"]["accent"]
        contrast = dict(ink_on_paper=cr(i, p), accent_on_paper=cr(a, p),
                        ink_on_accent=cr(i, a), white_on_accent=cr("#FFFFFF", a))
        if c["accent_role"] == "text" and contrast["accent_on_paper"] < 3.0:
            problems.append(f"{slug}: accent_role=text 但 contrast 只有 "
                            f"{contrast['accent_on_paper']}，应当改成 block")
        if (c["accent_role"] == "block" and contrast["accent_on_paper"] > 3.0
                and not c.get("accent_block_because")):
            problems.append(f"{slug}: accent_role=block 但对比度 {contrast['accent_on_paper']} "
                            f"已够当字，确认是有意为之就写 accent_block_because")

        t = c["type"]
        if t["display_pt"] / 15.0 < 3.5:
            problems.append(f"{slug}: display/body = {t['display_pt']/15.0:.2f} < 3.5")

        fp = FONTS[c["font"]]
        themes.append(dict(
            slug=slug,
            name=upstream.get("name", slug),
            name_zh=c["name_zh"],
            register=c["register"],
            density=upstream.get("density"),
            formality=upstream.get("formality"),
            mood=(upstream.get("mood") or [])[:4],
            palette=c["roles"],
            derived_roles=sorted(derived),
            accent_role=c["accent_role"],
            accent_text_color=("#FFFFFF" if contrast["white_on_accent"] > contrast["ink_on_accent"]
                               else c["roles"]["ink"]),
            contrast=contrast,
            font_preset=c["font"],
            fonts=fp["fonts"],
            fonts_fallback=fp["fallback"],
            fonts_bundled=fp["bundled"],
            fonts_html=upstream.get("typography", {}),
            type=dict(
                display_pt=t["display_pt"], h1_pt=36, lead_pt=24,
                body_pt=15, small_pt=12, label_pt=8.5,
                display_ratio=round(t["display_pt"] / 15.0, 2),
                # 实测：spcPct=100000 的行高 = 1.2 × 字号
                display_lh_pct=round(t["display_em"] / 1.2 * 100000),
                body_lh_pct=round(t["body_em"] / 1.2 * 100000),
                display_lh_em=t["display_em"], body_lh_em=t["body_em"],
                tracking_display_pt=round(t["tracking_display_em"] * t["display_pt"], 2),
                tracking_display_cjk_pt=0.0,
                tracking_label_pt=round(t["tracking_label_em"] * 8.5, 2),
            ),
            usage=dict(best_for=upstream.get("best_for", ""),
                       avoid_for=upstream.get("avoid_for", ""), note=c["note"]),
            notable_accents=c.get("notable", []),
            ppt_safe=True,
            source=f"beautiful-html-templates/templates/{slug}/template.json",
        ))

    for slug, reason in sorted(UNSAFE.items()):
        if slug in {t["slug"] for t in themes} or slug not in up:
            continue
        themes.append(dict(slug=slug, name=up.get(slug, {}).get("name", slug),
                           register="unsafe-for-pptx", ppt_safe=False,
                           ppt_limit=reason,
                           source=f"beautiful-html-templates/templates/{slug}/template.json"))

    doc = dict(
        schema=1,
        provenance=dict(
            upstream="beautiful-html-templates (34 templates)",
            local=REFS,
            adaptation=[
                "自由颜色键名 → 固定角色 paper/ink/accent/accent2/rule",
                "px 字号 → 实测过的 pt 阶梯 54/36/24/15/12/8.5（display 按主题调 54-72）",
                "web font → skill 自带 OFL 字体；缺字体时按 fonts_fallback 降级",
                "lineHeight: em → a:spcPct 千分比（实测 100% = 1.2x 字号）",
                "letterSpacing: em → 磅；中文强制 0（方块字加字距会散）",
                "按对比度判定 accent 只能当块还是可以当字",
            ],
        ),
        roles=["paper", "ink", "accent", "accent2", "rule"],
        type_ladder=dict(h1=36, lead=24, body=15, small=12, label=8.5),
        fonts_bundled=["Instrument Serif", "Bebas Neue", "Barlow",
                       "Barlow Condensed", "IBM Plex Mono"],
        themes=themes,
    )

    if problems:
        print("!! 校验失败：")
        for p in problems:
            print("   -", p)
        return 1

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
        f.write("\n")

    ok = [t for t in themes if t.get("ppt_safe")]
    print(f"写入 {OUT}")
    print(f"  pptx 可用主题 {len(ok)} 个，标记不可用 {len(themes)-len(ok)} 个")
    for t in ok:
        print(f"   {t['slug']:20} {t['name_zh']:6} {t['register']:13} "
              f"{t['font_preset']:16} ink/paper={t['contrast']['ink_on_paper']:5.2f} "
              f"accent={t['accent_role']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
