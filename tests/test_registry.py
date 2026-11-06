"""主题注册表的不变量。

`themes.json` 是可以随便改的数据文件——这些断言是它的"编译期检查"：
改了配色/字号，只要破坏下面任何一条，CI 就红，而不是等做完 13 页才发现字看不见。
"""

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skill" / "scripts"))

from layouts import Theme                            # noqa: E402
from pptx_kit import contrast                        # noqa: E402

FONTS = ROOT / "skill" / "assets" / "fonts"
DOC = json.loads((ROOT / "skill" / "assets" / "themes.json").read_text(encoding="utf-8"))
SAFE = [t for t in DOC["themes"] if t.get("ppt_safe")]
BUNDLED = {p.stem.split("-")[0].lower() for p in FONTS.glob("*.ttf")}
# 非分发的拉丁字体：靠系统字体兜底，限这几个（Windows / Office 都自带）
SYSTEM_LATIN = {"georgia", "constantia"}


def themes():
    return [(t["slug"], Theme(t)) for t in SAFE]


class Registry(unittest.TestCase):
    def test_14_usable_themes_with_unique_slugs(self):
        self.assertEqual(len(SAFE), 14, "ppt_safe 主题数变了，文档与预览图要一起改")
        slugs = [t["slug"] for t in SAFE]
        self.assertEqual(len(set(slugs)), len(slugs), f"slug 重复：{slugs}")

    def test_required_fields(self):
        for t in SAFE:
            with self.subTest(theme=t["slug"]):
                for k in ("slug", "palette", "fonts", "type"):
                    self.assertIn(k, t)
                for k in ("paper", "ink", "accent", "accent2", "rule"):
                    self.assertIn(k, t["palette"])
                for k in ("display_latin", "body_latin", "display_ea", "body_ea", "mono"):
                    self.assertIn(k, t["fonts"])
                for k in ("display_pt", "h1_pt", "lead_pt", "body_pt", "small_pt", "label_pt"):
                    self.assertIn(k, t["type"])

    def test_contrast_floors(self):
        for slug, th in themes():
            with self.subTest(theme=slug):
                self.assertGreaterEqual(contrast(th.ink, th.paper), 7.0, "墨/纸（正文必须好读）")
                self.assertGreaterEqual(contrast(th.accent_text, th.paper), 3.0, "强调色当文字")
                self.assertGreaterEqual(contrast(th.accent_mark, th.paper), 3.0, "强调色当细线")
                self.assertGreaterEqual(contrast(th.on_ink, th.ink), 3.0, "反色页上的字")

    def test_accent_readable_on_paper_and_zebra(self):
        """表格评级列用的颜色：纸色和斑马行上都要 ≥ 4.5。"""
        for slug, th in themes():
            for bg in (th.paper, th.zebra()):
                with self.subTest(theme=slug, bg=bg):
                    self.assertGreaterEqual(contrast(th.accent_readable(bg), bg), 4.5 - 1e-6)

    def test_type_scale_is_design_not_document(self):
        for slug, th in themes():
            with self.subTest(theme=slug):
                self.assertGreaterEqual(th.display / th.body, 3.5, "Display÷Body < 3.5 就成文档了")
                seq = [th.display, th.h1, th.lead, th.body, th.small, th.label]
                self.assertEqual(seq, sorted(seq, reverse=True), "字号阶必须逐级递减")

    def test_fonts_are_bundled_or_system(self):
        for slug, th in themes():
            for key in ("display_latin", "body_latin", "mono"):
                fam = th.__dict__[{"display_latin": "d_lat", "body_latin": "b_lat", "mono": "mono"}[key]]
                stem = fam.split()[0].lower()
                with self.subTest(theme=slug, font=fam):
                    self.assertTrue(
                        any(stem == b or fam.replace(" ", "").lower().startswith(b) for b in BUNDLED)
                        or fam.lower() in SYSTEM_LATIN,
                        f"{fam} 既没随仓库分发，也不在系统字体白名单里（换台机器会掉字体）",
                    )


if __name__ == "__main__":
    unittest.main()
