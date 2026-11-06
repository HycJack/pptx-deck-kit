"""linter 的回归测试：8 条规则必须**真的会响**，干净版面必须**真的安静**。

为什么要有这个：lint 是这套流程唯一的自动审计，规则写坏了不会有任何报错——
它只会变安静，于是所有人都以为版面干净。历史上就踩过两次：
表格（GraphicFrame）没有 `.text_frame` 直接崩、幻灯片 `<p:bg>` 没读导致白字压深底误报 1.00。
"""

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skill" / "scripts"))

from pptx import Presentation                       # noqa: E402
from pptx.dml.color import RGBColor                  # noqa: E402
from pptx.enum.text import PP_ALIGN                  # noqa: E402
from pptx.util import Inches, Pt                     # noqa: E402

import pptx_kit as K                                 # noqa: E402
from lint_pptx import check                          # noqa: E402

LATIN, EA = "Barlow", "微软雅黑"
INK, BODY = "1A1A1A", "3A3A3A"


def blank():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs


def page(prs, bg=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if bg:
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(bg)
    return s


def run(prs):
    """存盘再 lint，返回 (problems, warnings)。"""
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "t.pptx"
        prs.save(p)
        _, problems, warnings = check(p)
    return problems, warnings


def hits(items, rule):
    return [i for i in items if i["rule"] == rule]


class Clean(unittest.TestCase):
    def test_clean_page_is_silent(self):
        prs = blank()
        s = page(prs)
        K.textbox(s, 0.8, 0.8, 8.0, 1.4, text="设计稿标题", size=56, color=INK, latin=LATIN, ea=EA)
        K.textbox(s, 0.8, 3.2, 6.0, 0.6, text="正文一行说明文字", size=14, color=BODY, latin=LATIN, ea=EA)
        problems, warnings = run(prs)
        self.assertEqual((problems, warnings), ([], []))

    def test_every_rule_still_exists_in_source(self):
        """规则被误删时，上面那些用例会失效——先卡住规则名本身。"""
        src = (ROOT / "skill" / "scripts" / "lint_pptx.py").read_text(encoding="utf-8")
        for rule in ("R1", "R2", "R3", "R4", "R5", "R6", "R7", "R8"):
            self.assertIn(f'"{rule}"', src, f"{rule} 不见了")


class Rules(unittest.TestCase):
    def test_r1_out_of_bounds(self):
        prs = blank()
        s = page(prs)
        K.rect(s, 12.6, 1.0, 2.0, 1.0, fill="333333")
        problems, _ = run(prs)
        self.assertTrue(hits(problems, "R1"), "越界没被抓到")

    def test_r2_text_overflow(self):
        prs = blank()
        s = page(prs)
        K.textbox(s, 0.8, 1.0, 3.0, 0.45, text="这是一段很长的中文" * 4, size=32, color=INK,
                  latin=LATIN, ea=EA)
        problems, _ = run(prs)
        self.assertTrue(hits(problems, "R2"), "溢出没被抓到")

    def test_r4_textbox_overlap(self):
        prs = blank()
        s = page(prs)
        for _ in range(2):
            K.textbox(s, 1.0, 1.0, 4.0, 1.0, text="重叠测试", size=20, color=INK, latin=LATIN, ea=EA)
        problems, _ = run(prs)
        self.assertTrue(hits(problems, "R4"), "文本框重叠没被抓到")

    def test_r5_flat_type_scale(self):
        prs = blank()
        s = page(prs)
        K.textbox(s, 0.8, 1.0, 6.0, 0.6, text="不算标题的标题", size=20, color=INK, latin=LATIN, ea=EA)
        K.textbox(s, 0.8, 3.0, 6.0, 0.6, text="正文跟它差不多大", size=18, color=BODY, latin=LATIN, ea=EA)
        _, warnings = run(prs)
        self.assertTrue(hits(warnings, "R5"), "字号对比度不足没被抓到")

    def test_r6_missing_ea(self):
        prs = blank()
        s = page(prs)
        tb = s.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(6.0), Inches(0.8))
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = "裸写的中文没有 a:ea"
        r.font.size = Pt(32)
        r.font.name = LATIN
        _, warnings = run(prs)
        self.assertTrue(hits(warnings, "R6"), "缺 a:ea 没被抓到")

    def test_r7_justify(self):
        prs = blank()
        s = page(prs)
        K.textbox(s, 0.8, 1.0, 6.0, 0.8, text="两端对齐的中文段落没有用", size=30, color=INK,
                  latin=LATIN, ea=EA, align=PP_ALIGN.JUSTIFY)
        _, warnings = run(prs)
        self.assertTrue(hits(warnings, "R7"), "algn=just 没被抓到")

    def test_r8_low_contrast(self):
        prs = blank()
        s = page(prs)
        K.rect(s, 1.0, 1.0, 5.0, 2.5, fill="ED7D31")          # 橙底
        K.textbox(s, 1.2, 1.6, 4.6, 1.0, text="白字压橙", size=40, color="FFFFFF",
                  latin=LATIN, ea=EA)
        problems, _ = run(prs)
        found = hits(problems, "R8")
        self.assertTrue(found, "低对比度没被抓到")
        self.assertIn("2.7", found[0]["msg"])                  # 白/橙 实测 2.77


class Regressions(unittest.TestCase):
    """三次真实 bug：表格、幻灯片背景、整表当一段算溢出。"""

    def test_table_cells_do_not_crash(self):
        prs = blank()
        s = page(prs)
        tbl = s.shapes.add_table(3, 2, Inches(1.0), Inches(1.0), Inches(6.0), Inches(2.0)).table
        for r in range(3):
            for c in range(2):
                tbl.cell(r, c).text = f"格{r}{c}"
        problems, warnings = run(prs)                      # 不抛异常即通过
        self.assertFalse(hits(problems, "R2"), f"短文本表格被误报溢出：{hits(problems, 'R2')}")

    def test_slide_background_is_read(self):
        prs = blank()
        s = page(prs, bg="1E3A5F")                          # 深底
        K.textbox(s, 1.0, 2.0, 8.0, 1.2, text="深底上的白字", size=48, color="FFFFFF",
                  latin=LATIN, ea=EA)
        problems, _ = run(prs)
        self.assertFalse(hits(problems, "R8"), "深底白字被误报（<p:bg> 没读到）")

    def test_table_cell_fill_is_local_background(self):
        prs = blank()
        s = page(prs)
        tbl = s.shapes.add_table(1, 1, Inches(1.0), Inches(1.0), Inches(5.0), Inches(1.2)).table
        cell = tbl.cell(0, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string("ED7D31")
        cell.text = "表头白字"
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(24)
                r.font.color.rgb = RGBColor.from_string("FFFFFF")
        problems, _ = run(prs)
        self.assertTrue(hits(problems, "R8"), "单元格底色没被当成局部底色")


if __name__ == "__main__":
    unittest.main()
