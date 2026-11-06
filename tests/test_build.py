"""成品链路：版式能排出来、排出来能过 lint、页数对得上。

这是 `tasks.py verify` 的快速版（只跑几套主题，约 6 秒），CI 里跑这个；
全量 14×13 交给 `tasks.py verify`。
"""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skill" / "scripts"))

from layouts import LAYOUTS                            # noqa: E402

PY = sys.executable
LAYOUTS_PY = ROOT / "skill" / "scripts" / "preview_layouts.py"
LINT_PY = ROOT / "skill" / "scripts" / "lint_pptx.py"


def build(theme, out):
    r = subprocess.run([PY, str(LAYOUTS_PY), str(out), "--theme", theme],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    assert r.returncode == 0, r.stderr
    return r.stdout


def lint(path):
    r = subprocess.run([PY, str(LINT_PY), str(path)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.returncode, r.stdout


class Build(unittest.TestCase):
    def test_13_layouts_registered(self):
        self.assertEqual(len(LAYOUTS), 13, f"版式数变了：{sorted(LAYOUTS)}")

    def test_every_layout_slug_is_reachable(self):
        """13 个版式都要出现在自动排版的清单里（防止漏挂）。"""
        r = subprocess.run([PY, str(LAYOUTS_PY), "--list"], capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        self.assertEqual(r.returncode, 0, r.stderr)
        for slug in LAYOUTS:
            self.assertIn(slug, r.stdout, f"{slug} 没出现在 --list 里")

    def test_all_13_pages_lint_clean(self):
        for theme in ("editorial-forest", "studio"):
            with self.subTest(theme=theme), tempfile.TemporaryDirectory() as td:
                out = Path(td) / f"{theme}.pptx"
                build(theme, out)
                rc, log = lint(out)
                self.assertEqual(rc, 0, log)
                first = log.splitlines()[0]
                self.assertIn("error 0", log)
                self.assertIn("warn 0", log, f"{theme} 有 warn：\n{log}")
                self.assertEqual(int(first.split()[1]), 13, f"页数不对：{first}")

    def test_dark_and_high_saturation_themes_stay_clean(self):
        """换色最容易暴露版式问题：纯色深底与高饱和强调各跑一遍。"""
        for theme in ("neo-grid-bold", "monochrome"):
            with self.subTest(theme=theme), tempfile.TemporaryDirectory() as td:
                out = Path(td) / "d.pptx"
                build(theme, out)
                rc, log = lint(out)
                self.assertIn("error 0", log, f"{theme} 有 error：\n{log}")
                self.assertIn("warn 0", log, f"{theme} 有 warn：\n{log}")


if __name__ == "__main__":
    unittest.main()
