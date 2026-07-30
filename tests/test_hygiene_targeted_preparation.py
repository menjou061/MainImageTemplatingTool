"""Keep the Windows workflow scoped to the layout groups selected from Excel."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.ps1"
BATCH = ROOT / "L0_Windows命令行版" / "_internal" / "batch_template.jsx"


class HygieneTargetedPreparationTest(unittest.TestCase):
    def test_data_preflight_selects_layout_groups_before_photoshop_opens(self) -> None:
        source = RUNNER.read_text(encoding="utf-8-sig")

        preflight = source.index("$activeLayoutGroups = Get-ActiveRecordLayoutGroups")
        photoshop = source.index("$photoshop = Start-Photoshop", preflight)
        self.assertLess(preflight, photoshop)
        self.assertIn("active_layout_groups", source)
        self.assertIn("仅检查并映射本次商品实际使用的版式图层", source)

    def test_template_preparation_timing_is_recorded(self) -> None:
        source = RUNNER.read_text(encoding="utf-8-sig")

        self.assertIn('Get-ElapsedText -StartedAt $startedAt -Label ("PSD 模板 " + $Mode)', source)
        self.assertIn("本次映射版式：", source)

    def test_hygiene_prefers_business_provided_prepared_sibling(self) -> None:
        source = RUNNER.read_text(encoding="utf-8-sig")

        self.assertIn("function Get-PreparedTemplateSibling", source)
        self.assertIn("_套版模板.psd", source)
        self.assertIn("标准模板副本体检通过", source)
        self.assertIn("E_TEMPLATE_PREP_REQUIRED", source)

    def test_batch_preflight_uses_active_layout_scope(self) -> None:
        source = BATCH.read_text(encoding="utf-8")

        self.assertIn("profile.active_layout_groups && profile.active_layout_groups.length > 0", source)
        self.assertIn("collectRecordLayoutMatches(template, requested, matches)", source)


if __name__ == "__main__":
    unittest.main()
