"""Regression coverage for the hygiene template's three gift slots."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "L0_Windows命令行版" / "_internal" / "template_prepare.jsx"


class HygieneTemplatePreparationTest(unittest.TestCase):
    def test_gift_slots_keep_top_and_lower_pairs_in_business_order(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn('renameFirstText(topGiftCopy, "赠品文案1")', source)
        self.assertIn('"@赠品文案" + (copyIndex + 2)', source)
        self.assertIn('normalizeGiftSlot(document, giftGroups[0], "赠品图2")', source)
        self.assertIn('normalizeGiftSlot(document, giftGroups[1], "赠品图3")', source)

    def test_fixed_gift_button_and_disclaimer_are_not_bound_to_row_data(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn('candidateName === "拍即赠"', source)
        self.assertIn('candidateName.charAt(0) === "*"', source)

    def test_composite_lower_slot_is_normalized_before_material_replacement(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        start = source.index("function normalizeGiftSlot")
        end = source.index("function prepareHygieneLayoutGroup", start)
        helper = source[start:end]
        self.assertIn("convertToSmartObject(document, slotGroup, targetName)", helper)
        self.assertNotIn("renameGiftImage(document, sourceObjects", helper)


if __name__ == "__main__":
    unittest.main()
