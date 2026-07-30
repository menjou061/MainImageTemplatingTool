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
        self.assertIn("sortLayersByVisualPosition(giftGroups)", source)

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

    def test_every_layout_must_keep_all_dynamic_bindings(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn("findNamedLayersWithin(layout, expected[index], layerIndex)", source)
        self.assertIn('problems.push("E_VAR_UNBOUND: " + configured[layoutIndex] + "/" + expected[index])', source)

    def test_preparation_uses_only_the_layouts_selected_for_this_task(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn("function hygieneLayoutGroupNames()", source)
        self.assertIn("CHANNEL_PROFILE.active_layout_groups", source)
        self.assertIn("var groups = hygieneLayoutGroupNames();", source)
        self.assertIn("prepareHygieneTemplate(document, layerIndex", source)

    def test_large_psd_preparation_builds_one_index_and_skips_full_layer_report(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn("function buildLayerIndex(document)", source)
        self.assertIn("var layerIndex = buildLayerIndex(document);", source)
        report_start = source.index("function writePreparationReport")
        report_end = source.index("function main()", report_start)
        report = source[report_start:report_end]
        self.assertNotIn("addAllLayers(document, all)", report)
        self.assertIn("仅处理本次任务版式", report)

    def test_prepared_task_copy_removes_layouts_not_used_by_selected_rows(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn("function inactiveHygieneLayouts(document, layerIndex)", source)
        self.assertIn("function removeInactiveHygieneLayouts(document, layerIndex)", source)
        self.assertIn("inactive[index].layer.remove()", source)
        self.assertIn('status: "NEEDS_ISOLATION"', source)
        removal = source.index("var removedLayouts = removeInactiveHygieneLayouts(document")
        mapping = source.index("prepareHygieneTemplate(document, buildLayerIndex(document))")
        save = source.index("document.saveAs(outputFile")
        self.assertLess(removal, mapping)
        self.assertLess(removal, save)
        self.assertIn("未使用版式已从任务副本移除，原始 PSD 未修改", source)
        self.assertIn('inspection.status === "NEEDS_PREP" ? "建立图层映射" : "保留已通过体检的字段映射"', source)


if __name__ == "__main__":
    unittest.main()
