"""Regression coverage for the hygiene template's business-approved gift cards."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "L0_Windows命令行版" / "_internal" / "template_prepare.jsx"


class HygieneTemplatePreparationTest(unittest.TestCase):
    def test_750_gift_slots_target_the_three_member_cards_in_business_order(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        start = source.index("function prepareHygieneLayoutGroup")
        end = source.index("function prepareHygieneTemplate", start)
        helper = source[start:end]
        self.assertIn('renameGiftImage(document, topGiftLayers[0], "赠品图1")', helper)
        self.assertIn('normalizeGiftSlot(document, giftGroups[0], "赠品图2")', helper)
        self.assertIn('normalizeGiftSlot(document, giftGroups[1], "赠品图3")', helper)
        self.assertIn('dynamicGiftText[copyIndex].name = "@赠品文案" + (copyIndex + 1)', helper)

    def test_750_preflight_checks_the_named_material_not_card_decorations(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        start = source.index("function hygieneStructureProblems")
        end = source.index("function hygieneSourceHasMultipleLayouts", start)
        helper = source[start:end]
        self.assertIn("var slotObjects = smartObjectLayersWithin(slotGroups[slotIndex]);", helper)
        self.assertIn("if (slotObjects.length !== 1)", helper)

    def test_switch_groups_are_checked_with_their_hash_prefix(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        start = source.index("function hygieneProblems")
        end = source.index("function hygieneStructureProblems", start)
        helper = source[start:end]
        self.assertIn('var expectedGroups = usesStaticSupportArt() || giftFieldsAreOptional ? [] : ["#赠品区域"];', helper)

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

    def test_multi_layout_source_requires_prepared_copy(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8")

        self.assertIn("function inactiveHygieneLayouts(document, layerIndex)", source)
        self.assertIn("function hygieneSourceHasMultipleLayouts(document)", source)
        self.assertIn('status: "BLOCKED_PREP_REQUIRED"', source)
        self.assertIn("E_TEMPLATE_PREP_REQUIRED", source)
        self.assertNotIn("function removeInactiveHygieneLayouts", source)
        self.assertIn("请先生成并确认标准副本", source)
        blocked_return = source.index("if (result.status !== \"PREPARED\")")
        save = source.index("document.saveAs(outputFile")
        self.assertLess(blocked_return, save)


if __name__ == "__main__":
    unittest.main()
