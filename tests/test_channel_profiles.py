"""Contract tests for versioned channel profiles; no Photoshop integration is faked."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INTERNAL = ROOT / "L0_Windows命令行版" / "_internal"
sys.path.insert(0, str(INTERNAL))

from channel_profile import ProfileError, get_profile, get_variant_for_sheet, load_profiles, map_vertical_headers


class ChannelProfileTest(unittest.TestCase):
    def test_tmall_mapping_and_ignored_columns(self) -> None:
        profile = get_profile("tmall-positional-v1", require_enabled=False)
        mapping = map_vertical_headers(
            ["变量名称", "SKU", "变量01", "变量02", "变量03", "变量04", "变量05", "变量06", "变量07", "图片目录路径", "变量08", "变量09"],
            profile,
        )
        self.assertEqual(mapping["SKU"], "文件名称")
        self.assertEqual(mapping["变量01"], "利益点1")
        self.assertEqual(mapping["变量03"], "预估到手价")
        self.assertEqual(mapping["图片目录路径"], "产品")
        self.assertNotIn("变量08", mapping)
        self.assertNotIn("变量09", mapping)
        self.assertEqual(profile["execution_mode"], "legacy_layer_names")
        self.assertEqual(profile["variant_selection"], "sheet")
        self.assertEqual(profile["sheet"]["sku_header"], "变量名称")
        self.assertEqual(profile["required_psd_variables"][0], {"name": "利益点1", "type": "text"})
        self.assertEqual(profile["required_psd_variables"][-1], {"name": "产品", "type": "smart_object"})

    def test_schema_mismatch_is_rejected(self) -> None:
        profile = get_profile("tmall-positional-v1", require_enabled=False)
        with self.assertRaisesRegex(ProfileError, "E_PROFILE_SCHEMA_MISMATCH"):
            map_vertical_headers(["变量名称", "SKU", "变量01"], profile)

    def test_750_and_800_variants_are_isolated(self) -> None:
        profile = get_profile("tmall-positional-v1", require_enabled=False)
        tmall750 = get_profile("tmall-positional-v1", "main-750", require_enabled=False)
        tmall800 = get_profile("tmall-positional-v1", "main-800", require_enabled=False)
        self.assertEqual(tmall750["target_size"], {"width": 1440, "height": 1920})
        self.assertEqual(tmall800["target_size"], {"width": 1440, "height": 1440})
        self.assertNotIn("export_size", tmall750)
        self.assertNotIn("export_size", tmall800)
        self.assertEqual(tmall750["sheet_name"], "现货-750")
        self.assertEqual(tmall800["sheet_name"], "现货-800")
        self.assertEqual(tmall750["template_bindings"]["产品"], "DT17100-20")
        self.assertEqual(get_variant_for_sheet(profile, "现货-800"), "main-800")
        self.assertEqual(get_variant_for_sheet(profile, "现货-750"), "main-750")
        with self.assertRaisesRegex(ProfileError, "E_PROFILE_SHEET_MISMATCH"):
            get_variant_for_sheet(profile, "隐藏 Sheet")
        with self.assertRaisesRegex(ProfileError, "E_PROFILE_UNSUPPORTED"):
            get_profile("jd-main-800-v1", "main-750", require_enabled=False)

    def test_tmall_profile_is_enabled_and_unapproved_profile_is_rejected(self) -> None:
        self.assertEqual(get_profile("tmall-positional-v1")["profile_id"], "tmall-positional-v1")
        with self.assertRaisesRegex(ProfileError, "E_PROFILE_UNSUPPORTED"):
            get_profile("jd-main-800-v1")

    def test_legacy_profile_uses_default_variant_without_sheet_mapping(self) -> None:
        profile = get_profile("legacy-v1", require_enabled=False)
        self.assertEqual(get_variant_for_sheet(profile, "任意可见 Sheet"), "main-800")

    def test_hygiene_tmall_profile_uses_the_record_row_contract(self) -> None:
        profile = get_profile("hygiene-tmall-v1.2")
        self.assertEqual(profile["layout"], "record_rows")
        self.assertEqual(profile["target_size"], {"width": 750, "height": 1000})
        self.assertEqual(profile["output_label"], "750")
        self.assertEqual(profile["category"], "卫品")
        self.assertEqual(profile["channel"], "天猫官旗")
        self.assertIn("输出文件名", profile["sheet"]["required_headers"])
        self.assertTrue(any("备注" in headers for headers in profile["sheet"]["required_header_sets"]))
        self.assertNotIn("价格优惠券", profile["record_required_fields"])
        self.assertIn("小马无侧边", profile["record_layout"]["groups"])
        self.assertEqual(profile["variant_selection"], "sheet")
        self.assertNotIn("batch_variants", profile)
        self.assertEqual(get_variant_for_sheet(profile, "出图数据"), "main-750")
        self.assertNotIn("main-800", profile["variants"])
        self.assertEqual(profile["pending_variants"]["main-800"]["status"], "pending_asset_approval")
        with self.assertRaisesRegex(ProfileError, "E_PROFILE_SHEET_MISMATCH"):
            get_variant_for_sheet(profile, "卫品-800")

    def test_enabled_category_channel_combinations_are_complete_and_unique(self) -> None:
        enabled = [profile for profile in load_profiles().values() if profile["status"] == "enabled"]
        combinations = [(profile.get("category"), profile.get("channel")) for profile in enabled]
        self.assertNotIn((None, None), combinations)
        self.assertTrue(all(category and channel for category, channel in combinations))
        self.assertEqual(len(combinations), len(set(combinations)))
        self.assertIn(("纸品", "京东自营"), combinations)
        self.assertIn(("纸品", "天猫官旗"), combinations)
        self.assertIn(("卫品", "天猫官旗"), combinations)


if __name__ == "__main__":
    unittest.main()
