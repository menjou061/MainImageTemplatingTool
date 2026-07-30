"""Regression tests for optional multi-channel preflight fields."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INTERNAL = ROOT / "L0_Windows命令行版" / "_internal"
sys.path.insert(0, str(INTERNAL))

from clean_data import add_material_precheck, add_required_field_precheck


class OptionalFieldPrecheckTest(unittest.TestCase):
    def test_optional_profile_records_warnings_without_blocking_errors(self) -> None:
        record = {"商品图": r"C:\missing\sku.png"}

        add_required_field_precheck(record, {"活动时间", "卖点"}, warning=True)
        add_material_precheck(record, {"商品图"}, warning=True)

        self.assertIn("预检提醒", record)
        self.assertNotIn("预检异常", record)

    def test_strict_profile_keeps_missing_fields_as_errors(self) -> None:
        record = {"商品图": r"C:\missing\sku.png"}

        add_required_field_precheck(record, {"活动时间"})
        add_material_precheck(record, {"商品图"})

        self.assertIn("预检异常", record)
        self.assertNotIn("预检提醒", record)


if __name__ == "__main__":
    unittest.main()
