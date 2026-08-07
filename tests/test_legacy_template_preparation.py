"""Regression coverage for deterministic JD legacy PSD preparation."""
from __future__ import annotations

import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "L0_Windows命令行版" / "_internal" / "template_prepare.jsx"
RUNNER = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.ps1"
BATCH = ROOT / "L0_Windows命令行版" / "_internal" / "batch_template.jsx"
PROFILES = ROOT / "L0_Windows命令行版" / "_internal" / "channel_profiles.json"


class LegacyTemplatePreparationTest(unittest.TestCase):
    def test_known_legacy_aliases_are_preparable(self) -> None:
        source = SCRIPT.read_text(encoding="utf-8-sig")

        self.assertIn('"@时间": "@活动时间"', source)
        self.assertIn('"!堆图": true', source)
        self.assertIn("function hasDeterministicLegacyMapping", source)
        self.assertIn("已识别京东旧版模板的唯一字段映射", source)
        self.assertIn("prepareLegacyPackageLayers(document)", source)
        self.assertIn('candidate.exists || File(candidate.fsName + ".template.json").exists', source)

        runner = RUNNER.read_text(encoding="utf-8-sig")
        self.assertIn("$selectedProfile.profile_id -eq 'legacy-v1'", runner)
        self.assertIn("$check.Status -eq 'NEEDS_PREP'", runner)

        batch = BATCH.read_text(encoding="utf-8")
        self.assertIn("function isOptionalProfileVariable", batch)
        self.assertIn("isOptionalProfileVariable(profile, required)", batch)

    def test_legacy_coupon_layers_are_optional_when_data_disables_coupon(self) -> None:
        profile_doc = json.loads(PROFILES.read_text(encoding="utf-8-sig"))
        profile = next(item for item in profile_doc["profiles"] if item["profile_id"] == "legacy-v1")

        self.assertEqual(profile["optional_psd_variables"], ["券名", "折扣", "券门槛"])
        required = {item["name"] for item in profile["required_psd_variables"]}
        self.assertTrue(set(profile["optional_psd_variables"]).issubset(required))

    def test_field_compliant_templates_prepare_without_identity_metadata(self) -> None:
        runner = RUNNER.read_text(encoding="utf-8-sig")
        source = SCRIPT.read_text(encoding="utf-8-sig")

        self.assertIn("$sourcePsdPath = $psdPath", runner)
        self.assertIn("Resolve-TemplateForTask -Application $photoshop -TemplatePath $sourcePsdPath", runner)
        self.assertNotIn("Test-TemplateIdentity -Python $python -TemplatePath $sourcePsdPath", runner)
        self.assertNotIn("E_TEMPLATE_IDENTITY_MISMATCH：PSD 模板身份未通过", runner)
        self.assertIn("通过字段体检的模板副本", runner)
        self.assertIn("[bool]$ForceCopy = $false", runner)
        self.assertIn("forceCopy: ", runner)
        self.assertIn("var forceCopy = !!inputs.forceCopy;", source)
        self.assertIn('inspection.status === "READY" && !forceCopy', source)
        self.assertIn('已生成本次独立副本', source)

    def test_interactive_failures_remain_visible(self) -> None:
        runner = RUNNER.read_text(encoding="utf-8-sig")

        self.assertIn("function Show-TaskFailureDialog", runner)
        self.assertIn("'套版未完成'", runner)
        self.assertIn("Show-TaskFailureDialog -ErrorSummary $message", runner)


if __name__ == "__main__":
    unittest.main()
