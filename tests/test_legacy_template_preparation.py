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

    def test_first_import_bootstraps_only_a_prepared_legacy_copy(self) -> None:
        runner = RUNNER.read_text(encoding="utf-8-sig")
        source = SCRIPT.read_text(encoding="utf-8-sig")

        self.assertIn("$sourcePsdPath = $psdPath", runner)
        self.assertIn("$preparedCopyIsSeparate", runner)
        self.assertIn("$preparedCopyCreatedThisRun", runner)
        self.assertIn("$profileId -ne 'legacy-v1'", runner)
        self.assertIn("不会为历史副本或原 PSD 自动签发身份文件", runner)
        self.assertIn("Write-PreparedTemplateIdentity", runner)
        self.assertIn("Get-PreparedTemplatePaths", runner)
        self.assertIn("$preparedCopyWasPresentBefore", runner)
        self.assertIn("-AllowExistingPreparedSibling:(-not $identityMissing)", runner)
        self.assertIn("-ForceFreshCopy:$identityMissing", runner)
        self.assertIn("-not ($ForceFreshCopy -and $check.Status -eq 'READY')", runner)
        self.assertIn("忽略历史模板副本，只为本次任务生成新的独立副本", runner)
        self.assertIn("if (-not $preparedCopyCreatedThisRun)", runner)
        self.assertIn("不会为历史副本或原 PSD 自动签发身份文件", runner)
        self.assertIn("-SourcePsdPath $sourcePsdPath", runner)
        self.assertIn("function Test-TemplateIdentity", runner)
        self.assertIn("$ErrorActionPreference = 'Continue'", runner)
        self.assertIn("$exitCode -eq 0", runner)
        self.assertIn("[bool]$ForceCopy = $false", runner)
        self.assertIn("forceCopy: ", runner)
        self.assertIn("var forceCopy = !!inputs.forceCopy;", source)
        self.assertIn('inspection.status === "READY" && !forceCopy', source)
        self.assertIn('已生成本次独立副本用于建立模板身份', source)

    def test_interactive_failures_remain_visible(self) -> None:
        runner = RUNNER.read_text(encoding="utf-8-sig")

        self.assertIn("function Show-TaskFailureDialog", runner)
        self.assertIn("'套版未完成'", runner)
        self.assertIn("Show-TaskFailureDialog -ErrorSummary $message", runner)


if __name__ == "__main__":
    unittest.main()
