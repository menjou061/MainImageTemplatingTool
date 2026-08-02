"""Static regression checks for Windows Photoshop timeout/output guards.

These checks intentionally do not claim a Windows Photoshop success. They keep
the failure contract present and reviewable on CI/Mac hosts without PowerShell.
"""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.ps1"


class WindowsOutputGuardTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = RUNNER.read_text(encoding="utf-8-sig")

    def test_timeout_is_configurable_and_has_bounded_default(self) -> None:
        self.assertIn("$script:DefaultPhotoshopTimeoutSeconds = 900", self.source)
        self.assertIn("MAINIMAGE_PHOTOSHOP_TIMEOUT_SECONDS", self.source)
        self.assertIn("photoshopTimeoutSeconds", self.source)
        self.assertIn("Start-Job -ScriptBlock", self.source)
        self.assertIn("E_PHOTOSHOP_TIMEOUT", self.source)
        self.assertIn("Stop-Job -Job $job", self.source)

    def test_save_user_settings_preserves_timeout_override(self) -> None:
        save_start = self.source.index("function Save-UserSettings")
        save_end = self.source.index("function New-RunProgressWindow", save_start)
        save_source = self.source[save_start:save_end]
        self.assertIn(
            "Get-SettingText -Settings $script:Settings -Name 'photoshopTimeoutSeconds'",
            save_source,
        )
        self.assertIn(
            "Add-Member -NotePropertyName 'photoshopTimeoutSeconds'",
            save_source,
        )

    def test_photoshop_job_receives_scalar_progid_payload(self) -> None:
        self.assertIn(
            '$progIdPayload = (($progIds | ForEach-Object { [string]$_ }) -join "`n")',
            self.source,
        )
        self.assertIn(
            'param([string]$candidateProgIdPayload, [string]$jsx)',
            self.source,
        )
        self.assertIn('$candidateProgIdPayload -split "`n"', self.source)
        self.assertIn('-ArgumentList $progIdPayload, $ScriptText', self.source)
        self.assertNotIn('-ArgumentList (,$progIds), $ScriptText', self.source)

    def test_success_path_requires_report_and_real_jpg_psd_signatures(self) -> None:
        self.assertIn("function Assert-PhotoshopOutputArtifacts", self.source)
        self.assertIn("E_OUTPUT_INCOMPLETE", self.source)
        self.assertIn("$rowsWithOutput = @($rows | Where-Object", self.source)
        self.assertIn("Kind 'jpg'", self.source)
        self.assertIn("Kind 'psd'", self.source)
        self.assertIn("$bytes[0] -ne 0xFF", self.source)
        self.assertIn("缺少 8BPS 文件头", self.source)
        self.assertIn("PSD 没有可编辑图层", self.source)
        self.assertIn("PSD 图层/资源区不完整", self.source)
        self.assertIn("Assert-PhotoshopOutputArtifacts -ResultReport $resultReport", self.source)

    def test_integrity_gate_precedes_report_move_and_success_status(self) -> None:
        integrity_index = self.source.index("Assert-PhotoshopOutputArtifacts -ResultReport $resultReport")
        move_index = self.source.index("Move-Item -LiteralPath $resultReport")
        success_index = self.source.index("Write-EntryStatus -Status 'success'")
        self.assertLess(integrity_index, move_index)
        self.assertLess(integrity_index, success_index)

    def test_optional_template_fields_warn_without_failing_batch(self) -> None:
        batch_source = (
            ROOT / "L0_Windows命令行版" / "_internal" / "batch_template.jsx"
        ).read_text(encoding="utf-8-sig")
        self.assertIn("recordBindingWarnings(layerIndex, record)", batch_source)
        self.assertIn("模板未绑定可选字段：", batch_source)
        self.assertIn("if (isOptionalProfileVariable(required))", batch_source)
        self.assertIn(
            "disableUnboundOptionalGroups(layerIndex, record, bindingWarnings)",
            batch_source,
        )

    def test_optional_template_fields_are_checked_before_batch_generation(self) -> None:
        runner_source = self.source
        template_source = (
            ROOT / "L0_Windows命令行版" / "_internal" / "template_prepare.jsx"
        ).read_text(encoding="utf-8-sig")
        self.assertIn("function Assert-TemplateDataBindings", runner_source)
        self.assertIn("-Mode 'data_check'", runner_source)
        self.assertIn("不通过时不会开始批量生成", runner_source)
        self.assertIn("function dataBindingProblems(document, layerIndex)", template_source)
        self.assertIn("DATA_BINDING_ERROR", template_source)
        self.assertIn("data_fields_with_values", template_source)
        self.assertIn("请换用带对应区域的 PSD，或清空这些字段后重试", runner_source)
        self.assertIn("$friendlyMessage = [string]$_.Exception.Message", runner_source)
        self.assertIn("本次任务未开始。", runner_source)
        self.assertLess(
            runner_source.index("Assert-TemplateDataBindings -Application $photoshop"),
            runner_source.index("Set-RunProgress -Stage 'Photoshop 正在导出'"),
        )


if __name__ == "__main__":
    unittest.main()
