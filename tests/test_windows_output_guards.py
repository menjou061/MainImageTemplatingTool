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

    def test_success_path_requires_report_and_real_jpg_psd_signatures(self) -> None:
        self.assertIn("function Assert-PhotoshopOutputArtifacts", self.source)
        self.assertIn("E_OUTPUT_INCOMPLETE", self.source)
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


if __name__ == "__main__":
    unittest.main()
