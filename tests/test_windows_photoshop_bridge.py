"""Static regression coverage for the Windows Photoshop bridge."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.ps1"


class WindowsPhotoshopBridgeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = RUNNER.read_text(encoding="utf-8-sig")

    def test_zero_main_window_handle_is_a_warning_not_a_blocker(self) -> None:
        self.assertIn("W_PS_WINDOW_HANDLE_UNAVAILABLE", self.source)
        self.assertIn("Warning = $windowWarning", self.source)
        self.assertNotIn("Code = 'PS_NOT_READY'", self.source)
        self.assertIn("继续验证 COM 自动化接口", self.source)

    def test_job_receives_a_scalar_progid_payload(self) -> None:
        self.assertIn("$progIdPayload = (($progIds | ForEach-Object { [string]$_ }) -join \"`n\")", self.source)
        self.assertIn("param([string]$candidateProgIdPayload, [string]$jsx)", self.source)
        self.assertIn("$candidateProgIdPayload -split \"`n\"", self.source)
        self.assertIn("-ArgumentList $progIdPayload, $ScriptText", self.source)
        self.assertNotIn("-ArgumentList (,$progIds), $ScriptText", self.source)

    def test_progress_does_not_claim_com_connected_before_the_probe(self) -> None:
        self.assertIn("正在连接 Photoshop 并准备打开模板", self.source)
        self.assertNotIn("数据预检通过，已连接 Photoshop，准备打开模板", self.source)

    def test_template_identity_blocks_before_photoshop_starts(self) -> None:
        identity = self.source.index("模板身份校验")
        start = self.source.index("$photoshop = Start-Photoshop")
        missing_identity_block = self.source.index("if ($identityMissing -and $profileId -ne 'legacy-v1')")
        self.assertLess(identity, start)
        self.assertLess(missing_identity_block, start)
        self.assertIn("E_TEMPLATE_IDENTITY_MISMATCH", self.source)
        self.assertIn("E_TEMPLATE_IDENTITY_MISSING", self.source)
        self.assertIn("template_identity.py", self.source)


if __name__ == "__main__":
    unittest.main()
