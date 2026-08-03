"""Static contract checks for the interactive Windows smoke helper."""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SMOKE = ROOT / "tests" / "windows_ui_smoke.ps1"


class WindowsUiSmokeContractTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = SMOKE.read_text(encoding="utf-8-sig")

    def test_ascii_entry_fallback_survives_archive_filename_decode(self) -> None:
        self.assertIn("Join-Path $ToolRoot 'L0_Start.cmd'", self.source)
        self.assertIn("Windows tar can decode a Unicode archive member", self.source)

    def test_product_count_requires_matching_report_count(self) -> None:
        self.assertIn("[int]$ProductCount = 0", self.source)
        self.assertIn("$selectionCount = if ($ProductCount -gt 0)", self.source)
        self.assertIn("for ($selectionIndex = 0; $selectionIndex -lt $selectionCount;", self.source)
        self.assertIn("$ProductCount -gt 0 -and $resultRows.Count -ne $ProductCount", self.source)


if __name__ == "__main__":
    unittest.main()
