"""Static regression checks for Windows Photoshop timeout/output guards.

These checks intentionally do not claim a Windows Photoshop success. They keep
the failure contract present and reviewable on CI/Mac hosts without PowerShell.
"""
from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.ps1"
RUNNER_BAT = ROOT / "L0_Windows命令行版" / "_internal" / "L0_Run.bat"
START_CMD = ROOT / "L0_Windows命令行版" / "L0_Start.cmd"


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

    def test_interactive_failure_has_primary_and_launcher_fallback_dialogs(self) -> None:
        self.assertIn("function Show-TaskFailureDialog", self.source)
        self.assertIn("failure_dialog_shown.marker", self.source)
        self.assertIn("Show-TaskFailureDialog -ErrorSummary $message", self.source)

        launcher = RUNNER_BAT.read_text(encoding="utf-8-sig")
        self.assertIn("failure_dialog_shown.marker", launcher)
        self.assertIn(":SHOW_FALLBACK_FAILURE", launcher)
        self.assertIn("checkpoint=showing_fallback_failure_dialog", launcher)
        self.assertNotIn("details: %PS_CONSOLE%", launcher)
        self.assertIn('del /q "%STARTUP_DIR%\\failure_dialog_shown.marker"', launcher)
        self.assertGreaterEqual(launcher.count("call :SHOW_FALLBACK_FAILURE"), 3)

    def test_early_launcher_paths_clear_stale_marker_and_show_a_dialog(self) -> None:
        launcher = RUNNER_BAT.read_text(encoding="utf-8-sig")
        marker_clear = launcher.index('del /q "%STARTUP_DIR%\\failure_dialog_shown.marker"')
        entry_failure = launcher.index('call :WRITE_FAILURE "entry"')
        runner_failure = launcher.index('call :WRITE_FAILURE "runner"')
        powershell_failure = launcher.index('call :WRITE_FAILURE "powershell"')

        self.assertLess(marker_clear, entry_failure)
        for failure in (entry_failure, runner_failure, powershell_failure):
            self.assertIn("call :SHOW_FALLBACK_FAILURE", launcher[failure : failure + 220])

    def test_outer_launcher_has_a_final_fallback_when_runner_exits_early(self) -> None:
        launcher = START_CMD.read_text(encoding="utf-8-sig")
        runner_return = launcher.index("checkpoint=runner_returned")
        fallback = launcher.index("call :SHOW_STARTUP_FAILURE", runner_return)

        self.assertLess(runner_return, fallback)
        self.assertIn("failure_dialog_shown.marker", launcher[runner_return:fallback])
        self.assertNotIn("details: %LOG%", launcher)


if __name__ == "__main__":
    unittest.main()
