"""Regression coverage for the one-product-per-row channel Excel layout."""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
CLEAN_DATA_PATH = ROOT / "L0_Windows命令行版" / "_internal" / "clean_data.py"
sys.path.insert(0, str(CLEAN_DATA_PATH.parent))
SPEC = importlib.util.spec_from_file_location("l0_clean_data", CLEAN_DATA_PATH)
assert SPEC and SPEC.loader
clean_data = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(clean_data)


class VerticalChannelCompatibilityTest(unittest.TestCase):
    def test_unapproved_vertical_profile_is_rejected_instead_of_guessed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary_path = Path(temporary_directory)
            workbook_path = temporary_path / "channel.xlsx"
            output_directory = temporary_path / "output"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "渠道"
            worksheet.append(
                [
                    "变量名称",
                    "变量01",
                    "变量02",
                    "变量03",
                    "变量04",
                    "变量05",
                    "变量06",
                    "变量07",
                    "图片目录路径",
                    "变量08",
                    "变量09",
                ]
            )
            worksheet.append(
                [
                    "SKU-001",
                    "限时直降",
                    "立减115元",
                    "到手价",
                    119,
                    ".8",
                    "纸巾规格",
                    "柔软亲肤",
                    '=K2&"\\"&J2',
                    "SKU-001.png",
                    r"\\server\materials",
                ]
            )
            workbook.save(workbook_path)

            with self.assertRaisesRegex(ValueError, "E_PROFILE_SCHEMA_MISMATCH"):
                clean_data.build_data(workbook_path, "渠道", output_directory, None)


if __name__ == "__main__":
    unittest.main()
