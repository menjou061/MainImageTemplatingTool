"""Regression coverage for the one-product-per-row channel Excel layout."""

from __future__ import annotations

import csv
import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
CLEAN_DATA_PATH = ROOT / "L0_Windows命令行版" / "_internal" / "clean_data.py"
SPEC = importlib.util.spec_from_file_location("l0_clean_data", CLEAN_DATA_PATH)
assert SPEC and SPEC.loader
clean_data = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(clean_data)


class VerticalChannelCompatibilityTest(unittest.TestCase):
    def test_positional_variables_and_image_formula_are_normalized(self) -> None:
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

            with patch.object(clean_data.os.path, "isfile", return_value=True):
                count, errors, data_path, _, _ = clean_data.build_data(
                    workbook_path, "渠道", output_directory, None
                )

            self.assertEqual((count, errors), (1, 0))
            with data_path.open(encoding="utf-8-sig", newline="") as handle:
                record = next(csv.DictReader(handle))
            self.assertEqual(record["商品文件名"], "SKU-001")
            self.assertEqual(record["折扣"], "限时直降")
            self.assertEqual(record["券名"], "立减115元")
            self.assertEqual(record["到手"], "到手价")
            self.assertEqual(record["价格1"], "119")
            self.assertEqual(record["价格2"], ".8")
            self.assertEqual(record["规格"], "纸巾规格")
            self.assertEqual(record["卖点"], "柔软亲肤")
            self.assertEqual(record["商品图"], r"\\server\materials\SKU-001.png")


if __name__ == "__main__":
    unittest.main()
