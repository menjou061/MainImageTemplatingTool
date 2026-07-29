"""Regression coverage for the hygiene/Tmall one-row-per-product workbook."""
from __future__ import annotations

import csv
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
INTERNAL = ROOT / "L0_Windows命令行版" / "_internal"
sys.path.insert(0, str(INTERNAL))
SPEC = importlib.util.spec_from_file_location("l0_clean_data", INTERNAL / "clean_data.py")
assert SPEC and SPEC.loader
clean_data = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(clean_data)


class HygieneRecordRowsTest(unittest.TestCase):
    def test_record_rows_parse_prices_pieces_and_gift_pairs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            product = root / "product.png"
            gift_one = root / "gift-one.png"
            gift_two = root / "gift-two.png"
            gift_three = root / "gift-three.png"
            for path in (product, gift_one, gift_two, gift_three):
                path.touch()
            workbook_path = root / "hygiene.xlsx"
            output = root / "output"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "跑批数据"
            sheet.append(clean_data.RECORD_ROW_HEADERS)
            sheet.append(
                [
                    "618预售品", "新超薄", "100%纯棉* 更干爽不黏腻", "*指棉面层",
                    "售卖07/18 00:00:00 - 07/19 23:59:59", str(product), "2套含赠到手242片", "2",
                    "58.3(第2套到手预估)=79.5（活动价）-2（入会领商品券）-19.2（官方立减12%）",
                    "\n".join((str(gift_one), str(gift_two), str(gift_three))),
                    "会员0元试用\n新会员0.01元拍下得\n直播间下单前100送", "",
                ]
            )
            workbook.save(workbook_path)

            count, errors, data_path, _, error_path = clean_data.build_data(
                workbook_path, "跑批数据", output, None, profile_id="hygiene-tmall-v1.2", variant="main-750"
            )

            self.assertEqual((count, errors), (1, 0))
            with data_path.open(encoding="utf-8-sig", newline="") as handle:
                record = next(csv.DictReader(handle))
            self.assertEqual(record["卖点"], "100%纯棉* 更干爽不黏腻")
            self.assertEqual(record["备注"], "*指棉面层")
            self.assertEqual(record["片数套"], "2套到手")
            self.assertEqual(record["片数数量"], "242片")
            self.assertEqual(record["到手"], "¥58.3")
            self.assertEqual(record["价格活动价"], "79.5")
            self.assertEqual(record["价格优惠券"], "2")
            self.assertEqual(record["价格立减"], "19.2")
            self.assertEqual(record["赠品图3"], str(gift_three))
            self.assertEqual(record["赠品文案3"], "直播间下单前100送")
            self.assertEqual(record["版式组"], "无代言人")
            self.assertFalse(error_path.read_text(encoding="utf-8-sig").splitlines()[1:])


if __name__ == "__main__":
    unittest.main()
