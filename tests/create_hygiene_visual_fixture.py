#!/usr/bin/env python3
"""Create a one-row hygiene workbook that exercises all three gift slots."""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook


HEADERS = [
    "活动", "系列", "主卖点", "备注", "时间", "堆品路径", "片数", "主推数", "价格条",
    "赠品路径", "赠品文案", "代言IP路径",
]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--product-image", required=True)
    parser.add_argument("--gift-one", required=True)
    parser.add_argument("--gift-two", required=True)
    parser.add_argument("--gift-three", required=True)
    args = parser.parse_args()

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "出图数据"
    worksheet.append(HEADERS)
    worksheet.append([
        "超级88", "新超薄", "100%纯棉柔护* 清爽不闷肌", "*指棉面层",
        "07/07-07/09", args.product_image, "2套到手188片", "2",
        "54.1(第2套到手预估)=84.5（活动价）-10（入会领商品券）-20.4（官方立减12%）",
        "\n".join((args.gift_one, args.gift_two, args.gift_three)),
        "赠品一 文案\n赠品二 文案\n赠品三 文案", "",
    ])
    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output)


if __name__ == "__main__":
    main()
