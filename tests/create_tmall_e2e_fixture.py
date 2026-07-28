#!/usr/bin/env python3
"""Create a two-sheet fixture for the real Tmall channel template regression."""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook


def add_sheet(workbook: Workbook, name: str, headers: list[str], row: list[object]) -> None:
    worksheet = workbook.create_sheet(name)
    worksheet.append(headers)
    worksheet.append(row)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("--product-image", required=True)
    args = parser.parse_args()

    output_path = Path(args.output).resolve()
    image_path = Path(args.product_image).resolve()
    if not image_path.is_file():
        raise SystemExit(f"Product image does not exist: {image_path}")

    workbook = Workbook()
    workbook.remove(workbook.active)
    add_sheet(
        workbook,
        "现货-750",
        ["变量名称", "变量01", "变量02", "变量03", "变量04", "变量05", "变量06", "变量07", "图片目录路径"],
        ["BT2610-27", "买1减16", "买2减36", "2件预估均价", 71, ".9", "茶语卷纸4层200克27卷", "200克大胖纸 大卷更耐用", str(image_path)],
    )
    add_sheet(
        workbook,
        "现货-800",
        ["变量名称", "变量01", "变量02", "变量03", "变量04", "变量05", "变量06", "变量07", "图片目录路径", "变量08", "变量09"],
        ["DT31100-18-2-一元预定", "限时直降", "立减115元", "到手价", 119, ".8", "婴儿乳霜柔纸巾奶被纸100抽2箱共36包", "添加40%乳霜  安全0刺激", str(image_path), "", ""],
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


if __name__ == "__main__":
    main()
