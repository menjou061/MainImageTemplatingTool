#!/usr/bin/env python3
"""List visible sheets or product names in an Excel workbook for the Windows L0 launcher."""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def as_text(value) -> str:
    return "" if value is None else str(value).strip()


def is_display_image_formula(value) -> bool:
    if not isinstance(value, str):
        return False
    normalized = "".join(value.split()).lower()
    return normalized.startswith("=") and "dispimg" in normalized


def list_sheets(workbook_path: Path) -> int:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    for worksheet in workbook.worksheets:
        if worksheet.sheet_state == "visible" and worksheet.title != "WpsReserved_CellImgList":
            print(worksheet.title)
    return 0


def list_products(workbook_path: Path, sheet_name: str) -> int:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    if sheet_name not in workbook.sheetnames:
        print(f"Sheet 不存在：{sheet_name}", file=sys.stderr)
        return 2
    worksheet = workbook[sheet_name]
    if worksheet.sheet_state != "visible" or worksheet.title == "WpsReserved_CellImgList":
        print(f"Sheet 不可处理：{sheet_name}", file=sys.stderr)
        return 2

    seen: set[str] = set()
    for column in range(2, worksheet.max_column + 1):
        value = worksheet.cell(1, column).value
        if is_blank(value) or is_display_image_formula(value):
            continue
        product = as_text(value)
        if product in seen:
            continue
        seen.add(product)
        print(product)
    return 0


def main() -> int:
    if len(sys.argv) not in {2, 4}:
        print("用法：l0_list_sheets.py <xlsx> 或 l0_list_sheets.py --products <xlsx> <sheet>", file=sys.stderr)
        return 2

    product_mode = len(sys.argv) == 4 and sys.argv[1] == "--products"
    workbook_path = Path(sys.argv[2] if product_mode else sys.argv[1])
    if not workbook_path.is_file():
        print(f"Excel 文件不存在：{workbook_path}", file=sys.stderr)
        return 2

    if product_mode:
        return list_products(workbook_path, sys.argv[3])
    if len(sys.argv) == 4:
        print("用法：l0_list_sheets.py --products <xlsx> <sheet>", file=sys.stderr)
        return 2
    return list_sheets(workbook_path)


if __name__ == "__main__":
    raise SystemExit(main())
