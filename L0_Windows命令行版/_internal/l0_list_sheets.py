#!/usr/bin/env python3
"""List visible sheets or product names in an Excel workbook for the Windows L0 launcher."""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook


VERTICAL_FIELD_NAMES = {
    "变量名称",
    "图片目录路径",
    "图片路径",
    "图片文件路径",
    "商品图",
    "商品图片",
}


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def as_text(value) -> str:
    return "" if value is None else str(value).strip()


def is_display_image_formula(value) -> bool:
    if not isinstance(value, str):
        return False
    normalized = "".join(value.split()).lower()
    return normalized.startswith("=") and "dispimg" in normalized


def looks_like_vertical_header(value) -> bool:
    text = as_text(value)
    return text in VERTICAL_FIELD_NAMES or text.startswith("变量")


def is_vertical_layout(worksheet) -> bool:
    """Detect the channel format with one row per product.

    The standard workbook is transposed: field names are in column A and
    product names are in row 1.  New channel sheets put field names in row 1
    and product IDs in column A.  Requiring at least two variable-like headers
    keeps ordinary product names from being mistaken for a schema.
    """
    headers = {as_text(worksheet.cell(1, column).value) for column in range(1, worksheet.max_column + 1)}
    if "文件名称" in headers and any(
        header in headers for header in ("产品（精确到图片名）", "产品图路径", "图片目录路径")
    ):
        return worksheet.max_row >= 2 and worksheet.max_column >= 2
    row_score = sum(
        1
        for column in range(1, worksheet.max_column + 1)
        if looks_like_vertical_header(worksheet.cell(1, column).value)
    )
    return row_score >= 2 and worksheet.max_row >= 2 and worksheet.max_column >= 2


def list_sheets(workbook_path: Path) -> int:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    for worksheet in workbook.worksheets:
        if worksheet.sheet_state == "visible" and worksheet.title != "WpsReserved_CellImgList":
            print(worksheet.title)
    return 0


def list_products(workbook_path: Path, sheet_name: str) -> int:
    # Read cached formula results where available so a product/image formula is
    # not shown to the designer as the literal formula text.
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if sheet_name not in workbook.sheetnames:
        print(f"Sheet 不存在：{sheet_name}", file=sys.stderr)
        return 2
    worksheet = workbook[sheet_name]
    if worksheet.sheet_state != "visible" or worksheet.title == "WpsReserved_CellImgList":
        print(f"Sheet 不可处理：{sheet_name}", file=sys.stderr)
        return 2

    seen: set[str] = set()
    if is_vertical_layout(worksheet):
        product_values = (worksheet.cell(row, 1).value for row in range(2, worksheet.max_row + 1))
    else:
        product_values = (worksheet.cell(1, column).value for column in range(2, worksheet.max_column + 1))
    for value in product_values:
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
