#!/usr/bin/env python3
"""Build a Windows designer ZIP with UTF-8 entry names.

Windows Explorer and Windows 10 tar.exe both preserve the tool's Chinese
entrypoint when the ZIP UTF-8 filename flag is present.  Do not substitute a
platform zip utility here: their filename-encoding defaults differ.
"""
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("L0_Windows命令行版"))
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not (source / "开始套版.cmd").is_file():
        raise SystemExit(f"无效工具目录：{source}")
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(source.rglob("*")):
            if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc":
                archive.write(path, path.relative_to(source).as_posix())


if __name__ == "__main__":
    main()
