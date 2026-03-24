#!/usr/bin/env python3
"""
从 .xlsx 中导出内嵌图片（解压 zip 内的 xl/media/）。

用法:
  python3 export_xlsx_media.py /path/to/脚本范例.xlsx
  python3 export_xlsx_media.py /path/to/脚本范例.xlsx -o ~/Desktop/脚本范例_内嵌图

默认输出目录: 与 xlsx 同目录下的「<文件名>_media」文件夹。
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

MEDIA_PREFIX = "xl/media/"


def export_media(xlsx_path: Path, out_dir: Path) -> list[Path]:
    xlsx_path = xlsx_path.expanduser().resolve()
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"找不到文件: {xlsx_path}")
    if xlsx_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        print("提示：标准 Office xlsx/xlsm 为 zip；其他格式可能无效。", file=sys.stderr)

    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        names = [n for n in zf.namelist() if n.startswith(MEDIA_PREFIX) and not n.endswith("/")]
        if not names:
            print("未在压缩包内发现 xl/media/ 下的文件（该表可能没有内嵌图）。", file=sys.stderr)
            return written

        for name in sorted(names):
            # xl/media/image1.png -> image1.png
            base = name[len(MEDIA_PREFIX) :].replace("/", "_")
            if not base:
                continue
            dest = out_dir / base
            # 避免子目录名冲突：若 base 含路径分隔（少见），已 replace
            with zf.open(name) as src, open(dest, "wb") as dst:
                shutil.copyfileobj(src, dst)
            written.append(dest)

    return written


def main() -> None:
    p = argparse.ArgumentParser(description="从 xlsx 导出 xl/media 内嵌资源")
    p.add_argument("xlsx", type=Path, help="输入 .xlsx / .xlsm 路径")
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="输出目录（默认: 与 xlsx 同目录下的 <文件名>_media）",
    )
    args = p.parse_args()

    xlsx_path = args.xlsx
    if args.output is not None:
        out_dir = args.output
    else:
        out_dir = xlsx_path.parent / f"{xlsx_path.stem}_media"

    try:
        files = export_media(xlsx_path, out_dir)
    except zipfile.BadZipFile:
        print("错误：不是有效的 zip/xlsx 文件。", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        sys.exit(1)

    print(f"输出目录: {out_dir}")
    print(f"共导出 {len(files)} 个文件")
    for f in files:
        print(f"  {f.name}")


if __name__ == "__main__":
    main()
