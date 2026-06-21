from __future__ import annotations

import argparse
from pathlib import Path

from .converter import convert_jobs, discover_jobs


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pdf2office", description="把 PDF 文件转换为 Word、PPT 或 Excel 文件")
    parser.add_argument("input", type=Path, help="PDF 文件或包含 PDF 的目录")
    parser.add_argument("-o", "--output", type=Path, help="输出文件或输出目录")
    parser.add_argument(
        "-f",
        "--format",
        default="word",
        choices=["word", "docx", "ppt", "pptx", "excel", "xlsx"],
        help="输出格式，默认 word",
    )
    parser.add_argument("-r", "--recursive", action="store_true", help="输入为目录时递归扫描子目录")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    parser.add_argument("--start", type=int, help="起始页码，从 1 开始")
    parser.add_argument("--end", type=int, help="结束页码，从 1 开始，包含该页")
    return parser


def validate_page_range(start: int | None, end: int | None) -> None:
    if start is not None and start < 1:
        raise ValueError("--start 必须大于等于 1")
    if end is not None and end < 1:
        raise ValueError("--end 必须大于等于 1")
    if start is not None and end is not None and end < start:
        raise ValueError("--end 不能小于 --start")


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        validate_page_range(args.start, args.end)
        jobs = discover_jobs(args.input, args.output, args.recursive, args.format)
    except Exception as exc:
        parser.error(str(exc))

    if not jobs:
        print("没有找到 PDF 文件")
        return 1

    results = convert_jobs(jobs, overwrite=args.overwrite, start_page=args.start, end_page=args.end)
    for result in results:
        if result.status == "converted":
            print(f"[成功] {result.source} -> {result.target}")
        elif result.status == "skipped":
            print(f"[跳过] {result.source}：{result.message}")
        else:
            print(f"[失败] {result.source}：{result.message}")

    return 1 if any(result.status == "failed" for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
