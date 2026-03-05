#!/usr/bin/env python3
"""hwpx-gen — JSON/Markdown to HWPX document converter.

Usage examples:
    python cli.py --input notice.json --output notice.hwpx
    python cli.py --input notice.md   --output notice.hwpx
    python cli.py --input notice.json --output notice.hwpx --preset gov_document
    python cli.py -i spec.json -o result.hwpx --no-page-flow
    python cli.py -i report.md -o report.hwpx --style 외부 --org 경기도교육청
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# Allow running from the repository root without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from hwpx_generator.json_loader import load_from_file as load_json, _build_preset
from hwpx_generator.md_loader import load_from_md_file

# 외부 스타일 동적 빌더
sys.path.insert(
    0,
    str(Path(__file__).resolve().parent / "gonggong_hwpxskills-main" / "gonggong_hwpxskills-main" / "scripts"),
)
import dynamic_builder as _dyn


# Supported presets (shown in --help)
_AVAILABLE_PRESETS = ["gov_document"]


def _run_external_style(
    input_path: str,
    output_path: str,
    org: str,
    color_main: str = "1F4E79",
    color_sub: str = "2E75B6",
    color_accent: str = "C00000",
) -> None:
    """외부 스타일(동적 빌더) 변환 실행."""
    _dyn.run(
        input_path=input_path,
        output_path=output_path,
        org=org,
        color_main=color_main,
        color_sub=color_sub,
        color_accent=color_accent,
    )


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="hwpx-gen",
        description="JSON 또는 Markdown 문서를 HWPX(한글) 파일로 변환합니다.",
        epilog=(
            "examples:\n"
            "  python cli.py -i notice.json -o notice.hwpx\n"
            "  python cli.py -i notice.md   -o notice.hwpx\n"
            "  python cli.py -i notice.json -o notice.hwpx --preset gov_document\n"
            "  python cli.py -i notice.md   -o notice.hwpx --preset gov_document\n"
            "  python cli.py -i report.md -o report.hwpx --style 외부 --org 경기도교육청\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        metavar="FILE",
        help="입력 파일 경로 (.json 또는 .md). 확장자로 포맷을 자동 감지합니다.",
    )
    parser.add_argument(
        "--output", "-o",
        required=True,
        metavar="FILE",
        help="출력 .hwpx 파일 경로.",
    )
    parser.add_argument(
        "--preset", "-p",
        default=None,
        choices=_AVAILABLE_PRESETS,
        help=(
            "스타일 프리셋 적용. "
            f"사용 가능: {', '.join(_AVAILABLE_PRESETS)}. "
            "JSON 파일에 preset 키가 이미 있으면 이 옵션은 무시됩니다."
        ),
    )
    parser.add_argument(
        "--no-page-flow",
        action="store_true",
        default=False,
        help="자동 페이지 흐름(페이지 나누기 삽입)을 비활성화합니다.",
    )
    parser.add_argument(
        "--style", "-s",
        default="기존",
        choices=["기존", "외부"],
        help="변환 방식 선택. 기존: hwpx_generator, 외부: gonggong 템플릿 (기본값: 기존)",
    )
    parser.add_argument(
        "--org",
        default="",
        help="보고서 기관명 (--style 외부 일 때만 사용, 기본값: 빈 문자열)",
    )
    parser.add_argument(
        "--color-main",
        default="1F4E79",
        help="섹션바 배경색 (--style 외부, 기본값: 1F4E79)",
    )
    parser.add_argument(
        "--color-sub",
        default="2E75B6",
        help="□ 기호색 (--style 외부, 기본값: 2E75B6)",
    )
    parser.add_argument(
        "--color-accent",
        default="C00000",
        help="※ 주석색 (--style 외부, 기본값: C00000)",
    )

    args = parser.parse_args(argv)

    input_path: str = args.input
    output_path: str = args.output

    # Validate input exists
    inp = Path(input_path)
    if not inp.exists():
        print(f"Error: 입력 파일을 찾을 수 없습니다: {input_path}", file=sys.stderr)
        sys.exit(1)

    # 외부 스타일 분기
    if args.style == "외부":
        _run_external_style(
            input_path, output_path, args.org,
            color_main=args.color_main,
            color_sub=args.color_sub,
            color_accent=args.color_accent,
        )
        return

    # === 기존 스타일 (아래 코드 변경 없음) ===
    preset_name: str | None = args.preset
    auto_page_flow: bool = not args.no_page_flow
    ext = inp.suffix.lower()

    try:
        if ext == ".md":
            doc = load_from_md_file(input_path, preset_name=preset_name)
        elif ext == ".json":
            doc = load_json(input_path)
            # Apply --preset if JSON doesn't already specify one
            if preset_name and not _json_has_preset(input_path):
                preset = _build_preset(preset_name, {})
                doc.apply_preset(preset)
        else:
            print(
                f"Error: 지원하지 않는 파일 형식입니다: '{ext}'. "
                ".json 또는 .md 파일을 사용하세요.",
                file=sys.stderr,
            )
            sys.exit(1)

        doc.save(output_path, auto_page_flow=auto_page_flow)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    abs_out = str(Path(output_path).resolve())
    print(f"Success: {abs_out}")


def _json_has_preset(path: str) -> bool:
    """Check if the JSON file already contains a 'preset' key."""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return bool(data.get("preset"))
    except Exception:
        return False


if __name__ == "__main__":
    main()
