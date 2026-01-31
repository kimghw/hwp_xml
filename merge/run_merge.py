# -*- coding: utf-8 -*-
"""
HWPX 병합 CLI

HWPX 파일을 병합하고 형식을 검토/수정합니다.

사용법:
    # 기본 병합 (형식 검토 포함)
    python -m merge.run_merge -o output.hwpx template.hwpx addition.hwpx

    # 단순 병합 (형식 검토 없음)
    python -m merge.run_merge -o output.hwpx --simple template.hwpx addition.hwpx

    # YAML 설정 사용
    python -m merge.run_merge -o output.hwpx --config formatter_config.yaml template.hwpx addition.hwpx

    # 개요 구조만 출력
    python -m merge.run_merge --list-outlines template.hwpx addition.hwpx

    # 자동 수정 비활성화 (검토만)
    python -m merge.run_merge -o output.hwpx --no-fix template.hwpx addition.hwpx

    # SDK 비활성화 (정규식만 사용)
    python -m merge.run_merge -o output.hwpx --no-sdk template.hwpx addition.hwpx
"""

import argparse
import sys
import io
from pathlib import Path
from typing import List, Optional

# Windows 콘솔 인코딩 설정 (main 함수에서 처리)

# 로컬 모듈
from .merge_hwpx import HwpxMerger, get_outline_structure
from .outline import print_outline_tree
from .merge_with_review import MergeReviewPipeline, print_validation_result

# formatters 모듈
try:
    from .formatters import load_config
    HAS_FORMATTERS = True
except ImportError:
    HAS_FORMATTERS = False

# 기본 설정 파일 경로
DEFAULT_CONFIG_PATH = Path(__file__).parent / "formatters" / "formatter_config.yaml"


def print_outline_structures(hwpx_paths: List[str]):
    """파일들의 개요 구조 출력"""
    for i, path in enumerate(hwpx_paths):
        print(f"\n[파일 {i+1}] {path}")
        print("-" * 40)
        tree = get_outline_structure(path)
        print_outline_tree(tree)


def merge_simple(hwpx_paths: List[str], output_path: str, exclude: Optional[List[str]] = None):
    """단순 병합 (형식 검토 없음)"""
    print("=" * 60)
    print("HWPX 단순 병합")
    print("=" * 60)
    print(f"입력 파일: {len(hwpx_paths)}개")
    print(f"출력 파일: {output_path}")

    merger = HwpxMerger()
    for path in hwpx_paths:
        merger.add_file(path)

    if exclude:
        merger.merge(output_path, exclude_outlines=exclude)
    else:
        merger.merge(output_path)

    print(f"\n[OK] 병합 완료: {output_path}")


def merge_with_review(
    hwpx_paths: List[str],
    output_path: str,
    config_path: Optional[str] = None,
    auto_fix: bool = True,
    use_sdk: bool = True
):
    """형식 검토 포함 병합 (SDK 기반)"""
    print("=" * 60)
    print("HWPX 병합 + 형식 검토" + (" [SDK]" if use_sdk else " [정규식]"))
    print("=" * 60)
    print(f"입력 파일: {len(hwpx_paths)}개")
    print(f"출력 파일: {output_path}")
    print(f"자동 수정: {'활성화' if auto_fix else '비활성화'}")
    print(f"SDK 사용: {'예' if use_sdk else '아니오'}")

    # 설정 파일 결정 (지정 없으면 기본 설정 파일 사용)
    effective_config_path = config_path
    if not effective_config_path and HAS_FORMATTERS and DEFAULT_CONFIG_PATH.exists():
        effective_config_path = str(DEFAULT_CONFIG_PATH)
        print(f"설정 파일: {effective_config_path} (기본)")
    elif effective_config_path:
        print(f"설정 파일: {effective_config_path}")

    print("=" * 60)

    # 파이프라인 생성
    if effective_config_path and HAS_FORMATTERS:
        pipeline = MergeReviewPipeline.from_config_file(effective_config_path, use_sdk=use_sdk)
    else:
        pipeline = MergeReviewPipeline(use_sdk=use_sdk)

    # 각 파일 구조 출력
    for i, path in enumerate(hwpx_paths):
        print(f"\n[파일 {i+1}] {path}")
        print("-" * 40)
        tree = get_outline_structure(path)
        print_outline_tree(tree)

    print("\n" + "=" * 60)

    # 병합 실행
    result = pipeline.merge_and_review(hwpx_paths, output_path, auto_fix=auto_fix)

    # 결과 출력
    print("\n" + "=" * 60)
    print("결과")
    print("=" * 60)
    print(f"병합 성공: {'[OK]' if result.merge_success else '[FAIL]'}")
    print(f"검토 성공: {'[OK]' if result.review_success else '[FAIL]'}")
    print(f"수정 적용: {len(result.fixes_applied)}건")

    if result.fixes_applied:
        print("\n적용된 수정:")
        for fix in result.fixes_applied[:10]:
            fix_type = fix.get('type', 'unknown')
            if fix_type == 'bullet_fix':
                print(f"  - 글머리: {fix.get('original_bullet')} → {fix.get('new_bullet')}")
            elif fix_type == 'caption_fix':
                print(f"  - 캡션: {fix.get('original', '')[:30]}...")
            else:
                print(f"  - {fix_type}: {str(fix)[:50]}...")

        if len(result.fixes_applied) > 10:
            print(f"  ... 외 {len(result.fixes_applied) - 10}건")

    print_validation_result(result.validation)

    return result


def main():
    # Windows 콘솔 인코딩 설정
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

    parser = argparse.ArgumentParser(
        description="HWPX 병합 CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  # 기본 병합
  python -m merge.run_merge -o output.hwpx template.hwpx addition.hwpx

  # YAML 설정 사용
  python -m merge.run_merge -o output.hwpx --config config.yaml file1.hwpx file2.hwpx

  # SDK 비활성화 (정규식 기반 포맷터만 사용)
  python -m merge.run_merge -o output.hwpx --no-sdk file1.hwpx file2.hwpx

  # 개요 구조만 확인
  python -m merge.run_merge --list-outlines file1.hwpx file2.hwpx
"""
    )

    parser.add_argument("files", nargs="+", help="병합할 HWPX 파일들 (첫 번째: template, 나머지: addition)")
    parser.add_argument("-o", "--output", help="출력 파일 경로")
    parser.add_argument("--config", help="YAML 설정 파일 경로")
    parser.add_argument("--simple", action="store_true", help="단순 병합 (형식 검토 없음)")
    parser.add_argument("--no-fix", action="store_true", help="자동 수정 비활성화 (검토만)")
    parser.add_argument("--no-sdk", action="store_true", help="SDK 비활성화 (정규식 기반 포맷터만 사용)")
    parser.add_argument("--list-outlines", action="store_true", help="개요 구조만 출력")
    parser.add_argument("--exclude", nargs="*", help="제외할 개요 이름")

    args = parser.parse_args()

    # 개요 구조 출력 모드
    if args.list_outlines:
        print_outline_structures(args.files)
        return

    # 출력 파일 필수
    if not args.output:
        parser.error("출력 파일을 지정해주세요 (-o/--output)")

    # 파일 개수 확인
    if len(args.files) < 2:
        parser.error("최소 2개의 HWPX 파일이 필요합니다 (template + addition)")

    # 파일 존재 확인
    for f in args.files:
        if not Path(f).exists():
            print(f"[ERROR] 파일을 찾을 수 없습니다: {f}", file=sys.stderr)
            sys.exit(1)

    # 병합 실행
    if args.simple:
        merge_simple(args.files, args.output, exclude=args.exclude)
    else:
        result = merge_with_review(
            args.files,
            args.output,
            config_path=args.config,
            auto_fix=not args.no_fix,
            use_sdk=not args.no_sdk
        )
        if not result.merge_success:
            sys.exit(1)


if __name__ == "__main__":
    main()
