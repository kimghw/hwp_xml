# -*- coding: utf-8 -*-
"""
HWPX 병합 + 형식 검토 통합 모듈

Claude Agent SDK Sub Agent를 사용하여:
1. HWPX 파일 병합
2. 형식 검토 (캡션 스타일, 글머리 기호)
3. 자동 수정 후 최종 병합

사용법:
    # 동기 방식
    python merge_with_review.py -o output.hwpx file1.hwpx file2.hwpx

    # Claude Agent SDK 사용 (비동기)
    from merge_with_review import merge_with_review_async
    await merge_with_review_async([...], output_path, use_agent=True)
"""

import asyncio
import json
import sys
import io
from pathlib import Path
from typing import List, Dict, Optional, Union
from dataclasses import dataclass, field

# Windows 콘솔 인코딩 설정
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 로컬 모듈
from .merge_hwpx import (
    HwpxMerger,
    get_outline_structure,
)
from .parser import HwpxParser
from .models import HwpxData
from .outline import (
    build_outline_tree,
    merge_outline_trees,
    flatten_outline_tree,
    print_outline_tree,
)
from .format_validator import (
    FormatValidator, FormatFixer, ValidationResult,
    validate_and_fix, print_validation_result
)


# 글머리 기호 순서 정의
BULLET_ORDER = {
    0: "■",  # 1단계: 네모
    1: "●",  # 2단계: 원
    2: "-",  # 3단계: 대시
}


@dataclass
class MergeReviewResult:
    """병합 및 검토 결과"""
    output_path: str = ""
    validation: ValidationResult = field(default_factory=ValidationResult)
    fixes_applied: List[Dict] = field(default_factory=list)
    merge_success: bool = False
    review_success: bool = False
    agent_feedback: str = ""


class MergeReviewPipeline:
    """병합 + 검토 파이프라인"""

    def __init__(self, bullet_styles: Dict[int, str] = None, caption_styles: Dict[str, str] = None):
        self.bullet_styles = bullet_styles or BULLET_ORDER
        self.caption_styles = caption_styles or {
            "table": "표 {num}. {title}",
            "figure": "그림 {num}. {title}",
        }
        self.parser = HwpxParser()
        self.validator = FormatValidator(self.caption_styles, self.bullet_styles)
        self.fixer = FormatFixer(self.bullet_styles)

    def _apply_fixes_to_hwpx_data(self, hwpx_data_list: List[HwpxData], merged_tree):
        """
        수정된 개요 트리의 변경 사항을 원본 HwpxData의 element에 반영

        FormatFixer가 para.text와 para.element를 동시에 수정하므로,
        수정된 element가 원본 데이터에 이미 반영되어 있음.
        이 메서드는 추가 동기화가 필요한 경우를 위해 유지.
        """
        # FormatFixer가 para.element를 직접 수정하므로
        # 별도 동기화 불필요 (element는 참조이므로 원본에 반영됨)
        pass

    def merge_and_review(
        self,
        hwpx_paths: List[Union[str, Path]],
        output_path: Union[str, Path],
        auto_fix: bool = True
    ) -> MergeReviewResult:
        """
        HWPX 파일 병합 후 형식 검토 및 수정

        Args:
            hwpx_paths: 병합할 HWPX 파일 경로들
            output_path: 출력 파일 경로
            auto_fix: 자동 수정 여부

        Returns:
            MergeReviewResult
        """
        result = MergeReviewResult(output_path=str(output_path))

        try:
            # 1. 파일 파싱
            print(f"[1/4] 파일 파싱 중... ({len(hwpx_paths)}개)")
            hwpx_data_list = []
            for path in hwpx_paths:
                data = self.parser.parse(path)
                hwpx_data_list.append(data)

            # 2. 개요 트리 병합
            print("[2/4] 개요 트리 병합 중...")
            trees = [data.outline_tree for data in hwpx_data_list]
            merged_tree = merge_outline_trees(trees)

            # 3. 형식 검토 및 수정
            print("[3/4] 형식 검토 및 수정 중...")
            if auto_fix:
                # 글머리 기호 수정 (텍스트 + XML 요소 동시 수정)
                bullet_fixes = self.fixer.fix_bullets_in_tree(merged_tree)
                result.fixes_applied.extend(bullet_fixes)

                # 캡션 형식 수정 (텍스트 + XML 요소 동시 수정)
                merged_paragraphs = flatten_outline_tree(merged_tree)
                caption_fixes = self.fixer.fix_caption_format(merged_paragraphs)
                result.fixes_applied.extend(caption_fixes)

                if bullet_fixes or caption_fixes:
                    print(f"    - 글머리 기호 수정: {len(bullet_fixes)}건")
                    print(f"    - 캡션 수정: {len(caption_fixes)}건")

                # 수정된 내용을 원본 HwpxData에 반영
                self._apply_fixes_to_hwpx_data(hwpx_data_list, merged_tree)

            result.review_success = True

            # 4. 병합된 파일 생성
            print("[4/4] 병합 파일 생성 중...")
            merger = HwpxMerger()
            for data in hwpx_data_list:
                merger.hwpx_data_list.append(data)

            # 병합 (수정된 element가 반영됨)
            merger.merge(output_path)
            result.merge_success = True

            # 5. 최종 검증
            print("\n최종 검증...")
            result.validation = self.validator.validate(output_path)

            print(f"\n[OK] 병합 완료: {output_path}")

        except Exception as e:
            result.validation.errors.append(f"병합 실패: {str(e)}")
            print(f"\n[ERROR] 오류 발생: {e}")

        return result


# ============================================================
# Claude Agent SDK 통합
# ============================================================

FORMAT_REVIEW_PROMPT = """당신은 HWPX 문서 형식 검토 전문가입니다.

다음 규칙에 따라 문서 형식을 검토하세요:

## 캡션 스타일 규칙
- 테이블: "표 N. 제목" 형식 (예: "표 1. 실험 결과")
- 그림: "그림 N. 제목" 형식 (예: "그림 1. 시스템 구조도")

## 개요 글머리 기호 규칙 (3단계)
- 1단계: ■ (검정 네모)
- 2단계: ● (검정 원)
- 3단계: - (대시)

## 검토 시 확인할 사항
1. 모든 캡션이 표준 형식을 따르는지
2. 글머리 기호가 계층 구조에 맞는지
3. 번호가 연속적인지

검토 결과를 JSON 형식으로 반환하세요:
{
    "is_valid": true/false,
    "errors": [...],
    "warnings": [...],
    "suggestions": [...]
}"""


async def merge_with_review_async(
    hwpx_paths: List[Union[str, Path]],
    output_path: Union[str, Path],
    use_agent: bool = True
) -> MergeReviewResult:
    """
    Claude Agent SDK를 사용한 비동기 병합 + 검토

    Args:
        hwpx_paths: 병합할 HWPX 파일 경로들
        output_path: 출력 파일 경로
        use_agent: Claude Agent 사용 여부

    Returns:
        MergeReviewResult
    """
    pipeline = MergeReviewPipeline()

    if not use_agent:
        return pipeline.merge_and_review(hwpx_paths, output_path, auto_fix=True)

    # SDK import 시도
    try:
        from claude_code_sdk import query, ClaudeCodeOptions
    except ImportError as e:
        print(f"경고: claude-code-sdk import 실패 ({e}). 동기 방식으로 진행합니다.")
        return pipeline.merge_and_review(hwpx_paths, output_path, auto_fix=True)

    result = MergeReviewResult(output_path=str(output_path))

    try:
        # Agent 옵션 설정
        options = ClaudeCodeOptions(
            allowed_tools=["Read", "Grep", "Glob"],
            cwd=str(Path(hwpx_paths[0]).parent),
        )

        # 1단계: Agent로 형식 검토 요청
        print("[Agent] 형식 검토 시작...")

        file_list = '\n'.join(f'- {p}' for p in hwpx_paths)
        review_prompt = f"""다음 HWPX 파일들의 형식을 검토해주세요:
{file_list}

{FORMAT_REVIEW_PROMPT}
"""

        agent_responses = []
        async for message in query(prompt=review_prompt, options=options):
            if hasattr(message, 'content'):
                for block in message.content:
                    if hasattr(block, 'text'):
                        agent_responses.append(block.text)
            elif hasattr(message, 'result'):
                agent_responses.append(str(message.result))

        result.agent_feedback = '\n'.join(agent_responses)
        if result.agent_feedback:
            print(f"[Agent 검토 결과]\n{result.agent_feedback[:500]}...")

    except Exception as e:
        print(f"[Agent 오류] {e}")
        result.agent_feedback = f"Agent 실행 중 오류: {e}"

    # 2단계: 로컬 병합 + 수정
    print("\n[로컬] 병합 및 수정 시작...")
    merge_result = pipeline.merge_and_review(hwpx_paths, output_path, auto_fix=True)

    result.validation = merge_result.validation
    result.fixes_applied = merge_result.fixes_applied
    result.merge_success = merge_result.merge_success
    result.review_success = merge_result.review_success

    return result


# ============================================================
# CLI
# ============================================================

def main():
    import argparse

    parser = argparse.ArgumentParser(description="HWPX 병합 + 형식 검토")
    parser.add_argument("-o", "--output", required=True, help="출력 파일 경로")
    parser.add_argument("files", nargs="+", help="병합할 HWPX 파일들")
    parser.add_argument("--no-fix", action="store_true", help="자동 수정 비활성화")
    parser.add_argument("--agent", action="store_true", help="Claude Agent 사용")

    args = parser.parse_args()

    print("=" * 60)
    print("HWPX 병합 + 형식 검토")
    print("=" * 60)
    print(f"입력 파일: {len(args.files)}개")
    print(f"출력 파일: {args.output}")
    print(f"자동 수정: {'비활성화' if args.no_fix else '활성화'}")
    print(f"Agent 사용: {'예' if args.agent else '아니오'}")
    print("=" * 60)

    # 각 파일 구조 출력
    for i, path in enumerate(args.files):
        print(f"\n[파일 {i+1}] {path}")
        print("-" * 40)
        tree = get_outline_structure(path)
        print_outline_tree(tree)

    print("\n" + "=" * 60)

    if args.agent:
        # 비동기 실행
        result = asyncio.run(
            merge_with_review_async(args.files, args.output, use_agent=True)
        )
    else:
        # 동기 실행
        pipeline = MergeReviewPipeline()
        result = pipeline.merge_and_review(args.files, args.output, auto_fix=not args.no_fix)

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
            print(f"  - {fix.get('type', 'unknown')}: {fix.get('original', '')[:30]}...")

    print_validation_result(result.validation)

    if result.agent_feedback:
        print("\n[Agent 피드백]")
        print(result.agent_feedback[:1000])


if __name__ == "__main__":
    main()
