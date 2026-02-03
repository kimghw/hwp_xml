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

    # YAML 설정 사용
    python merge_with_review.py -o output.hwpx --config formatter_config.yaml file1.hwpx file2.hwpx

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

# Windows 콘솔 인코딩 설정 (run_merge.py main에서 처리)

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
    FormatValidator, ValidationResult,
    validate_and_fix, print_validation_result
)

# formatters 모듈 (통합)
from .formatters import (
    BaseFormatter,
    BulletFormatter,
    CaptionFormatter,
    BULLET_STYLES,
    load_config,
    FormatterConfig,
)


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

    def __init__(
        self,
        config: Optional[FormatterConfig] = None,
        outline_formatter: Optional[BaseFormatter] = None,
        add_formatter: Optional[BaseFormatter] = None,
    ):
        """
        Args:
            config: FormatterConfig 객체 (YAML 설정 사용 시)
            outline_formatter: 개요 본문용 포맷터 (BaseFormatter 상속)
            add_formatter: add_ 필드용 포맷터 (BaseFormatter 상속)
        """
        self.config = config

        # 스타일 설정 (YAML 설정 또는 기본값)
        if config is not None:
            self.bullet_style_name = config.bullet.style
            self.bullet_styles = self._extract_bullet_styles_from_config(config)
            self.caption_styles = {
                "table": f"{config.table_caption.type_prefix} {{num}}{config.table_caption.separator}{{title}}",
                "figure": f"{config.image_caption.type_prefix} {{num}}{config.image_caption.separator}{{title}}",
            }
        else:
            self.bullet_style_name = "filled"
            self.bullet_styles = self._extract_bullet_styles_from_name("filled")
            self.caption_styles = {
                "table": "표 {num}. {title}",
                "figure": "그림 {num}. {title}",
            }

        # 포맷터 생성 (formatters 모듈 사용)
        default_formatter = BulletFormatter(style=self.bullet_style_name)
        self._caption_formatter = CaptionFormatter()

        # 개요/add_ 별도 포맷터 (지정 안하면 기본 포맷터 사용)
        self._outline_formatter = outline_formatter or default_formatter
        self._add_formatter = add_formatter or default_formatter
        self._bullet_formatter = default_formatter  # 기존 호환성 유지

        print(f"    [BulletFormatter] style: {self.bullet_style_name}")
        if outline_formatter:
            print(f"    [개요용] {outline_formatter.get_style_name()} 포맷터")
        if add_formatter:
            print(f"    [add_용] {add_formatter.get_style_name()} 포맷터")

        self.parser = HwpxParser()
        self.validator = FormatValidator(self.caption_styles, self.bullet_styles)

    def _extract_bullet_styles_from_name(self, style_name: str) -> Dict[int, str]:
        """스타일 이름에서 글머리 스타일 추출"""
        if style_name in BULLET_STYLES:
            style_config = BULLET_STYLES[style_name]
            # symbol_tuple = (symbol, indent) → indent + symbol 형태로 반환
            return {
                level: symbol_tuple[1] + symbol_tuple[0]
                for level, symbol_tuple in style_config.items()
            }
        # 기본값: filled 스타일
        return self._extract_bullet_styles_from_name("filled")

    def _extract_bullet_styles_from_config(self, config: FormatterConfig) -> Dict[int, str]:
        """FormatterConfig에서 글머리 스타일 추출"""
        return self._extract_bullet_styles_from_name(config.bullet.style)

    @classmethod
    def from_config(
        cls,
        config: FormatterConfig,
        outline_formatter: Optional[BaseFormatter] = None,
        add_formatter: Optional[BaseFormatter] = None,
    ) -> 'MergeReviewPipeline':
        """
        YAML 설정에서 파이프라인 생성

        Args:
            config: FormatterConfig 객체
            outline_formatter: 개요 본문용 포맷터 (BaseFormatter 상속)
            add_formatter: add_ 필드용 포맷터 (BaseFormatter 상속)

        Returns:
            MergeReviewPipeline 인스턴스

        사용 예:
            from merge.formatters import load_config, BulletFormatter
            config = load_config("formatter_config.yaml")

            # 개요와 add_에 다른 스타일 적용
            pipeline = MergeReviewPipeline.from_config(
                config,
                outline_formatter=BulletFormatter(style="default"),
                add_formatter=BulletFormatter(style="filled"),
            )
        """
        return cls(
            config=config,
            outline_formatter=outline_formatter,
            add_formatter=add_formatter,
        )

    @classmethod
    def from_config_file(cls, config_path: Union[str, Path]) -> 'MergeReviewPipeline':
        """
        YAML 설정 파일에서 파이프라인 생성

        Args:
            config_path: YAML 설정 파일 경로

        Returns:
            MergeReviewPipeline 인스턴스
        """
        config = load_config(str(config_path))
        return cls.from_config(config)

    def _fix_bullets_in_tree(self, outline_tree, parent_level: int = -1) -> List[Dict]:
        """
        개요 트리에서 글머리 기호 수정 (SDK 레벨 분석 사용)

        처리 순서:
        1. 기존 글머리가 YAML 설정과 일치하면 패스
        2. 기존 글머리가 다르거나 없으면 SDK로 레벨 분석
        3. 분석된 레벨에 맞는 글머리 적용
        """
        fixes = []

        # YAML 설정에서 유효한 글머리 기호 (문자열 형태)
        valid_bullets = list(self.bullet_styles.values())

        for node in outline_tree:
            # 현재 노드의 상대적 깊이 계산 (부모 기준)
            if parent_level < 0:
                relative_depth = 0
            else:
                relative_depth = node.level - parent_level
            relative_depth = max(0, min(relative_depth, 2))

            # 문단별로 처리
            paras_needing_analysis = []  # SDK 분석이 필요한 문단들

            for i, para in enumerate(node.paragraphs):
                if not para.text:
                    continue

                # 개요 제목(첫 번째 문단, is_outline=True)은 글머리 수정 스킵
                if i == 0 and para.is_outline:
                    continue

                text = para.text.strip()

                # 1. 기존 글머리가 YAML 설정과 일치하는지 확인
                has_valid_bullet = False
                for bullet in valid_bullets:
                    bullet_stripped = bullet.strip()
                    if text.startswith(bullet_stripped):
                        has_valid_bullet = True
                        break

                if has_valid_bullet:
                    # 이미 유효한 글머리 → 스킵
                    continue
                else:
                    # SDK 분석 필요
                    paras_needing_analysis.append((i, para))

            # 2. 레벨 분석 (자동 감지 사용)
            if paras_needing_analysis:
                texts_to_analyze = [para.text for _, para in paras_needing_analysis]
                # 레벨 자동 감지
                analyzed_levels = [
                    self._bullet_formatter._detect_bullet_level(t.strip())
                    if hasattr(self._bullet_formatter, '_detect_bullet_level')
                    else relative_depth
                    for t in texts_to_analyze
                ]
                # 감지 실패(-1)인 경우 기본값 사용
                analyzed_levels = [l if l >= 0 else relative_depth for l in analyzed_levels]
                stripped_texts = [t.strip() for t in texts_to_analyze]
            else:
                analyzed_levels = []
                stripped_texts = []

            # 3. 분석된 레벨에 맞는 글머리 적용 (outline_formatter 사용)
            if stripped_texts and analyzed_levels:
                # outline_formatter.format_with_levels 사용
                format_result = self._outline_formatter.format_with_levels(
                    stripped_texts, analyzed_levels
                )
                formatted_lines = format_result.formatted_text.split('\n') if format_result.success else []
            else:
                formatted_lines = []

            for idx, (para_idx, para) in enumerate(paras_needing_analysis):
                if idx < len(formatted_lines) and formatted_lines[idx]:
                    new_text = formatted_lines[idx]
                else:
                    # fallback: 기존 방식
                    if idx < len(analyzed_levels):
                        level = analyzed_levels[idx]
                    else:
                        level = relative_depth
                    level = max(0, min(level, 2))
                    expected_bullet = self.bullet_styles.get(level, '- ')
                    clean_text = stripped_texts[idx] if idx < len(stripped_texts) else para.text.strip()
                    new_text = expected_bullet + clean_text

                # 원본과 다르면 수정 기록
                if para.text.strip() != new_text.strip():
                    # 레벨 정보 추출
                    fix_level = analyzed_levels[idx] if idx < len(analyzed_levels) else relative_depth
                    fix_level = max(0, min(fix_level, 2))
                    fixes.append({
                        'type': 'bullet_fix',
                        'para_index': para.index,
                        'level': fix_level,
                        'original_text': para.text,
                        'new_text': new_text,
                        'formatter': self._outline_formatter.get_style_name(),
                    })
                    para.text = new_text
                    self._update_element_text(para.element, new_text)

            # 하위 개요 재귀 처리
            if node.children:
                child_fixes = self._fix_bullets_in_tree(node.children, node.level)
                fixes.extend(child_fixes)

        return fixes

    def _fix_captions(self, paragraphs: list) -> List[Dict]:
        """캡션 형식 수정"""
        import re
        fixes = []

        for para in paragraphs:
            if not para.text:
                continue

            text = para.text

            # 테이블 캡션 수정
            table_match = re.match(r'^(?:표|Table|\[표)\s*(\d+)[.\s\]]*(.*)$', text, re.IGNORECASE)
            if table_match:
                table_num = int(table_match.group(1))
                title = table_match.group(2).strip()
                new_text = f"표 {table_num}. {title}"

                if text != new_text:
                    fixes.append({
                        'type': 'caption_fix',
                        'para_index': para.index,
                        'caption_type': 'table',
                        'original': text,
                        'new': new_text,
                    })
                    para.text = new_text
                    self._update_element_text(para.element, new_text)
                continue

            # 그림 캡션 수정
            figure_match = re.match(r'^(?:그림|Figure|\[그림)\s*(\d+)[.\s\]]*(.*)$', text, re.IGNORECASE)
            if figure_match:
                figure_num = int(figure_match.group(1))
                title = figure_match.group(2).strip()
                new_text = f"그림 {figure_num}. {title}"

                if text != new_text:
                    fixes.append({
                        'type': 'caption_fix',
                        'para_index': para.index,
                        'caption_type': 'figure',
                        'original': text,
                        'new': new_text,
                    })
                    para.text = new_text
                    self._update_element_text(para.element, new_text)

        return fixes

    def _update_element_text(self, elem, new_text: str):
        """문단 요소의 텍스트를 새 텍스트로 교체"""
        if elem is None:
            return

        # 첫 번째 run의 t 요소 찾아서 텍스트 교체
        for run in elem:
            if run.tag.endswith('}run'):
                t_elements = [t for t in run if t.tag.endswith('}t')]
                if t_elements:
                    t_elements[0].text = new_text
                    for t in t_elements[1:]:
                        run.remove(t)
                    return

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
                # 글머리 기호 수정 (SDK 레벨 분석 + 정규식 적용)
                bullet_fixes = self._fix_bullets_in_tree(merged_tree)
                result.fixes_applied.extend(bullet_fixes)

                # 캡션 형식 수정
                merged_paragraphs = flatten_outline_tree(merged_tree)
                caption_fixes = self._fix_captions(merged_paragraphs)
                result.fixes_applied.extend(caption_fixes)

                if bullet_fixes or caption_fixes:
                    print(f"    - 글머리 기호 수정: {len(bullet_fixes)}건")
                    print(f"    - 캡션 수정: {len(caption_fixes)}건")

            result.review_success = True

            # 4. 병합된 파일 생성 (수정된 트리 사용)
            print("[4/4] 병합 파일 생성 중...")
            # format_content=False: 이미 _fix_bullets_in_tree에서 글머리 적용했으므로 중복 방지
            # add_formatter: 테이블 add_ 필드용 포맷터 전달
            merger = HwpxMerger(format_content=False, add_formatter=self._add_formatter)
            for data in hwpx_data_list:
                merger.hwpx_data_list.append(data)

            # 수정된 merged_tree로 병합 (element 수정이 반영됨)
            merger.merge_with_tree(output_path, merged_tree)
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
