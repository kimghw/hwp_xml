# -*- coding: utf-8 -*-
"""
HWPX 병합 파이프라인 모듈

사용법:
    from merge.merge_pipeline import MergePipeline

    pipeline = MergePipeline()
    result = pipeline.merge(["file1.hwpx", "file2.hwpx"], "output.hwpx")
"""

import asyncio
from pathlib import Path
from typing import List, Dict, Optional, Union
from dataclasses import dataclass, field

from .merge_hwpx import HwpxMerger, get_outline_structure
from .merge_table import TableMergeHandler, TableMergePlan
from .parser import HwpxParser
from .models import HwpxData
from .outline import merge_outline_trees, flatten_outline_tree, print_outline_tree
from .format_validator import (
    FormatValidator, ValidationResult, FormatFixer, print_validation_result
)
from .formatters import (
    BaseFormatter, BulletFormatter, CaptionFormatter, load_config, FormatterConfig
)


@dataclass
class MergeResult:
    """병합 결과"""
    output_path: str = ""
    validation: ValidationResult = field(default_factory=ValidationResult)
    fixes_applied: List[Dict] = field(default_factory=list)
    table_merges: List[TableMergePlan] = field(default_factory=list)  # 테이블 병합 계획
    success: bool = False
    agent_feedback: str = ""


class MergePipeline:
    """HWPX 병합 파이프라인"""

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
            self.bullet_styles = BulletFormatter.get_bullet_dict_by_name(config.bullet.style)
            self.caption_styles = {
                "table": f"{config.table_caption.type_prefix} {{num}}{config.table_caption.separator}{{title}}",
                "figure": f"{config.image_caption.type_prefix} {{num}}{config.image_caption.separator}{{title}}",
            }
        else:
            self.bullet_style_name = "filled"
            self.bullet_styles = BulletFormatter.get_bullet_dict_by_name("filled")
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

        print(f"    [BulletFormatter] style: {self.bullet_style_name}")
        if outline_formatter:
            print(f"    [개요용] {outline_formatter.get_style_name()} 포맷터")
        if add_formatter:
            print(f"    [add_용] {add_formatter.get_style_name()} 포맷터")

        self.parser = HwpxParser()
        self.validator = FormatValidator(self.caption_styles, self.bullet_styles)

        # 형식 수정기 (SDK 레벨 분석 + 글머리 적용)
        self._fixer = FormatFixer(bullet_styles=self.bullet_styles)

        # 테이블 병합 처리기
        self._table_handler = TableMergeHandler(
            format_content=True,
            use_sdk_for_levels=True,
            add_formatter=self._add_formatter,
        )

    @classmethod
    def from_config(
        cls,
        config: FormatterConfig,
        outline_formatter: Optional[BaseFormatter] = None,
        add_formatter: Optional[BaseFormatter] = None,
    ) -> 'MergePipeline':
        """YAML 설정에서 파이프라인 생성"""
        return cls(
            config=config,
            outline_formatter=outline_formatter,
            add_formatter=add_formatter,
        )

    @classmethod
    def from_config_file(cls, config_path: Union[str, Path]) -> 'MergePipeline':
        """
        YAML 설정 파일에서 파이프라인 생성

        Args:
            config_path: YAML 설정 파일 경로

        Returns:
            MergePipeline 인스턴스
        """
        config = load_config(str(config_path))
        return cls.from_config(config)

    def merge(
        self,
        hwpx_paths: List[Union[str, Path]],
        output_path: Union[str, Path],
        auto_fix: bool = True
    ) -> MergeResult:
        """
        HWPX 파일 병합 후 형식 검토 및 수정

        Args:
            hwpx_paths: 병합할 HWPX 파일 경로들
            output_path: 출력 파일 경로
            auto_fix: 자동 수정 여부

        Returns:
            MergeResult
        """
        result = MergeResult(output_path=str(output_path))

        try:
            # 1. 파일 파싱
            print(f"[1/5] 파일 파싱 중... ({len(hwpx_paths)}개)")
            hwpx_data_list = []
            for path in hwpx_paths:
                data = self.parser.parse(path)
                hwpx_data_list.append(data)

            # 2. 개요 트리 병합 (메모리)
            print("[2/5] 개요 트리 병합 중...")
            trees = [data.outline_tree for data in hwpx_data_list]
            merged_tree = merge_outline_trees(trees)

            # 3. 형식 검토 및 수정
            print("[3/5] 형식 검토 및 수정 중...")
            if auto_fix:
                # 글머리 기호 수정 (SDK 레벨 분석 + 정규식 적용)
                bullet_fixes = self._fixer.fix_bullets_in_tree(merged_tree)
                result.fixes_applied.extend(bullet_fixes)

                # 캡션 형식 수정
                merged_paragraphs = flatten_outline_tree(merged_tree)
                caption_fixes = self._fixer.fix_caption_format(merged_paragraphs)
                result.fixes_applied.extend(caption_fixes)

                if bullet_fixes or caption_fixes:
                    print(f"    - 글머리 기호 수정: {len(bullet_fixes)}건")
                    print(f"    - 캡션 수정: {len(caption_fixes)}건")

            # 4. 테이블 병합
            print("[4/5] 테이블 병합 중...")
            table_merge_plans = self._table_handler.collect_and_merge(hwpx_data_list, merged_tree)
            result.table_merges = table_merge_plans
            if table_merge_plans:
                total_rows = sum(len(p.addition_data) for p in table_merge_plans)
                print(f"    - {len(table_merge_plans)}개 테이블, {total_rows}행 병합 완료")
            else:
                print("    - 병합할 테이블 없음")

            # 5. 파일 생성 (본문만 처리, 테이블은 이미 병합됨)
            print("[5/5] 파일 생성 중...")
            merger = HwpxMerger(format_content=False, add_formatter=self._add_formatter)
            for data in hwpx_data_list:
                merger.hwpx_data_list.append(data)
            merger.merge_with_tree(output_path, merged_tree)
            result.success = True

            # 최종 검증
            print("\n최종 검증...")
            result.validation = self.validator.validate(output_path)

            print(f"\n[OK] 병합 완료: {output_path}")

        except Exception as e:
            result.validation.errors.append(f"병합 실패: {str(e)}")
            print(f"\n[ERROR] 오류 발생: {e}")

        return result


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
        # 비동기 실행 (agent 모듈 사용)
        try:
            from agent import merge_with_review_async
            result = asyncio.run(
                merge_with_review_async(args.files, args.output, use_agent=True)
            )
        except ImportError as e:
            print(f"경고: agent 모듈 import 실패 ({e}). 동기 방식으로 진행합니다.")
            pipeline = MergePipeline()
            result = pipeline.merge(args.files, args.output, auto_fix=not args.no_fix)
    else:
        # 동기 실행
        pipeline = MergePipeline()
        result = pipeline.merge(args.files, args.output, auto_fix=not args.no_fix)

    # 결과 출력
    print("\n" + "=" * 60)
    print("결과")
    print("=" * 60)
    print(f"병합 성공: {'[OK]' if result.success else '[FAIL]'}")
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
