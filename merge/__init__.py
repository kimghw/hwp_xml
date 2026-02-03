# -*- coding: utf-8 -*-
"""
HWPX 병합 모듈

개요 기준 파일 병합 및 테이블 셀 병합 기능을 제공합니다.

사용 예:
    # 파일 병합
    from merge import merge_hwpx_files, get_outline_structure
    merge_hwpx_files(["file1.hwpx", "file2.hwpx"], "output.hwpx")

    # 테이블 셀 병합
    from merge import TableMerger
    merger = TableMerger()
    merger.load_base_table("input.hwpx", table_index=0)
    merger.add_rows_smart(data_list)
    merger.save("output.hwpx")
"""

# 데이터 모델
from .models import Paragraph, OutlineNode, HwpxData
from .table import CellInfo, HeaderConfig, RowAddPlan, TableInfo

# 파서
from .parser import HwpxParser
from .table import TableParser

# 개요 트리 함수
from .outline import (
    build_outline_tree,
    merge_outline_trees,
    flatten_outline_tree,
    filter_outline_tree,
    get_all_outline_names,
    print_outline_tree,
)

# 병합 클래스 및 함수
from .merge_hwpx import (
    HwpxMerger,
    get_outline_structure,
    merge_hwpx_files,
)

from .table import TableMerger

# 검증 및 수정
from .format_validator import (
    FormatValidator,
    FormatFixer,
    ValidationResult,
    validate_and_fix,
    print_validation_result,
)

# 병합 파이프라인
from .merge_pipeline import MergePipeline, MergeResult
from .merge_table import TableMergePlan

# 내용 양식 변환
from .content_formatter import ContentFormatter, OutlineContentFormatter

# 정규식 기반 포맷터
from .formatters import (
    BulletFormatter,
    BulletItem,
    FormatResult,
    BULLET_STYLES,
    CaptionFormatter,
    CaptionInfo,
    CAPTION_PATTERNS,
    get_captions,
    print_captions,
    renumber_captions,
)

__all__ = [
    # 데이터 모델
    'Paragraph',
    'OutlineNode',
    'HwpxData',
    'CellInfo',
    'HeaderConfig',
    'RowAddPlan',
    'TableInfo',

    # 파서
    'HwpxParser',
    'TableParser',

    # 개요 트리 함수
    'build_outline_tree',
    'merge_outline_trees',
    'flatten_outline_tree',
    'filter_outline_tree',
    'get_all_outline_names',
    'print_outline_tree',

    # 병합 클래스 및 함수
    'HwpxMerger',
    'TableMerger',
    'MergePipeline',
    'MergeResult',
    'TableMergePlan',
    'get_outline_structure',
    'merge_hwpx_files',

    # 검증 및 수정
    'FormatValidator',
    'FormatFixer',
    'ValidationResult',
    'validate_and_fix',
    'print_validation_result',

    # 내용 양식 변환
    'ContentFormatter',
    'OutlineContentFormatter',

    # 정규식 기반 포맷터 (merge.formatters)
    'BulletFormatter',
    'BulletItem',
    'FormatResult',
    'BULLET_STYLES',
    'CaptionFormatter',
    'CaptionInfo',
    'CAPTION_PATTERNS',
    'get_captions',
    'print_captions',
    'renumber_captions',
]
