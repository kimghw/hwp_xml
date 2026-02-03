# -*- coding: utf-8 -*-
"""
테이블 병합 및 필드 관리 모듈

- models: 테이블/셀 데이터 모델
- parser: HWPX 테이블 파싱
- merger: 테이블 셀 내용 병합
- formatter_config: add_ 필드 포맷터 설정
"""

from .models import CellInfo, TableInfo, HeaderConfig, RowAddPlan
from .parser import TableParser
from .merger import TableMerger
from .row_extractor import RowExtractor, extract_table_rows
from .formatter_config import (
    TableFormatterConfigLoader,
    TableFormatterConfig,
    FieldFormatterConfig,
    load_table_formatter_config,
    format_add_field_value,
)

__all__ = [
    'CellInfo',
    'TableInfo',
    'HeaderConfig',
    'RowAddPlan',
    'TableParser',
    'TableMerger',
    # 행 추출
    'RowExtractor',
    'extract_table_rows',
    # 포맷터 설정
    'TableFormatterConfigLoader',
    'TableFormatterConfig',
    'FieldFormatterConfig',
    'load_table_formatter_config',
    'format_add_field_value',
]
