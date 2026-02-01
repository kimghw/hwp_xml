# -*- coding: utf-8 -*-
"""
HWPX 테이블 필드 관리 모듈

필드명 자동 생성, 빈 셀 필드 채우기, 시각화 기능을 제공합니다.

사용 예:
    # 자동 필드명 생성
    from merge.field import insert_auto_fields
    insert_auto_fields("template.hwpx")

    # 빈 셀에 위 셀 필드명 복사
    from merge.field import fill_empty_fields
    fill_empty_fields("template.hwpx")

    # 필드 시각화 (빨간 배경 / 파란 텍스트)
    from merge.field import highlight_empty_fields, insert_field_text
    highlight_empty_fields("template.hwpx", "output.hwpx")
    insert_field_text("template.hwpx", "output.hwpx")
"""

from .auto_field import (
    AutoFieldInserter,
    insert_auto_fields,
    FieldNameGenerator,
    CellForNaming,
    generate_field_names,
)

from .fill_empty import (
    EmptyFieldFiller,
    fill_empty_fields,
)

from .visualizer import (
    FieldVisualizer,
    highlight_empty_fields,
    insert_field_text,
)

__all__ = [
    # 자동 필드 생성
    'AutoFieldInserter',
    'insert_auto_fields',
    'FieldNameGenerator',
    'CellForNaming',
    'generate_field_names',

    # 빈 셀 필드 채우기
    'EmptyFieldFiller',
    'fill_empty_fields',

    # 시각화
    'FieldVisualizer',
    'highlight_empty_fields',
    'insert_field_text',
]
