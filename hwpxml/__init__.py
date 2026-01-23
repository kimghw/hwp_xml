# -*- coding: utf-8 -*-
"""
xml 모듈 - HWPX XML 파싱

HWPX 파일에서 테이블, 페이지 속성 등을 추출
"""

from .get_table_property import (
    GetTableProperty,
    TableProperty,
    CellProperty,
    extract_tables_from_hwpx,
    extract_table_data_as_list,
)
from .get_page_property import (
    GetPageProperty,
    PageProperty,
    PageSize,
    PageMargin,
    Unit,
    get_page_property,
    get_all_page_properties,
)
from .extract_cell_index import (
    ExtractCellIndex,
    CellIndexInfo,
    DocumentIndexMap,
    extract_indexes_from_hwpx,
    get_index_mapping,
)
from .get_cell_detail import (
    GetCellDetail,
    CellDetail,
    FontInfo,
    ParaInfo,
    BorderInfo,
    get_cell_details,
)

__all__ = [
    # Table
    'GetTableProperty',
    'TableProperty',
    'CellProperty',
    'extract_tables_from_hwpx',
    'extract_table_data_as_list',
    # Page
    'GetPageProperty',
    'PageProperty',
    'PageSize',
    'PageMargin',
    'Unit',
    'get_page_property',
    'get_all_page_properties',
    # Index
    'ExtractCellIndex',
    'CellIndexInfo',
    'DocumentIndexMap',
    'extract_indexes_from_hwpx',
    'get_index_mapping',
    # Cell Detail
    'GetCellDetail',
    'CellDetail',
    'FontInfo',
    'ParaInfo',
    'BorderInfo',
    'get_cell_details',
]
