# -*- coding: utf-8 -*-
"""
HWPX 테이블 관련 데이터 모델

개요:
- CellInfo: 테이블 셀 정보
- HeaderConfig: 새 행 추가 시 헤더 열 설정
- RowAddPlan: 행 추가 계획
- TableInfo: 테이블 정보
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any


@dataclass
class CellInfo:
    """테이블 셀 정보"""
    row: int = 0
    col: int = 0
    row_span: int = 1
    col_span: int = 1
    end_row: int = 0
    end_col: int = 0

    # 내용
    text: str = ""
    field_name: str = ""  # nc.name 필드명

    # XML 요소 참조
    element: Any = None

    # 빈 셀 여부
    is_empty: bool = True

    # 셀 크기 (HWPUNIT)
    width: int = 0
    height: int = 0

    # 배경색 (자동 필드명 생성에 사용)
    bg_color: str = ""

    # 셀이 차지하는 영역 (rowspan/colspan으로 확장된 영역)
    def covers(self, row: int, col: int) -> bool:
        """특정 (row, col) 위치를 이 셀이 커버하는지 확인"""
        return (self.row <= row <= self.end_row and
                self.col <= col <= self.end_col)


@dataclass
class HeaderConfig:
    """
    새 행 추가 시 헤더 열 설정

    예시:
    - col=0, action='extend' → 기존 헤더의 rowspan 확장
    - col=1, action='new', text='새헤더', rowspan=2 → 새 헤더 셀 생성
    - col=2, action='data' → 데이터 열 (새 셀 생성)
    """
    col: int = 0
    col_span: int = 1
    action: str = 'extend'  # 'extend' | 'new' | 'data'
    text: str = ""  # action='new'일 때 헤더 텍스트
    rowspan: int = 1  # action='new'일 때 새 헤더의 rowspan


@dataclass
class RowAddPlan:
    """
    행 추가 계획

    복잡한 테이블 구조에서 어떻게 행을 추가할지 정의
    """
    headers: List[HeaderConfig] = field(default_factory=list)
    data_cols: List[int] = field(default_factory=list)  # 데이터가 들어갈 열
    rows_to_add: int = 1  # 추가할 행 수


@dataclass
class TableInfo:
    """테이블 정보"""
    table_id: str = ""
    row_count: int = 0
    col_count: int = 0

    # 셀 정보 (row, col) -> CellInfo
    cells: Dict[Tuple[int, int], CellInfo] = field(default_factory=dict)

    # 필드명 -> 셀 위치 매핑
    field_to_cell: Dict[str, Tuple[int, int]] = field(default_factory=dict)

    # XML 요소
    element: Any = None

    # 열별 너비 (col -> width)
    col_widths: Dict[int, int] = field(default_factory=dict)

    # 행별 높이 (row -> height)
    row_heights: Dict[int, int] = field(default_factory=dict)

    def get_col_width(self, col: int) -> int:
        """특정 열의 너비 반환 (colspan 고려)"""
        if col in self.col_widths:
            return self.col_widths[col]
        # 해당 열의 셀에서 너비 찾기
        for (r, c), cell in self.cells.items():
            if c == col and cell.col_span == 1 and cell.width > 0:
                return cell.width
        # 기본값
        return 5000

    def get_row_height(self, row: int) -> int:
        """특정 행의 높이 반환"""
        if row in self.row_heights:
            return self.row_heights[row]
        # 해당 행의 셀에서 높이 찾기
        for (r, c), cell in self.cells.items():
            if r == row and cell.row_span == 1 and cell.height > 0:
                return cell.height
        # 기본값
        return 1000

    def get_cell(self, row: int, col: int) -> Optional[CellInfo]:
        """특정 위치의 셀 반환 (병합 셀 고려)"""
        # 정확한 위치의 셀
        if (row, col) in self.cells:
            return self.cells[(row, col)]

        # rowspan/colspan으로 커버되는 셀 찾기
        for cell in self.cells.values():
            if cell.covers(row, col):
                return cell

        return None

    def get_empty_cells_in_col(self, col: int) -> List[CellInfo]:
        """특정 열의 빈 셀 목록"""
        empty = []
        for row in range(self.row_count):
            cell = self.get_cell(row, col)
            if cell and cell.is_empty and cell.col == col:
                empty.append(cell)
        return empty

    def get_cells_by_field(self, field_name: str) -> List[CellInfo]:
        """필드명으로 셀 찾기"""
        cells = []
        for cell in self.cells.values():
            if cell.field_name == field_name:
                cells.append(cell)
        return cells
