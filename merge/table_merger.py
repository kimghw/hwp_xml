# -*- coding: utf-8 -*-
"""
HWPX 테이블 셀 내용 병합 모듈

테이블을 추가하는 것이 아니라, 기존 테이블의 셀에 내용을 병합합니다.

주요 기능:
1. 필드명(nc.name) 기준으로 셀 매칭
2. 빈 셀에 내용 채우기
3. 행이 부족하면 행 추가 (rowspan 고려)

사용 예:
    merger = TableMerger()
    merger.load_base_table(base_hwpx_path, table_index=0)
    merger.merge_data(data_list)  # [{field_name: value}, ...]
    merger.save(output_path)
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Union, Tuple, Any
from pathlib import Path
from io import BytesIO
import copy
import re


# XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


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

    # 열별 너비
    col_widths: List[int] = field(default_factory=list)

    # 행별 높이
    row_heights: List[int] = field(default_factory=list)

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


class TableParser:
    """HWPX 테이블 파싱"""

    def __init__(self):
        self._ns = ""  # 네임스페이스 접두사

    def parse_tables(self, hwpx_path: Union[str, Path]) -> List[TableInfo]:
        """HWPX 파일에서 모든 테이블 파싱"""
        hwpx_path = Path(hwpx_path)
        tables = []

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                content = zf.read(section_file)
                section_tables = self._parse_section(content)
                tables.extend(section_tables)

        return tables

    def _parse_section(self, xml_content: bytes) -> List[TableInfo]:
        """section XML에서 테이블 파싱"""
        tables = []
        root = ET.parse(BytesIO(xml_content)).getroot()

        # 네임스페이스 추출
        if '}' in root.tag:
            self._ns = root.tag.split('}')[0] + '}'

        # 테이블 찾기
        self._find_tables_recursive(root, tables)

        return tables

    def _find_tables_recursive(self, element, tables: List[TableInfo]):
        """재귀적으로 테이블 찾기"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table = self._parse_table(child)
                tables.append(table)

                # 중첩 테이블도 찾기
                for cell in table.cells.values():
                    if cell.element is not None:
                        self._find_tables_recursive(cell.element, tables)
            else:
                self._find_tables_recursive(child, tables)

    def _parse_table(self, tbl_elem) -> TableInfo:
        """테이블 요소 파싱"""
        table = TableInfo(
            table_id=tbl_elem.get('id', ''),
            element=tbl_elem
        )

        # 열 너비 파싱
        for child in tbl_elem:
            if child.tag.endswith('}tr'):
                # 첫 번째 행에서 열 개수 확인
                col_count = 0
                for tc in child:
                    if tc.tag.endswith('}tc'):
                        col_count += 1
                table.col_count = max(table.col_count, col_count)

        # 행 파싱
        row_idx = 0
        for child in tbl_elem:
            if child.tag.endswith('}tr'):
                self._parse_row(child, row_idx, table)
                row_idx += 1

        table.row_count = row_idx

        # 필드명 매핑 생성
        for (row, col), cell in table.cells.items():
            if cell.field_name:
                table.field_to_cell[cell.field_name] = (row, col)

        return table

    def _parse_row(self, tr_elem, row_idx: int, table: TableInfo):
        """행 파싱"""
        col_idx = 0

        for tc_elem in tr_elem:
            if not tc_elem.tag.endswith('}tc'):
                continue

            cell = CellInfo(
                row=row_idx,
                col=col_idx,
                element=tc_elem
            )

            # 셀 속성 파싱
            for child in tc_elem:
                tag = child.tag.split('}')[-1]

                if tag == 'cellAddr':
                    cell.col = int(child.get('colAddr', col_idx))
                    cell.row = int(child.get('rowAddr', row_idx))
                    col_idx = cell.col

                elif tag == 'cellSpan':
                    cell.col_span = int(child.get('colSpan', 1))
                    cell.row_span = int(child.get('rowSpan', 1))

                elif tag == 'subList':
                    # 텍스트 추출
                    text = self._extract_text(child)
                    cell.text = text
                    cell.is_empty = not text.strip()

                    # 필드명 추출
                    field_name = self._extract_field_name(child)
                    if field_name:
                        cell.field_name = field_name

            cell.end_row = cell.row + cell.row_span - 1
            cell.end_col = cell.col + cell.col_span - 1

            table.cells[(cell.row, cell.col)] = cell
            col_idx += 1

    def _extract_text(self, sublist_elem) -> str:
        """subList에서 텍스트 추출"""
        texts = []
        for p in sublist_elem:
            if p.tag.endswith('}p'):
                for run in p:
                    if run.tag.endswith('}run'):
                        for t in run:
                            if t.tag.endswith('}t') and t.text:
                                texts.append(t.text)
        return ''.join(texts)

    def _extract_field_name(self, sublist_elem) -> str:
        """subList에서 필드명(nc.name) 추출"""
        for p in sublist_elem:
            if p.tag.endswith('}p'):
                for run in p:
                    if run.tag.endswith('}run'):
                        for ctrl in run:
                            # fieldBegin에서 name 속성 찾기
                            if ctrl.tag.endswith('}fieldBegin'):
                                # command 속성에서 필드명 추출
                                # MERGEFIELD 필드명
                                # 또는 nc.name 속성
                                name = ctrl.get('name', '')
                                if name:
                                    return name
        return ""


class TableMerger:
    """테이블 셀 내용 병합"""

    def __init__(self):
        self.parser = TableParser()
        self.base_table: Optional[TableInfo] = None
        self.hwpx_path: Optional[Path] = None
        self.hwpx_data: Dict[str, bytes] = {}  # HWPX 파일 내용

    def load_base_table(self, hwpx_path: Union[str, Path], table_index: int = 0):
        """
        기준 테이블 로드

        Args:
            hwpx_path: HWPX 파일 경로
            table_index: 테이블 인덱스 (0부터 시작)
        """
        self.hwpx_path = Path(hwpx_path)

        # HWPX 파일 내용 로드
        with zipfile.ZipFile(self.hwpx_path, 'r') as zf:
            for name in zf.namelist():
                self.hwpx_data[name] = zf.read(name)

        # 테이블 파싱
        tables = self.parser.parse_tables(self.hwpx_path)

        if table_index >= len(tables):
            raise ValueError(f"테이블 인덱스 {table_index}가 범위를 벗어났습니다. (총 {len(tables)}개)")

        self.base_table = tables[table_index]

        return self.base_table

    def get_table_structure(self) -> Dict[str, Any]:
        """
        테이블 구조 정보 반환

        Returns:
            {
                'row_count': int,
                'col_count': int,
                'fields': [{'name': str, 'row': int, 'col': int}, ...],
                'empty_cells': [{'row': int, 'col': int}, ...],
            }
        """
        if not self.base_table:
            raise ValueError("기준 테이블이 로드되지 않았습니다.")

        fields = []
        empty_cells = []

        for (row, col), cell in self.base_table.cells.items():
            if cell.field_name:
                fields.append({
                    'name': cell.field_name,
                    'row': row,
                    'col': col,
                    'row_span': cell.row_span,
                    'col_span': cell.col_span,
                })

            if cell.is_empty:
                empty_cells.append({
                    'row': row,
                    'col': col,
                })

        return {
            'row_count': self.base_table.row_count,
            'col_count': self.base_table.col_count,
            'fields': fields,
            'empty_cells': empty_cells,
        }

    def merge_data(
        self,
        data_list: List[Dict[str, str]],
        mode: str = "fill_empty"
    ) -> TableInfo:
        """
        데이터 병합

        Args:
            data_list: 병합할 데이터 리스트 [{field_name: value}, ...]
            mode: 병합 모드
                - "fill_empty": 빈 셀만 채우기
                - "append_row": 항상 새 행 추가
                - "smart": 빈 셀 먼저 채우고, 부족하면 행 추가

        Returns:
            병합된 테이블 정보
        """
        if not self.base_table:
            raise ValueError("기준 테이블이 로드되지 않았습니다.")

        if mode == "fill_empty":
            self._fill_empty_cells(data_list)
        elif mode == "append_row":
            self._append_rows(data_list)
        elif mode == "smart":
            self._smart_merge(data_list)
        else:
            raise ValueError(f"알 수 없는 모드: {mode}")

        return self.base_table

    def _fill_empty_cells(self, data_list: List[Dict[str, str]]):
        """빈 셀에 데이터 채우기"""
        for data in data_list:
            for field_name, value in data.items():
                # 필드명으로 셀 찾기
                cells = self.base_table.get_cells_by_field(field_name)

                for cell in cells:
                    if cell.is_empty:
                        self._set_cell_text(cell, value)
                        cell.is_empty = False
                        break  # 첫 번째 빈 셀만 채우기

    def _append_rows(self, data_list: List[Dict[str, str]]):
        """새 행 추가하여 데이터 넣기"""
        for data in data_list:
            self._add_row_with_data(data)

    def add_rows_smart(
        self,
        data_list: List[Dict[str, str]],
        fill_empty_first: bool = True
    ):
        """
        필드명(nc.name) 접두사를 분석하여 자동으로 행 추가

        - 필드명이 'header_'로 시작: 같은 값끼리 병합 (rowspan)
        - 필드명이 'data_'로 시작: 개별 데이터로 처리
        - 그 외: 항상 확장 (extend)

        Args:
            data_list: 데이터 리스트 [{field_name: value, ...}, ...]
            fill_empty_first: True면 빈 셀 먼저 채우고, 없으면 새 행 추가

        예시:
            # 테이블에 필드명이 설정되어 있다면:
            # - Col 0: field_name 없음 (항상 확장)
            # - Col 1: field_name="header_group" (헤더 병합)
            # - Col 2: field_name="data_col1" (데이터)
            # - Col 3: field_name="data_col2" (데이터)

            data = [
                {"header_group": "GroupA", "data_col1": "A", "data_col2": "B"},
                {"header_group": "GroupA", "data_col1": "C", "data_col2": "D"},  # GroupA 확장
                {"header_group": "GroupB", "data_col1": "E", "data_col2": "F"},  # GroupB 새로 생성
            ]
            merger.add_rows_smart(data)
        """
        if self.base_table is None:
            return

        # 필드명 분석하여 열 분류
        header_cols = []  # 'header_'로 시작하는 필드명의 열
        data_cols = []    # 'data_'로 시작하는 필드명의 열
        extend_cols = []  # 그 외 열 (항상 확장)

        for (row, col), cell in self.base_table.cells.items():
            if cell.field_name:
                if cell.field_name.startswith('header_'):
                    if col not in header_cols:
                        header_cols.append(col)
                elif cell.field_name.startswith('data_'):
                    if col not in data_cols:
                        data_cols.append(col)

        # 필드명이 없거나 header_/data_ 접두사 없는 열 찾기
        for col in range(self.base_table.col_count):
            if col not in header_cols and col not in data_cols:
                extend_cols.append(col)

        # 헤더 열이 없으면 첫 번째 열을 extend로
        if not header_cols:
            if data_cols:
                extend_cols = [min(data_cols)]
                data_cols = [c for c in data_cols if c != min(data_cols)]

        # 헤더 열이 여러 개면 마지막 것만 사용 (나머지는 extend)
        if len(header_cols) > 1:
            header_cols.sort()
            main_header_col = header_cols[-1]  # 가장 오른쪽 헤더 열
            extend_cols.extend([c for c in header_cols if c != main_header_col])
            header_cols = [main_header_col]

        header_col = header_cols[0] if header_cols else 0

        # 헤더 필드명 찾기
        header_field_name = None
        for (row, col), cell in self.base_table.cells.items():
            if col == header_col and cell.field_name:
                header_field_name = cell.field_name
                break

        # add_rows_auto 호출
        if header_field_name:
            self.add_rows_auto(
                data_list,
                header_col=header_col,
                data_cols=data_cols,
                extend_cols=extend_cols,
                header_key=header_field_name,
                fill_empty_first=fill_empty_first
            )
        else:
            # 헤더 필드가 없으면 단순히 행 추가
            for data in data_list:
                self._add_row_with_data(data)

    def add_rows_auto(
        self,
        data_list: List[Dict[str, str]],
        header_col: int,
        data_cols: List[int],
        extend_cols: Optional[List[int]] = None,
        header_key: str = "_header",
        fill_empty_first: bool = True
    ):
        """
        헤더 이름을 기준으로 자동 행 추가

        Args:
            data_list: 데이터 리스트 [{header_key: "헤더명", field1: value1, ...}, ...]
            header_col: 헤더가 바뀌는 열 (예: 1)
            data_cols: 데이터 열 목록 (예: [2, 3])
            extend_cols: 항상 확장할 열 (예: [0]) - 생략시 header_col, data_cols 외 모든 열
            header_key: 헤더 이름을 담은 키 (기본값: "_header")
            fill_empty_first: True면 빈 셀 먼저 채우고, 없으면 새 행 추가

        예시:
            # 테이블 구조:
            # +-------+-------+-------+-------+
            # | Head1 | Head2 | Col1  | Col2  |  Row 0
            # | (4행) | (2행) +-------+-------+
            # |       |       | A     | B     |  Row 1
            # |       +-------+-------+-------+
            # |       | Head3 | C     | D     |  Row 2
            # |       | (2행) +-------+-------+
            # |       |       |       |       |  Row 3 (빈 셀)
            # +-------+-------+-------+-------+

            data = [
                {"_header": "Head3", "Col1": "G", "Col2": "H"},  # 빈 셀에 채움
                {"_header": "Head3", "Col1": "I", "Col2": "J"},  # 새 행 추가, Head3 확장
                {"_header": "Head4", "Col1": "K", "Col2": "L"},  # Head4 새로 생성
            ]
            merger.add_rows_auto(data, header_col=1, data_cols=[2, 3], extend_cols=[0])
        """
        if self.base_table is None:
            return

        # extend_cols 기본값 설정
        if extend_cols is None:
            extend_cols = [c for c in range(self.base_table.col_count)
                          if c != header_col and c not in data_cols]

        # 현재 헤더 상태 추적
        current_header_text = None
        current_header_remaining = 0  # 새 헤더의 남은 rowspan

        for data in data_list:
            header_name = data.get(header_key, "")
            data_without_header = {k: v for k, v in data.items() if k != header_key}

            # 빈 셀 먼저 채우기 시도
            if fill_empty_first:
                filled = self._try_fill_empty_cells(
                    data_without_header, header_col, header_name, data_cols
                )
                if filled:
                    continue  # 빈 셀에 채웠으면 다음 데이터로

            # 빈 셀이 없거나 fill_empty_first=False면 새 행 추가
            last_row = self.base_table.row_count - 1

            # 마지막 행의 header_col 셀 확인
            header_cell = self.base_table.get_cell(last_row, header_col)
            existing_header_text = header_cell.text if header_cell else ""

            # 헤더 설정 생성
            header_config = []

            for col in range(self.base_table.col_count):
                cell = self.base_table.get_cell(last_row, col)

                if col in extend_cols:
                    # 항상 확장하는 열 (Head1 같은)
                    header_config.append(HeaderConfig(
                        col=col,
                        col_span=cell.col_span if cell else 1,
                        action='extend'
                    ))

                elif col == header_col:
                    # 헤더가 바뀌는 열
                    if current_header_remaining > 0:
                        # 이전에 새로 생성한 헤더의 rowspan 범위 내 → 이미 커버됨
                        current_header_remaining -= 1
                        continue
                    elif header_name == existing_header_text or header_name == current_header_text:
                        # 같은 헤더 → 확장
                        header_config.append(HeaderConfig(
                            col=col,
                            col_span=cell.col_span if cell else 1,
                            action='extend'
                        ))
                        current_header_text = header_name
                    else:
                        # 다른 헤더 → 새 헤더 생성
                        header_config.append(HeaderConfig(
                            col=col,
                            col_span=cell.col_span if cell else 1,
                            action='new',
                            text=header_name,
                            rowspan=2
                        ))
                        current_header_text = header_name
                        current_header_remaining = 1  # rowspan=2이므로 1행 더 커버

                elif col in data_cols:
                    # 데이터 열
                    header_config.append(HeaderConfig(
                        col=col,
                        action='data'
                    ))

            # 행 추가
            self.add_row_with_headers(data_without_header, header_config)

    def _try_fill_empty_cells(
        self,
        data: Dict[str, str],
        header_col: int,
        header_name: str,
        data_cols: List[int]
    ) -> bool:
        """
        같은 헤더 아래 빈 셀에 데이터 채우기 시도

        Returns:
            True: 빈 셀에 채움
            False: 빈 셀 없음 (새 행 추가 필요)
        """
        if self.base_table is None:
            return False

        # 같은 헤더 아래 빈 셀 찾기
        for row in range(self.base_table.row_count):
            # 이 행의 헤더 확인
            header_cell = self.base_table.get_cell(row, header_col)
            if not header_cell or header_cell.text != header_name:
                continue

            # 이 헤더가 커버하는 행 범위
            header_start = header_cell.row
            header_end = header_cell.end_row

            # 이 헤더 범위 내에서 빈 셀 행 찾기
            for check_row in range(header_start, header_end + 1):
                # 모든 data_cols가 비어있는지 확인
                all_empty = True
                cells_to_fill = []

                for col in data_cols:
                    cell = self.base_table.get_cell(check_row, col)
                    if cell and cell.row == check_row:  # 실제 셀 (rowspan 커버 아님)
                        if cell.is_empty:
                            cells_to_fill.append((cell, col))
                        else:
                            all_empty = False
                            break
                    else:
                        all_empty = False
                        break

                if all_empty and cells_to_fill:
                    # 빈 셀에 데이터 채우기
                    for cell, col in cells_to_fill:
                        # 필드명으로 값 찾기
                        for field_name, value in data.items():
                            if field_name in self.base_table.field_to_cell:
                                field_row, field_col = self.base_table.field_to_cell[field_name]
                                if field_col == col:
                                    self._set_cell_text(cell, value)
                                    cell.is_empty = False
                                    cell.text = value
                                    break
                    return True

        return False

    def add_row_with_headers(
        self,
        data: Dict[str, str],
        header_config: List[HeaderConfig]
    ):
        """
        헤더 설정에 따라 새 행 추가

        Args:
            data: {field_name: value} 데이터
            header_config: 각 열의 헤더 설정

        예시:
            # Head1은 확장, Head2는 새로 생성, 나머지는 데이터
            config = [
                HeaderConfig(col=0, action='extend'),
                HeaderConfig(col=1, action='new', text='새헤더2', rowspan=2),
                HeaderConfig(col=2, action='data'),
                HeaderConfig(col=3, action='data'),
            ]
            merger.add_row_with_headers({"Col1": "A", "Col2": "B"}, config)
        """
        if self.base_table is None or self.base_table.element is None:
            return

        last_row_idx = self.base_table.row_count - 1
        new_row_idx = self.base_table.row_count

        # 1. 헤더 설정 분석
        col_actions = {}  # col -> HeaderConfig
        for hc in header_config:
            for c in range(hc.col, hc.col + hc.col_span):
                col_actions[c] = hc

        # 2. rowspan 확장 처리
        for hc in header_config:
            if hc.action == 'extend':
                cell = self.base_table.get_cell(last_row_idx, hc.col)
                if cell:
                    # rowspan 셀이든 일반 셀이든 모두 확장
                    self._extend_rowspan(cell)

        # 3. 새 행 XML 생성
        self._create_new_row_with_headers(new_row_idx, data, header_config)

        self.base_table.row_count += 1

    def _create_new_row_with_headers(
        self,
        row_idx: int,
        data: Dict[str, str],
        header_config: List[HeaderConfig]
    ):
        """헤더 설정에 따라 새 행 XML 생성"""
        if self.base_table is None or self.base_table.element is None:
            return

        # 마지막 tr 찾기 (템플릿)
        last_tr = None
        for child in self.base_table.element:
            if child.tag.endswith('}tr'):
                last_tr = child

        if last_tr is None:
            return

        # 새 tr 생성
        new_tr = copy.deepcopy(last_tr)

        # 기존 셀 모두 제거
        for tc in list(new_tr):
            if tc.tag.endswith('}tc'):
                new_tr.remove(tc)

        # 필드명 -> 열 매핑
        field_to_col = {}
        for field_name, (_, col) in self.base_table.field_to_cell.items():
            field_to_col[field_name] = col

        # 값 -> 열 매핑
        cols_with_data = {}
        for field_name, value in data.items():
            if field_name in field_to_col:
                cols_with_data[field_to_col[field_name]] = value

        # 설정에 따라 셀 생성
        processed_cols = set()
        for hc in sorted(header_config, key=lambda x: x.col):
            if hc.col in processed_cols:
                continue

            if hc.action == 'extend':
                # rowspan 확장된 셀 - 새 행에 셀 없음
                for c in range(hc.col, hc.col + hc.col_span):
                    processed_cols.add(c)

            elif hc.action == 'new':
                # 새 헤더 셀 생성
                tc = self._create_cell_element(
                    row_idx, hc.col, hc.text,
                    rowspan=hc.rowspan, colspan=hc.col_span
                )
                if tc is not None:
                    new_tr.append(tc)

                    # CellInfo 추가
                    cell = CellInfo(
                        row=row_idx,
                        col=hc.col,
                        row_span=hc.rowspan,
                        col_span=hc.col_span,
                        end_row=row_idx + hc.rowspan - 1,
                        end_col=hc.col + hc.col_span - 1,
                        text=hc.text,
                        is_empty=not hc.text,
                        element=tc
                    )
                    self.base_table.cells[(row_idx, hc.col)] = cell

                for c in range(hc.col, hc.col + hc.col_span):
                    processed_cols.add(c)

            elif hc.action == 'data':
                # 데이터 셀 생성
                value = cols_with_data.get(hc.col, "")
                tc = self._create_cell_element(
                    row_idx, hc.col, value,
                    colspan=hc.col_span
                )
                if tc is not None:
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=hc.col,
                        col_span=hc.col_span,
                        end_col=hc.col + hc.col_span - 1,
                        text=value,
                        is_empty=not value,
                        element=tc
                    )
                    self.base_table.cells[(row_idx, hc.col)] = cell

                for c in range(hc.col, hc.col + hc.col_span):
                    processed_cols.add(c)

        # 테이블에 새 행 추가
        self.base_table.element.append(new_tr)

    def _create_cell_element(
        self,
        row: int,
        col: int,
        text: str,
        rowspan: int = 1,
        colspan: int = 1
    ) -> Optional[Any]:
        """
        새 셀 XML 요소 생성

        기존 셀을 템플릿으로 사용
        """
        if self.base_table is None:
            return None

        # 템플릿 셀 찾기 (같은 열의 기존 셀)
        template_cell = None
        for (r, c), cell in self.base_table.cells.items():
            if c == col and cell.element is not None:
                template_cell = cell
                break

        if template_cell is None:
            # 아무 셀이나 템플릿으로 사용
            for cell in self.base_table.cells.values():
                if cell.element is not None:
                    template_cell = cell
                    break

        if template_cell is None:
            return None

        # 셀 복사
        tc = copy.deepcopy(template_cell.element)

        # 속성 업데이트
        for child in tc:
            tag = child.tag.split('}')[-1]

            if tag == 'cellAddr':
                child.set('colAddr', str(col))
                child.set('rowAddr', str(row))

            elif tag == 'cellSpan':
                child.set('colSpan', str(colspan))
                child.set('rowSpan', str(rowspan))

            elif tag == 'subList':
                # 텍스트 설정
                for p in child:
                    if p.tag.endswith('}p'):
                        for run in p:
                            if run.tag.endswith('}run'):
                                for t in run:
                                    if t.tag.endswith('}t'):
                                        t.text = text
                                        break
                                break
                        break

        return tc

    def _smart_merge(self, data_list: List[Dict[str, str]]):
        """스마트 병합: 빈 셀 먼저, 부족하면 행 추가"""
        remaining_data = list(data_list)

        # 1단계: 빈 셀 채우기
        filled = []
        for data in remaining_data:
            all_filled = True
            for field_name, value in data.items():
                cells = self.base_table.get_cells_by_field(field_name)
                empty_found = False

                for cell in cells:
                    if cell.is_empty:
                        self._set_cell_text(cell, value)
                        cell.is_empty = False
                        empty_found = True
                        break

                if not empty_found:
                    all_filled = False

            if all_filled:
                filled.append(data)

        # 2단계: 남은 데이터는 행 추가
        for data in remaining_data:
            if data not in filled:
                self._add_row_with_data(data)

    def _set_cell_text(self, cell: CellInfo, text: str):
        """셀에 텍스트 설정"""
        if cell.element is None:
            return

        # subList > p > run > t 구조에서 텍스트 설정
        for child in cell.element:
            if child.tag.endswith('}subList'):
                for p in child:
                    if p.tag.endswith('}p'):
                        for run in p:
                            if run.tag.endswith('}run'):
                                for t in run:
                                    if t.tag.endswith('}t'):
                                        t.text = text
                                        cell.text = text
                                        return

                        # t 요소가 없으면 생성
                        # (실제 구현에서는 run과 t 요소를 생성해야 함)
                        break

    def _add_row_with_data(self, data: Dict[str, str]):
        """
        새 행 추가 (중첩 rowspan 고려)

        복잡한 rowspan 구조 처리:
        +---------------+---------------+-------+-------+
        |               |               | Col1  | Col2  |
        |    Head1      |    Head2      +-------+-------+
        |   (4행)       |   (2행)       | A     | B     |
        |               +---------------+-------+-------+
        |               |    Head3      | C     | D     |
        |               |   (2행)       +-------+-------+
        |               |               | E     | F     |
        +---------------+---------------+-------+-------+

        각 열별로 해당 열을 커버하는 rowspan 셀을 찾아 확장 여부 결정
        """
        if self.base_table is None or self.base_table.element is None:
            return

        # 마지막 행 찾기
        last_row_idx = self.base_table.row_count - 1

        # 각 열의 rowspan 상태 확인
        # col -> (status, covering_cell)
        # status: 'extend_rowspan' | 'new_cell' | 'skip' (colspan으로 커버됨)
        col_status = {}

        # 모든 열 순회하며 상태 결정
        col = 0
        while col < self.base_table.col_count:
            cell = self.base_table.get_cell(last_row_idx, col)

            if cell:
                # 이 셀의 colspan 범위 처리
                for c in range(cell.col, cell.col + cell.col_span):
                    if c == cell.col:
                        # 셀의 시작 열
                        if cell.row < last_row_idx:
                            # rowspan으로 위에서 내려온 셀
                            col_status[c] = ('extend_rowspan', cell)
                        else:
                            # 마지막 행에서 시작하는 셀
                            col_status[c] = ('new_cell', cell)
                    else:
                        # colspan으로 커버되는 열
                        col_status[c] = ('skip', cell)

                col = cell.col + cell.col_span
            else:
                col_status[col] = ('new_cell', None)
                col += 1

        # 새 행 추가
        new_row_idx = self.base_table.row_count

        # 데이터가 들어갈 열과 rowspan 확장할 열 결정
        cols_with_data = {}  # col -> value
        cols_to_extend = set()

        for field_name, value in data.items():
            if field_name in self.base_table.field_to_cell:
                _, col = self.base_table.field_to_cell[field_name]
                cols_with_data[col] = value

        # 열별로 처리
        for col in range(self.base_table.col_count):
            status, ref_cell = col_status.get(col, ('new_cell', None))

            if status == 'skip':
                continue

            if col in cols_with_data:
                # 이 열에 데이터가 있음
                if status == 'extend_rowspan' and ref_cell:
                    # rowspan 셀이지만 새 데이터 있음 → 확장하지 않고 새 셀 생성
                    pass
                # 새 셀 생성 (아래에서 처리)
            else:
                # 이 열에 데이터가 없음
                if status == 'extend_rowspan' and ref_cell:
                    cols_to_extend.add(col)

        # rowspan 확장 (셀 위치로 추적)
        extended_cell_positions = set()
        for col in cols_to_extend:
            _, ref_cell = col_status.get(col, (None, None))
            if ref_cell:
                cell_pos = (ref_cell.row, ref_cell.col)
                if cell_pos not in extended_cell_positions:
                    self._extend_rowspan(ref_cell)
                    extended_cell_positions.add(cell_pos)

        # 새 행 XML 생성
        self._create_new_row(new_row_idx, cols_with_data, col_status)

        self.base_table.row_count += 1

    def _create_new_row(
        self,
        row_idx: int,
        cols_with_data: Dict[int, str],
        col_status: Dict[int, Tuple[str, Optional[CellInfo]]]
    ):
        """새 행 XML 생성"""
        if self.base_table is None or self.base_table.element is None:
            return

        # 마지막 tr 요소 찾기 (복사 템플릿용)
        last_tr = None
        for child in self.base_table.element:
            if child.tag.endswith('}tr'):
                last_tr = child

        if last_tr is None:
            return

        # 새 tr 생성
        new_tr = copy.deepcopy(last_tr)

        # 기존 셀 모두 제거
        for tc in list(new_tr):
            if tc.tag.endswith('}tc'):
                new_tr.remove(tc)

        # 데이터가 들어갈 열 찾기 (rowspan으로 커버되지 않는 열)
        data_cols = []
        for col in range(self.base_table.col_count):
            status, _ = col_status.get(col, ('new_cell', None))
            if status == 'new_cell':
                data_cols.append(col)

        # 각 데이터 열에 새 셀 생성
        for col in data_cols:
            tc = self._create_cell_element(
                row_idx, col,
                cols_with_data.get(col, "")
            )
            if tc is not None:
                new_tr.append(tc)

                # CellInfo 추가
                cell = CellInfo(
                    row=row_idx,
                    col=col,
                    text=cols_with_data.get(col, ""),
                    is_empty=col not in cols_with_data,
                    element=tc
                )
                self.base_table.cells[(row_idx, col)] = cell

        # 테이블에 새 행 추가
        self.base_table.element.append(new_tr)

    def _set_tc_text(self, tc_elem, text: str):
        """tc 요소의 텍스트 설정"""
        for child in tc_elem:
            if child.tag.endswith('}subList'):
                for p in child:
                    if p.tag.endswith('}p'):
                        for run in p:
                            if run.tag.endswith('}run'):
                                for t in run:
                                    if t.tag.endswith('}t'):
                                        t.text = text
                                        return

    def _extend_rowspan(self, cell: CellInfo):
        """셀의 rowspan 확장"""
        if cell.element is None:
            return

        for child in cell.element:
            if child.tag.endswith('}cellSpan'):
                current_rowspan = int(child.get('rowSpan', 1))
                child.set('rowSpan', str(current_rowspan + 1))
                cell.row_span += 1
                cell.end_row += 1
                return

    def _create_new_cell(self, row: int, col: int, text: str, template_cell: Optional[CellInfo]):
        """새 셀 생성"""
        # 실제 구현에서는 template_cell을 복사하여 새 셀 생성
        # XML 구조: tr > tc > cellAddr, cellSpan, cellSz, subList > p > run > t
        pass  # 복잡한 XML 조작 필요

    def save(self, output_path: Union[str, Path]):
        """병합된 테이블을 HWPX 파일로 저장"""
        if not self.base_table:
            raise ValueError("기준 테이블이 로드되지 않았습니다.")

        output_path = Path(output_path)

        # section XML 업데이트
        # 기존 테이블 요소를 수정된 것으로 교체

        # 실제 구현에서는 self.base_table.element의 변경사항을
        # section XML에 반영하고 새 HWPX 파일 생성

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, content in self.hwpx_data.items():
                if name.startswith('Contents/section') and name.endswith('.xml'):
                    # 테이블이 수정된 section XML 재생성
                    modified_content = self._rebuild_section_xml(name, content)
                    zf.writestr(name, modified_content)
                elif name == 'mimetype':
                    zf.writestr(name, content, compress_type=zipfile.ZIP_STORED)
                else:
                    zf.writestr(name, content)

        return output_path

    def _rebuild_section_xml(self, section_name: str, original_content: bytes) -> bytes:
        """section XML 재구성"""
        # 현재는 원본 그대로 반환
        # 실제 구현에서는 self.base_table.element의 변경사항 반영
        root = ET.parse(BytesIO(original_content)).getroot()

        # 수정된 테이블 요소로 교체
        # TODO: 테이블 요소 찾아서 교체

        return ET.tostring(root, encoding='UTF-8', xml_declaration=True)


def create_header_config_from_table(
    table: TableInfo,
    data_cols: List[int],
    new_headers: Optional[Dict[int, Tuple[str, int]]] = None
) -> List[HeaderConfig]:
    """
    테이블 구조에서 자동으로 HeaderConfig 생성

    Args:
        table: 테이블 정보
        data_cols: 데이터가 들어갈 열 목록
        new_headers: 새로 생성할 헤더 {col: (text, rowspan)}
                     예: {1: ("새헤더", 2)}

    Returns:
        HeaderConfig 리스트

    예시:
        # 기존 테이블:
        # +-------+-------+-------+-------+
        # | Head1 | Head2 | Col1  | Col2  |
        # | (2행) | (2행) |       |       |
        # +-------+-------+-------+-------+

        # Col1, Col2에 데이터 추가, Head2 자리에 새 헤더 생성
        config = create_header_config_from_table(
            table,
            data_cols=[2, 3],
            new_headers={1: ("Head2-2", 2)}
        )
        # 결과:
        # [
        #   HeaderConfig(col=0, action='extend'),      # Head1 확장
        #   HeaderConfig(col=1, action='new', text='Head2-2', rowspan=2),
        #   HeaderConfig(col=2, action='data'),
        #   HeaderConfig(col=3, action='data'),
        # ]
    """
    new_headers = new_headers or {}
    last_row = table.row_count - 1
    configs = []
    processed = set()

    col = 0
    while col < table.col_count:
        if col in processed:
            col += 1
            continue

        cell = table.get_cell(last_row, col)

        if col in new_headers:
            # 새 헤더 생성
            text, rowspan = new_headers[col]
            colspan = cell.col_span if cell else 1
            configs.append(HeaderConfig(
                col=col,
                col_span=colspan,
                action='new',
                text=text,
                rowspan=rowspan
            ))
            for c in range(col, col + colspan):
                processed.add(c)
            col += colspan

        elif col in data_cols:
            # 데이터 열
            colspan = 1
            configs.append(HeaderConfig(
                col=col,
                col_span=colspan,
                action='data'
            ))
            processed.add(col)
            col += 1

        elif cell and cell.row < last_row:
            # 기존 rowspan 셀 - 확장
            configs.append(HeaderConfig(
                col=col,
                col_span=cell.col_span,
                action='extend'
            ))
            for c in range(col, col + cell.col_span):
                processed.add(c)
            col = cell.col + cell.col_span

        else:
            # 일반 셀 - 데이터로 처리
            colspan = cell.col_span if cell else 1
            configs.append(HeaderConfig(
                col=col,
                col_span=colspan,
                action='data'
            ))
            for c in range(col, col + colspan):
                processed.add(c)
            col += colspan

    return configs


def print_table_structure(table: TableInfo, max_rows: int = 20):
    """테이블 구조 출력"""
    print(f"테이블 ID: {table.table_id}")
    print(f"크기: {table.row_count}행 x {table.col_count}열")
    print()

    # 셀 그리드 출력
    for row in range(min(table.row_count, max_rows)):
        row_str = f"Row {row:2d}: "
        for col in range(table.col_count):
            cell = table.get_cell(row, col)
            if cell:
                if cell.row == row and cell.col == col:
                    # 셀 시작 위치
                    text = cell.text[:10] + "..." if len(cell.text) > 10 else cell.text
                    field = f"[{cell.field_name}]" if cell.field_name else ""
                    span = f"({cell.row_span}x{cell.col_span})" if cell.row_span > 1 or cell.col_span > 1 else ""
                    row_str += f" | {text or '(empty)'}{field}{span}"
                else:
                    # rowspan/colspan으로 커버되는 영역
                    row_str += " | ↓"
            else:
                row_str += " | -"
        print(row_str)

    if table.row_count > max_rows:
        print(f"... ({table.row_count - max_rows}행 더 있음)")


def analyze_rowspan_structure(table: TableInfo, target_row: int = -1):
    """
    rowspan 구조 분석 - 중첩 rowspan 시각화

    Args:
        table: 테이블 정보
        target_row: 분석할 행 (-1이면 마지막 행)

    Returns:
        각 행에서 어떤 셀이 rowspan으로 커버되는지 분석
    """
    print("\n[Rowspan 구조 분석]")
    print("=" * 60)

    # 대상 행 결정
    if target_row < 0:
        target_row = table.row_count - 1
    last_row = target_row
    print(f"Row {last_row} 기준 각 열 상태:\n")

    col = 0
    while col < table.col_count:
        cell = table.get_cell(last_row, col)
        if cell:
            if cell.row < last_row:
                # rowspan 셀
                span_rows = f"Row {cell.row}-{cell.end_row}"
                span_cols = f"Col {cell.col}-{cell.end_col}" if cell.col_span > 1 else f"Col {cell.col}"
                print(f"  Col {col}: ROWSPAN 셀 ({span_rows}, {span_cols})")
                print(f"           텍스트: '{cell.text[:30]}...' " if len(cell.text) > 30 else f"           텍스트: '{cell.text}'")
                print(f"           → 새 행 추가 시: rowspan 확장 ({cell.row_span} → {cell.row_span + 1})")
            else:
                # 일반 셀
                if cell.col_span > 1:
                    print(f"  Col {col}: 일반 셀 (colspan={cell.col_span})")
                else:
                    print(f"  Col {col}: 일반 셀")
                print(f"           → 새 행 추가 시: 새 셀 생성")

            col = cell.col + cell.col_span
        else:
            print(f"  Col {col}: 빈 슬롯")
            col += 1

    print()


def get_row_template_info(table: TableInfo, target_row: int = -1) -> Dict[str, Any]:
    """
    특정 행의 셀 구조 분석 (새 행 추가용 템플릿)

    Args:
        table: 테이블 정보
        target_row: 분석할 행 (-1이면 마지막 행)

    Returns:
        {
            'row': int,
            'cells': [
                {'col': int, 'status': str, 'cell': CellInfo or None},
                ...
            ]
        }
    """
    if target_row < 0:
        target_row = table.row_count - 1

    result = {
        'row': target_row,
        'cells': []
    }

    col = 0
    while col < table.col_count:
        cell = table.get_cell(target_row, col)

        if cell:
            if cell.row < target_row:
                status = 'extend_rowspan'
            else:
                status = 'new_cell'

            result['cells'].append({
                'col': col,
                'col_span': cell.col_span,
                'status': status,
                'cell': cell,
                'text': cell.text[:20] if cell.text else '',
            })
            col = cell.col + cell.col_span
        else:
            result['cells'].append({
                'col': col,
                'col_span': 1,
                'status': 'empty',
                'cell': None,
                'text': '',
            })
            col += 1

    return result


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법:")
        print("  python table_merger.py <hwpx_file> [table_index]")
        print("  python table_merger.py <hwpx_file> [table_index] --analyze")
        print()
        print("헤더 설정 예시 (Python):")
        print("""
from table_merger import TableMerger, HeaderConfig, create_header_config_from_table

merger = TableMerger()
table = merger.load_base_table("input.hwpx", 0)

# 방법 1: 수동 설정
config = [
    HeaderConfig(col=0, action='extend'),           # Head1 rowspan 확장
    HeaderConfig(col=1, action='new', text='새헤더', rowspan=2),  # 새 헤더
    HeaderConfig(col=2, action='data'),             # 데이터 열
    HeaderConfig(col=3, action='data'),             # 데이터 열
]
merger.add_row_with_headers({"Col1": "A", "Col2": "B"}, config)

# 방법 2: 자동 생성
config = create_header_config_from_table(
    table,
    data_cols=[2, 3],                # 데이터 열
    new_headers={1: ("새헤더", 2)}   # col 1에 새 헤더 (rowspan=2)
)
merger.add_row_with_headers({"Col1": "A", "Col2": "B"}, config)

merger.save("output.hwpx")
""")
        sys.exit(1)

    hwpx_path = sys.argv[1]
    table_index = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else 0
    analyze = '--analyze' in sys.argv

    print(f"파일: {hwpx_path}")
    print(f"테이블 인덱스: {table_index}")
    print("=" * 60)

    merger = TableMerger()

    try:
        table = merger.load_base_table(hwpx_path, table_index)
        print_table_structure(table)

        print("\n테이블 구조 정보:")
        structure = merger.get_table_structure()
        print(f"  행: {structure['row_count']}, 열: {structure['col_count']}")
        print(f"  필드: {len(structure['fields'])}개")
        for f in structure['fields']:
            print(f"    - {f['name']} @ ({f['row']}, {f['col']})")
        print(f"  빈 셀: {len(structure['empty_cells'])}개")

        if analyze:
            # 복잡한 rowspan 구조가 있는 행 찾기
            complex_rows = []
            for row in range(table.row_count):
                rowspan_count = 0
                for col in range(table.col_count):
                    cell = table.get_cell(row, col)
                    if cell and cell.row < row:
                        rowspan_count += 1
                if rowspan_count >= 2:  # 2개 이상의 rowspan이 교차하는 행
                    complex_rows.append((row, rowspan_count))

            if complex_rows:
                print(f"\n복잡한 rowspan 구조 행: {len(complex_rows)}개")
                for row, count in complex_rows[:5]:
                    print(f"  Row {row}: {count}개 열이 rowspan으로 커버됨")
                    analyze_rowspan_structure(table, row)
            else:
                analyze_rowspan_structure(table)

            print("\n[행 추가 템플릿 분석]")
            template = get_row_template_info(table)
            print(f"마지막 행 (Row {template['row']}) 기준:")
            for cell_info in template['cells']:
                status_str = {
                    'extend_rowspan': 'ROWSPAN 확장',
                    'new_cell': '새 셀 생성',
                    'empty': '빈 슬롯',
                }[cell_info['status']]
                print(f"  Col {cell_info['col']}: {status_str} (colspan={cell_info['col_span']})")

    except Exception as e:
        print(f"오류: {e}")
        import traceback
        traceback.print_exc()
