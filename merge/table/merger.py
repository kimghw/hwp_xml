# -*- coding: utf-8 -*-
"""
HWPX 테이블 셀 내용 병합 모듈

Base 파일(원본)의 테이블에 Add 파일(추가 데이터)의 내용을 병합합니다.
병합은 input_ 필드가 있는 영역에만 데이터를 추가하며, 기존 데이터(data_)는 유지됩니다.

필드명 접두사별 동작:
- header_: 유지 (변경 없음)
- data_: 유지 (변경 없음)
- add_: 기존 셀에 내용 추가 (행 추가 없음)
- stub_: 새 행 생성
- gstub_: 같은 값이면 rowspan 확장, 다른 값이면 새 셀 생성
- input_: 빈 셀 채우기, 없으면 새 행 추가

사용 예:
    merger = TableMerger()
    merger.load_base_table("base.hwpx", table_index=0)

    add_data = [
        {"gstub_abc": "그룹A", "input_123": "값1"},
        {"gstub_abc": "그룹A", "input_123": "값2"},  # gstub rowspan 확장
        {"gstub_abc": "그룹B", "input_123": "값3"},  # 새 gstub 셀
    ]
    merger.merge_with_stub(add_data)
    merger.save("output.hwpx")
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Union, Tuple, Any
from pathlib import Path
from io import BytesIO
import copy

from .models import CellInfo, HeaderConfig, TableInfo
from .parser import TableParser, NAMESPACES
from ..format_validator import AddFieldValidator, AddFieldValidationResult, CellStyleInfo

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class TableMerger:
    """테이블 셀 내용 병합"""

    def __init__(self, validate_format: bool = False, sdk_validator=None):
        """
        Args:
            validate_format: True면 add_/input_ 필드 병합 전 형식 검증
            sdk_validator: Claude Code SDK 검증 함수 (외부 주입)
        """
        self.parser = TableParser()
        self.base_table: Optional[TableInfo] = None
        self.hwpx_path: Optional[Path] = None
        self.hwpx_data: Dict[str, bytes] = {}  # HWPX 파일 내용
        self.validate_format = validate_format
        self.field_validator = AddFieldValidator(sdk_validator) if validate_format else None

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

    def merge_with_stub(
        self,
        data_list: List[Dict[str, str]],
        fill_empty_first: bool = True,
        field_styles: Optional[Dict[str, str]] = None,
        add_separator: str = " "
    ):
        """
        stub/gstub/input 기반 병합

        접두사별 처리:
        - header_: 유지 (변경 없음)
        - data_: 유지 (변경 없음)
        - add_: 기존 셀에 내용 추가 (같은 문단 시 빈칸 1개 구분)
        - stub_: 새 행 생성
        - gstub_: 같은 값이면 rowspan 확장, 다른 값이면 새 셀
        - input_: 빈 셀 채우기, 없으면 새 행 추가

        Args:
            data_list: [{field_name: value}, ...] 형태의 데이터
            fill_empty_first: True면 빈 input_ 셀 먼저 채움
            field_styles: {field_name: style_type} add_ 필드 스타일 지정 (검증용)
            add_separator: add_ 필드 구분자 (기본: 빈칸 1개)
        """
        if self.base_table is None:
            return

        # 필드명별 열 분류
        field_cols = self._classify_field_columns()

        # add_ 필드 먼저 처리 (행 추가 없음)
        add_data = {}
        row_data_list = []

        for data in data_list:
            row_item = {}
            for field_name, value in data.items():
                if field_name.startswith('add_'):
                    if field_name not in add_data:
                        add_data[field_name] = []
                    add_data[field_name].append(value)
                else:
                    row_item[field_name] = value
            if row_item:
                row_data_list.append(row_item)

        # add_ 처리 (같은 문단 = 빈칸 1개 구분)
        for field_name, values in add_data.items():
            combined = add_separator.join(values)
            self._process_add_fields(
                {field_name: combined},
                separator=add_separator,
                field_styles=field_styles
            )

        # 각 데이터 행 처리
        for data in row_data_list:
            self._merge_single_row(data, field_cols, fill_empty_first)

    def _classify_field_columns(self) -> Dict[str, List[int]]:
        """필드명 접두사별 열 분류"""
        result = {
            'header': [],
            'data': [],
            'add': [],
            'stub': [],
            'gstub': [],
            'input': [],
        }

        if self.base_table is None:
            return result

        for (row, col), cell in self.base_table.cells.items():
            if not cell.field_name:
                continue

            prefix = cell.field_name.split('_')[0] + '_'
            if prefix == 'header_' and col not in result['header']:
                result['header'].append(col)
            elif prefix == 'data_' and col not in result['data']:
                result['data'].append(col)
            elif prefix == 'add_' and col not in result['add']:
                result['add'].append(col)
            elif prefix == 'stub_' and col not in result['stub']:
                result['stub'].append(col)
            elif prefix == 'gstub_' and col not in result['gstub']:
                result['gstub'].append(col)
            elif prefix == 'input_' and col not in result['input']:
                result['input'].append(col)

        return result

    def _merge_single_row(
        self,
        data: Dict[str, str],
        field_cols: Dict[str, List[int]],
        fill_empty_first: bool
    ):
        """단일 데이터 행 병합"""
        if self.base_table is None:
            return

        # gstub 값 추출
        gstub_values = {}
        for field_name, value in data.items():
            if field_name.startswith('gstub_'):
                gstub_values[field_name] = value

        # stub 값 추출
        stub_values = {}
        for field_name, value in data.items():
            if field_name.startswith('stub_'):
                stub_values[field_name] = value

        # input 값 추출
        input_values = {}
        for field_name, value in data.items():
            if field_name.startswith('input_'):
                input_values[field_name] = value

        if not input_values:
            return  # input 데이터 없으면 스킵

        # 1. 빈 셀 먼저 채우기 시도
        if fill_empty_first:
            filled = self._try_fill_input_cells(input_values, gstub_values, stub_values)
            if filled:
                return

        # 2. 새 행 추가 필요
        self._add_row_with_stub(data, gstub_values, stub_values, input_values)

    def _try_fill_input_cells(
        self,
        input_values: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str]
    ) -> bool:
        """빈 input_ 셀에 데이터 채우기 시도"""
        if self.base_table is None:
            return False

        # 같은 gstub/stub 패턴의 빈 행 찾기
        for row in range(self.base_table.row_count):
            # 이 행의 모든 input 셀이 비어있는지 확인
            row_empty = True
            matching_gstub = True
            matching_stub = True

            # gstub 매칭 확인
            for field_name, expected_value in gstub_values.items():
                cells = self.base_table.get_cells_by_field(field_name)
                for cell in cells:
                    if cell.row <= row <= cell.end_row:
                        if cell.text != expected_value:
                            matching_gstub = False
                        break

            # stub 매칭 확인
            for field_name, expected_value in stub_values.items():
                cells = self.base_table.get_cells_by_field(field_name)
                for cell in cells:
                    if cell.row == row:
                        if cell.text != expected_value:
                            matching_stub = False
                        break

            if not matching_gstub or not matching_stub:
                continue

            # input 셀 빈 여부 확인
            cells_to_fill = []
            for field_name in input_values:
                cells = self.base_table.get_cells_by_field(field_name)
                for cell in cells:
                    if cell.row == row:
                        if cell.is_empty:
                            cells_to_fill.append((cell, field_name))
                        else:
                            row_empty = False
                        break

            if row_empty and cells_to_fill:
                # 빈 셀 채우기
                for cell, field_name in cells_to_fill:
                    value = input_values.get(field_name, "")
                    self._set_cell_text(cell, value)
                    cell.is_empty = False
                    cell.text = value
                return True

        return False

    def _add_row_with_stub(
        self,
        data: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str]
    ):
        """stub/gstub 고려하여 새 행 추가"""
        if self.base_table is None or self.base_table.element is None:
            return

        last_row = self.base_table.row_count - 1
        new_row = self.base_table.row_count

        # 각 열의 처리 방법 결정
        col_actions = {}  # col -> ('extend'|'new'|'data', cell, value)

        for col in range(self.base_table.col_count):
            cell = self.base_table.get_cell(last_row, col)
            if cell is None:
                continue

            field_name = cell.field_name

            if field_name and field_name.startswith('gstub_'):
                # gstub 처리
                new_value = gstub_values.get(field_name, "")
                if cell.text == new_value:
                    # 같은 값 → rowspan 확장
                    col_actions[col] = ('extend', cell, None)
                else:
                    # 다른 값 → 새 셀 생성
                    col_actions[col] = ('new', cell, new_value)

            elif field_name and field_name.startswith('stub_'):
                # stub는 항상 새 셀
                new_value = stub_values.get(field_name, cell.text)
                col_actions[col] = ('new', cell, new_value)

            elif field_name and field_name.startswith('input_'):
                # input 데이터
                new_value = input_values.get(field_name, "")
                col_actions[col] = ('data', cell, new_value)

            elif field_name and field_name.startswith('header_'):
                # header는 확장
                col_actions[col] = ('extend', cell, None)

            elif field_name and field_name.startswith('data_'):
                # data는 빈 셀로 생성
                col_actions[col] = ('data', cell, "")

            else:
                # 기타 - rowspan 확장
                if cell.row < last_row:
                    col_actions[col] = ('extend', cell, None)
                else:
                    col_actions[col] = ('data', cell, "")

        # rowspan 확장 처리
        extended_cells = set()
        for col, (action, cell, _) in col_actions.items():
            if action == 'extend' and cell:
                cell_key = (cell.row, cell.col)
                if cell_key not in extended_cells:
                    self._extend_rowspan(cell)
                    extended_cells.add(cell_key)

        # 새 행 생성
        self._create_row_with_actions(new_row, col_actions)
        self.base_table.row_count += 1

    def _create_row_with_actions(
        self,
        row_idx: int,
        col_actions: Dict[int, Tuple[str, Optional[CellInfo], Optional[str]]]
    ):
        """액션에 따라 새 행 생성"""
        if self.base_table is None or self.base_table.element is None:
            return

        # 마지막 tr 찾기
        last_tr = None
        for child in self.base_table.element:
            if child.tag.endswith('}tr'):
                last_tr = child

        if last_tr is None:
            return

        # 새 tr 생성
        new_tr = copy.deepcopy(last_tr)

        # 기존 셀 제거
        for tc in list(new_tr):
            if tc.tag.endswith('}tc'):
                new_tr.remove(tc)

        # 처리된 열 추적
        processed_cols = set()

        for col in sorted(col_actions.keys()):
            if col in processed_cols:
                continue

            action, ref_cell, value = col_actions[col]

            if action == 'extend':
                # rowspan 확장된 열 - 셀 생성 안 함
                if ref_cell:
                    for c in range(ref_cell.col, ref_cell.col + ref_cell.col_span):
                        processed_cols.add(c)

            elif action == 'new':
                # 새 셀 생성 (stub/gstub 새 값)
                colspan = ref_cell.col_span if ref_cell else 1
                tc = self._create_cell_element(row_idx, col, value or "", colspan=colspan)
                if tc is not None:
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        col_span=colspan,
                        end_col=col + colspan - 1,
                        text=value or "",
                        is_empty=not value,
                        element=tc,
                        field_name=ref_cell.field_name if ref_cell else None
                    )
                    self.base_table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)

            elif action == 'data':
                # 데이터 셀 생성
                colspan = ref_cell.col_span if ref_cell else 1
                tc = self._create_cell_element(row_idx, col, value or "", colspan=colspan)
                if tc is not None:
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        col_span=colspan,
                        end_col=col + colspan - 1,
                        text=value or "",
                        is_empty=not value,
                        element=tc,
                        field_name=ref_cell.field_name if ref_cell else None
                    )
                    self.base_table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)

        # 테이블에 새 행 추가
        self.base_table.element.append(new_tr)

    def add_rows_smart(
        self,
        data_list: List[Dict[str, str]],
        fill_empty_first: bool = True
    ):
        """
        필드명(nc.name) 접두사를 분석하여 자동으로 행 추가

        - 필드명이 'header_'로 시작: 같은 값끼리 병합 (rowspan)
        - 필드명이 'data_'로 시작: 개별 데이터로 처리
        - 필드명이 'add_'로 시작: 기존 셀에 내용 추가 (행 추가 없음)
        - 그 외: 항상 확장 (extend)

        Args:
            data_list: 데이터 리스트 [{field_name: value, ...}, ...]
            fill_empty_first: True면 빈 셀 먼저 채우고, 없으면 새 행 추가

        예시:
            # 테이블에 필드명이 설정되어 있다면:
            # - Col 0: field_name 없음 (항상 확장)
            # - Col 1: field_name="header_group" (헤더 병합)
            # - Col 2: field_name="data_col1" (데이터)
            # - Col 3: field_name="add_memo" (기존 셀에 추가)

            data = [
                {"header_group": "GroupA", "data_col1": "A", "add_memo": "메모1"},
                {"header_group": "GroupA", "data_col1": "C", "add_memo": "메모2"},  # add_memo는 기존 셀에 추가
                {"header_group": "GroupB", "data_col1": "E", "add_memo": "메모3"},
            ]
            merger.add_rows_smart(data)
        """
        if self.base_table is None:
            return

        # 필드명 분석하여 열 분류
        header_cols = []  # 'header_'로 시작하는 필드명의 열
        data_cols = []    # 'data_'로 시작하는 필드명의 열
        add_cols = []     # 'add_'로 시작하는 필드명의 열 (새 기능)
        extend_cols = []  # 그 외 열 (항상 확장)

        for (row, col), cell in self.base_table.cells.items():
            if cell.field_name:
                if cell.field_name.startswith('header_'):
                    if col not in header_cols:
                        header_cols.append(col)
                elif cell.field_name.startswith('data_'):
                    if col not in data_cols:
                        data_cols.append(col)
                elif cell.field_name.startswith('add_'):
                    if col not in add_cols:
                        add_cols.append(col)

        # 필드명이 없거나 header_/data_/add_ 접두사 없는 열 찾기
        for col in range(self.base_table.col_count):
            if col not in header_cols and col not in data_cols and col not in add_cols:
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

        # add_ 필드 데이터 분리
        add_field_data = {}
        row_data_list = []

        for data in data_list:
            add_data = {}
            row_data = {}
            for field_name, value in data.items():
                if field_name.startswith('add_'):
                    add_data[field_name] = value
                else:
                    row_data[field_name] = value
            if add_data:
                add_field_data.update(add_data)
            if row_data:
                row_data_list.append(row_data)

        # add_ 필드 데이터 처리 (기존 셀에 내용 추가)
        self._process_add_fields(add_field_data)

        # add_rows_auto 호출 (add_ 필드 제외한 데이터)
        if header_field_name and row_data_list:
            self.add_rows_auto(
                row_data_list,
                header_col=header_col,
                data_cols=data_cols,
                extend_cols=extend_cols,
                header_key=header_field_name,
                fill_empty_first=fill_empty_first
            )
        elif row_data_list:
            # 헤더 필드가 없으면 단순히 행 추가
            for data in row_data_list:
                self._add_row_with_data(data)

    def _process_add_fields(
        self,
        add_field_data: Dict[str, str],
        separator: str = " ",
        field_styles: Optional[Dict[str, str]] = None
    ):
        """
        add_ 접두사 필드 처리: 기존 셀에 내용 추가

        Args:
            add_field_data: {field_name: value} 딕셔너리
            separator: 기존 내용과 새 내용 사이 구분자 (기본: 빈칸 1개)
            field_styles: {field_name: style_type} 필드별 스타일 지정
        """
        if not self.base_table:
            return

        field_styles = field_styles or {}

        for field_name, value in add_field_data.items():
            # 형식 검증 (validate_format=True인 경우)
            if self.field_validator:
                style = field_styles.get(field_name, "plain")
                validation_result = self.field_validator.validate_add_content(
                    value,
                    base_cell_style=style,
                    separator=separator
                )
                if validation_result.changes_made:
                    print(f"  {field_name}: {', '.join(validation_result.changes_made)}")
                if validation_result.warnings:
                    for warning in validation_result.warnings:
                        print(f"  경고: {warning}")
                value = validation_result.validated_text

            # 필드명으로 셀 찾기
            cells = self.base_table.get_cells_by_field(field_name)

            for cell in cells:
                # 기존 내용에 새 내용 추가 (같은 문단 = 빈칸 1개 구분)
                if cell.text:
                    new_text = f"{cell.text}{separator}{value}"
                else:
                    new_text = value

                self._set_cell_text(cell, new_text)
                cell.text = new_text
                cell.is_empty = False

    def append_to_cell(
        self,
        field_name: str,
        value: str,
        separator: str = "\n",
        all_cells: bool = False
    ):
        """
        특정 필드명의 셀에 내용 추가 (행 추가 없음)

        Args:
            field_name: 필드명 (add_ 접두사 유무 무관)
            value: 추가할 값
            separator: 기존 내용과 새 내용 사이 구분자 (기본: 줄바꿈)
            all_cells: True면 같은 필드명의 모든 셀에 추가, False면 첫 번째 셀만

        예시:
            merger.append_to_cell("add_memo", "추가 메모")
            merger.append_to_cell("notes", "새 내용", separator=" / ")
        """
        if not self.base_table:
            return

        cells = self.base_table.get_cells_by_field(field_name)

        for cell in cells:
            if cell.text:
                new_text = f"{cell.text}{separator}{value}"
            else:
                new_text = value

            self._set_cell_text(cell, new_text)
            cell.text = new_text
            cell.is_empty = False

            if not all_cells:
                break

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

        기존 셀을 템플릿으로 사용하고, 기존 열 너비/행 높이를 적용
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

        # 기존 열 너비와 행 높이 가져오기
        # colspan이 1보다 크면 여러 열의 너비 합산
        total_width = 0
        for c in range(col, col + colspan):
            total_width += self.base_table.get_col_width(c)

        # 행 높이 (새 행이므로 기본값 또는 마지막 행 참조)
        cell_height = self.base_table.get_row_height(self.base_table.row_count - 1)

        # 속성 업데이트
        for child in tc:
            tag = child.tag.split('}')[-1]

            if tag == 'cellAddr':
                child.set('colAddr', str(col))
                child.set('rowAddr', str(row))

            elif tag == 'cellSpan':
                child.set('colSpan', str(colspan))
                child.set('rowSpan', str(rowspan))

            elif tag == 'cellSz':
                # 셀 크기를 기존 열 너비에 맞춤
                child.set('width', str(total_width))
                child.set('height', str(cell_height))

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
        """
        if self.base_table is None or self.base_table.element is None:
            return

        # 마지막 행 찾기
        last_row_idx = self.base_table.row_count - 1

        # 각 열의 rowspan 상태 확인
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
                pass
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

    def save(self, output_path: Union[str, Path]):
        """병합된 테이블을 HWPX 파일로 저장"""
        if not self.base_table:
            raise ValueError("기준 테이블이 로드되지 않았습니다.")

        output_path = Path(output_path)

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
        root = ET.parse(BytesIO(original_content)).getroot()
        return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
