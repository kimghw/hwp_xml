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
from typing import List, Dict, Optional, Union, Any
from pathlib import Path
from io import BytesIO
import copy

from .models import CellInfo, HeaderConfig, TableInfo
from .parser import TableParser, NAMESPACES
from .gstub_cell_splitter import GstubCellSplitter
from .row_builder import RowBuilder
from .formatter_config import TableFormatterConfigLoader
from ..format_validator import AddFieldValidator
from ..content_formatter import ContentFormatter

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class TableMerger:
    """테이블 셀 내용 병합"""

    def __init__(
        self,
        validate_format: bool = False,
        sdk_validator=None,
        formatter_config_path: Optional[str] = None,
        use_formatter: bool = True,
        format_add_content: bool = True,
        use_sdk_for_levels: bool = True,
    ):
        """
        Args:
            validate_format: True면 add_/input_ 필드 병합 전 형식 검증
            sdk_validator: Claude Code SDK 검증 함수 (외부 주입)
            formatter_config_path: 포맷터 설정 YAML 파일 경로 (기본: table_formatter_config.yaml)
            use_formatter: True면 add_ 필드에 포맷터 적용
            format_add_content: True면 add_ 필드에 글머리 기호 포맷팅 적용
            use_sdk_for_levels: True면 SDK로 레벨 분석 (format_add_content=True일 때)
        """
        self.parser = TableParser()
        self.base_table: Optional[TableInfo] = None
        self.hwpx_path: Optional[Path] = None
        self.hwpx_data: Dict[str, bytes] = {}  # HWPX 파일 내용
        self.validate_format = validate_format
        self.field_validator = AddFieldValidator(sdk_validator) if validate_format else None

        # 포맷터 설정 로드
        self.use_formatter = use_formatter
        self.formatter_loader: Optional[TableFormatterConfigLoader] = None
        if use_formatter:
            self.formatter_loader = TableFormatterConfigLoader(formatter_config_path)
            self.formatter_loader.load()

        # add_ 필드 글머리 포맷터
        self.format_add_content = format_add_content
        self.use_sdk_for_levels = use_sdk_for_levels
        self.content_formatter: Optional[ContentFormatter] = None
        if format_add_content:
            self.content_formatter = ContentFormatter(style="default", use_sdk=use_sdk_for_levels)

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

        # 1. 빈 셀 먼저 채우기 시도 (gstub 범위 내)
        if fill_empty_first:
            filled = self._try_fill_input_cells(input_values, gstub_values, stub_values)
            if filled:
                return

        # 2. gstub가 있으면 gstub 범위 확장하여 행 삽입
        if gstub_values:
            inserted = self._insert_row_in_gstub_range(data, gstub_values, stub_values, input_values)
            if inserted:
                return

        # 3. gstub 없으면 테이블 끝에 새 행 추가
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

        # 모든 gstub 값이 매칭되는 공통 행 범위 계산
        valid_rows = set(range(self.base_table.row_count))

        for field_name, expected_value in gstub_values.items():
            cells = self.base_table.get_cells_by_field(field_name)
            gstub_rows = set()
            for cell in cells:
                if cell.text == expected_value:
                    # 이 gstub가 커버하는 행들
                    gstub_rows.update(range(cell.row, cell.end_row + 1))
            # 공통 범위로 축소
            valid_rows &= gstub_rows

        if not valid_rows:
            return False  # 매칭되는 gstub 범위 없음

        # 유효한 행들 중에서 빈 셀 찾기
        for row in sorted(valid_rows):
            # stub 매칭 확인
            matching_stub = True
            for field_name, expected_value in stub_values.items():
                cells = self.base_table.get_cells_by_field(field_name)
                found = False
                for cell in cells:
                    if cell.row == row:
                        found = True
                        if cell.text != expected_value:
                            matching_stub = False
                        break
                if not found:
                    matching_stub = False

            if not matching_stub:
                continue

            # input 셀 빈 여부 확인
            row_empty = True
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

    def _insert_row_in_gstub_range(
        self,
        data: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str]
    ) -> bool:
        """
        gstub 범위 내에 새 행 삽입 (셀 나누기)

        GstubCellSplitter에 위임합니다.
        """
        if self.base_table is None:
            return False

        splitter = GstubCellSplitter(self.base_table)
        return splitter.insert_row_in_gstub_range(
            gstub_values,
            stub_values,
            input_values,
            self._extend_rowspan,
            self._create_cell_element
        )

    def _add_row_with_stub(
        self,
        data: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str]
    ):
        """stub/gstub 고려하여 새 행 추가 - RowBuilder에 위임"""
        if self.base_table is None:
            return

        builder = RowBuilder(
            self.base_table,
            self._extend_rowspan,
            self._create_cell_element,
            self._set_cell_text
        )
        builder.add_row_with_stub(data, gstub_values, stub_values, input_values)

    def add_rows_smart(
        self,
        data_list: List[Dict[str, str]],
        fill_empty_first: bool = True
    ):
        """
        필드명(nc.name) 접두사를 분석하여 자동으로 행 추가

        RowBuilder에 위임합니다.
        """
        if self.base_table is None:
            return

        builder = RowBuilder(
            self.base_table,
            self._extend_rowspan,
            self._create_cell_element,
            self._set_cell_text
        )
        builder.add_rows_smart(
            data_list,
            fill_empty_first,
            self._process_add_fields
        )

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
            # 필드별 구분자 초기화 (기본값으로 리셋)
            field_separator = separator

            # 1. 글머리 기호 포맷팅 적용 (format_add_content=True인 경우)
            # SDK가 글머리 제거 + 레벨 분석 + 글머리 재적용을 담당
            sdk_formatted = False
            if self.format_add_content and self.content_formatter:
                if self.use_sdk_for_levels:
                    result = self.content_formatter.format_with_analyzed_levels(value)
                else:
                    result = self.content_formatter.auto_format(value, use_sdk_for_levels=False)

                if result.success and result.formatted_text:
                    value = result.formatted_text
                    sdk_formatted = True

            # 2. 포맷터 적용 (use_formatter=True인 경우)
            # SDK가 이미 글머리 포맷팅을 했으면 구분자만 적용
            if self.use_formatter and self.formatter_loader:
                field_config = self.formatter_loader.get_config_for_field(field_name)

                # 필드별 구분자 사용
                if field_config.separator:
                    field_separator = field_config.separator

                # SDK가 포맷팅하지 않은 경우에만 포맷터 적용
                if not sdk_formatted:
                    value = self.formatter_loader.format_value(field_name, value)

            # 3. 형식 검증 (validate_format=True인 경우)
            if self.field_validator:
                style = field_styles.get(field_name, "plain")
                validation_result = self.field_validator.validate_add_content(
                    value,
                    base_cell_style=style,
                    separator=field_separator
                )
                if validation_result.changes_made:
                    print(f"  {field_name}: {', '.join(validation_result.changes_made)}")
                if validation_result.warnings:
                    for warning in validation_result.warnings:
                        print(f"  경고: {warning}")
                value = validation_result.validated_text

            # 4. 필드명으로 셀 찾기
            cells = self.base_table.get_cells_by_field(field_name)

            for cell in cells:
                # 기존 내용에 새 내용 추가 (같은 문단 = 빈칸 1개 구분)
                if cell.text:
                    new_text = f"{cell.text}{field_separator}{value}"
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

        RowBuilder에 위임합니다.
        """
        if self.base_table is None:
            return

        builder = RowBuilder(
            self.base_table,
            self._extend_rowspan,
            self._create_cell_element,
            self._set_cell_text
        )
        builder.add_rows_auto(
            data_list, header_col, data_cols, extend_cols, header_key, fill_empty_first
        )

    def add_row_with_headers(
        self,
        data: Dict[str, str],
        header_config: List[HeaderConfig]
    ):
        """
        헤더 설정에 따라 새 행 추가

        RowBuilder에 위임합니다.
        """
        if self.base_table is None:
            return

        builder = RowBuilder(
            self.base_table,
            self._extend_rowspan,
            self._create_cell_element,
            self._set_cell_text
        )
        builder.add_row_with_headers(data, header_config)

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

        # 템플릿 셀 찾기 (같은 열의 input_ 셀 우선)
        template_cell = None
        fallback_cell = None

        for (r, c), cell in self.base_table.cells.items():
            if c == col and cell.element is not None:
                # input_ 셀을 우선 사용 (데이터 행 스타일)
                if cell.field_name and cell.field_name.startswith('input_'):
                    template_cell = cell
                    break
                # 다른 셀은 fallback으로 저장
                if fallback_cell is None:
                    fallback_cell = cell

        if template_cell is None:
            template_cell = fallback_cell

        if template_cell is None:
            # 아무 input_ 셀이나 템플릿으로 사용
            for cell in self.base_table.cells.values():
                if cell.element is not None:
                    if cell.field_name and cell.field_name.startswith('input_'):
                        template_cell = cell
                        break
            # 그래도 없으면 아무 셀이나
            if template_cell is None:
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
        if self.base_table is None:
            return

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

        # 2단계: 남은 데이터는 행 추가 (RowBuilder 사용)
        builder = RowBuilder(
            self.base_table,
            self._extend_rowspan,
            self._create_cell_element,
            self._set_cell_text
        )
        for data in remaining_data:
            if data not in filled:
                builder._add_row_with_data(data)

    def _set_cell_text(self, cell: CellInfo, text: str):
        """셀에 텍스트 설정 (여러 줄이면 여러 문단으로)"""
        if cell.element is None:
            return

        lines = text.split('\n') if '\n' in text else [text]

        # subList 찾기
        for child in cell.element:
            if child.tag.endswith('}subList'):
                # 기존 문단들 수집
                existing_p = [p for p in child if p.tag.endswith('}p')]

                if not existing_p:
                    # 문단이 없으면 첫 줄만 설정하고 종료
                    cell.text = text
                    return

                # 첫 번째 문단을 템플릿으로 사용
                template_p = existing_p[0]

                # 기존 문단 모두 제거
                for p in existing_p:
                    child.remove(p)

                # 각 줄마다 문단 생성
                for i, line in enumerate(lines):
                    new_p = copy.deepcopy(template_p)

                    # 문단 내 텍스트 설정
                    text_set = False
                    for run in new_p:
                        if run.tag.endswith('}run'):
                            for t in run:
                                if t.tag.endswith('}t'):
                                    t.text = line
                                    text_set = True
                                    break
                            if text_set:
                                break

                            # t 요소가 없으면 생성
                            if not text_set:
                                ns = run.tag.split('}')[0] + '}' if '}' in run.tag else ''
                                t_elem = ET.Element(f'{ns}t')
                                t_elem.text = line
                                run.append(t_elem)
                                text_set = True
                                break

                    child.append(new_p)

                cell.text = text
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
        """section XML 재구성 - 수정된 테이블 요소 반영"""
        root = ET.parse(BytesIO(original_content)).getroot()

        # 원본 테이블 요소들을 찾아서 수정된 테이블로 교체
        if self.base_table and self.base_table.element is not None:
            table_elements = [elem for elem in root.iter() if elem.tag.endswith('}tbl')]

            for i, tbl_elem in enumerate(table_elements):
                # base_table의 element와 동일한 table_id를 가진 테이블 찾기
                if tbl_elem.get('id') == self.base_table.table_id:
                    # 부모 요소 찾기
                    parent = None
                    for p in root.iter():
                        if tbl_elem in list(p):
                            parent = p
                            break

                    if parent is not None:
                        # 기존 테이블 위치에 수정된 테이블 삽입
                        idx = list(parent).index(tbl_elem)
                        parent.remove(tbl_elem)
                        parent.insert(idx, self.base_table.element)
                    break

        return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
