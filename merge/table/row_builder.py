# -*- coding: utf-8 -*-
"""
테이블 행 자동 생성 모듈

헤더 기반 자동 병합, 스마트 행 추가 등의 기능을 제공합니다.

주요 기능:
- add_rows_smart: 필드명 접두사 분석하여 자동 행 추가
- add_rows_auto: 헤더 이름 기준 자동 행 추가
- add_row_with_headers: 헤더 설정에 따른 행 추가
"""

import copy
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, TYPE_CHECKING

from .models import CellInfo, HeaderConfig

if TYPE_CHECKING:
    from .models import TableInfo


class RowBuilder:
    """테이블 행 자동 생성"""

    def __init__(
        self,
        table: "TableInfo",
        extend_rowspan_callback,
        create_cell_callback,
        set_cell_text_callback
    ):
        """
        Args:
            table: 대상 테이블 정보
            extend_rowspan_callback: rowspan 확장 콜백
            create_cell_callback: 셀 생성 콜백
            set_cell_text_callback: 셀 텍스트 설정 콜백
        """
        self.table = table
        self._extend_rowspan = extend_rowspan_callback
        self._create_cell_element = create_cell_callback
        self._set_cell_text = set_cell_text_callback

    def add_rows_smart(
        self,
        data_list: List[Dict[str, str]],
        fill_empty_first: bool = True,
        process_add_fields_callback=None
    ):
        """
        필드명(nc.name) 접두사를 분석하여 자동으로 행 추가

        - 필드명이 'header_'로 시작: 같은 값끼리 병합 (rowspan)
        - 필드명이 'data_'로 시작: 개별 데이터로 처리
        - 필드명이 'add_'로 시작: 기존 셀에 내용 추가 (행 추가 없음)
        - 그 외: 항상 확장 (extend)
        """
        if self.table is None:
            return

        # 필드명 분석하여 열 분류
        header_cols = []  # 'header_'로 시작하는 필드명의 열
        data_cols = []    # 'data_'로 시작하는 필드명의 열
        add_cols = []     # 'add_'로 시작하는 필드명의 열
        extend_cols = []  # 그 외 열 (항상 확장)

        for (row, col), cell in self.table.cells.items():
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
        for col in range(self.table.col_count):
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
        for (row, col), cell in self.table.cells.items():
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
        if process_add_fields_callback and add_field_data:
            process_add_fields_callback(add_field_data)

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
        """
        if self.table is None:
            return

        # extend_cols 기본값 설정
        if extend_cols is None:
            extend_cols = [c for c in range(self.table.col_count)
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
            last_row = self.table.row_count - 1

            # 마지막 행의 header_col 셀 확인
            header_cell = self.table.get_cell(last_row, header_col)
            existing_header_text = header_cell.text if header_cell else ""

            # 헤더 설정 생성
            header_config = []

            for col in range(self.table.col_count):
                cell = self.table.get_cell(last_row, col)

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
        """
        if self.table is None:
            return False

        # 같은 헤더 아래 빈 셀 찾기
        for row in range(self.table.row_count):
            # 이 행의 헤더 확인
            header_cell = self.table.get_cell(row, header_col)
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
                    cell = self.table.get_cell(check_row, col)
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
                            if field_name in self.table.field_to_cell:
                                field_row, field_col = self.table.field_to_cell[field_name]
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
        """
        if self.table is None or self.table.element is None:
            return

        last_row_idx = self.table.row_count - 1
        new_row_idx = self.table.row_count

        # 1. 헤더 설정 분석
        col_actions = {}  # col -> HeaderConfig
        for hc in header_config:
            for c in range(hc.col, hc.col + hc.col_span):
                col_actions[c] = hc

        # 2. rowspan 확장 처리
        for hc in header_config:
            if hc.action == 'extend':
                cell = self.table.get_cell(last_row_idx, hc.col)
                if cell:
                    # rowspan 셀이든 일반 셀이든 모두 확장
                    self._extend_rowspan(cell)

        # 3. 새 행 XML 생성
        self._create_new_row_with_headers(new_row_idx, data, header_config)

        self.table.row_count += 1

    def _create_new_row_with_headers(
        self,
        row_idx: int,
        data: Dict[str, str],
        header_config: List[HeaderConfig]
    ):
        """헤더 설정에 따라 새 행 XML 생성"""
        if self.table is None or self.table.element is None:
            return

        # 마지막 tr 찾기 (템플릿)
        last_tr = None
        for child in self.table.element:
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
        for field_name, (_, col) in self.table.field_to_cell.items():
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
                    self.table.cells[(row_idx, hc.col)] = cell

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
                    self.table.cells[(row_idx, hc.col)] = cell

                for c in range(hc.col, hc.col + hc.col_span):
                    processed_cols.add(c)

        # 테이블에 새 행 추가
        self.table.element.append(new_tr)

    def _add_row_with_data(self, data: Dict[str, str]):
        """
        새 행 추가 (중첩 rowspan 고려)
        """
        if self.table is None or self.table.element is None:
            return

        # 마지막 행 찾기
        last_row_idx = self.table.row_count - 1

        # 각 열의 rowspan 상태 확인
        col_status = {}

        # 모든 열 순회하며 상태 결정
        col = 0
        while col < self.table.col_count:
            cell = self.table.get_cell(last_row_idx, col)

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
        new_row_idx = self.table.row_count

        # 데이터가 들어갈 열과 rowspan 확장할 열 결정
        cols_with_data = {}  # col -> value
        cols_to_extend = set()

        for field_name, value in data.items():
            if field_name in self.table.field_to_cell:
                _, col = self.table.field_to_cell[field_name]
                cols_with_data[col] = value

        # 열별로 처리
        for col in range(self.table.col_count):
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

        self.table.row_count += 1

        # tbl 요소의 rowCnt 속성 업데이트
        if self.table.element is not None:
            self.table.element.set('rowCnt', str(self.table.row_count))

    def _create_new_row(
        self,
        row_idx: int,
        cols_with_data: Dict[int, str],
        col_status: Dict
    ):
        """새 행 XML 생성"""
        if self.table is None or self.table.element is None:
            return

        # 마지막 tr 요소 찾기 (복사 템플릿용)
        last_tr = None
        for child in self.table.element:
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
        for col in range(self.table.col_count):
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
                self.table.cells[(row_idx, col)] = cell

        # 테이블에 새 행 추가
        self.table.element.append(new_tr)

    def add_row_with_stub(
        self,
        data: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str]
    ):
        """stub/gstub 고려하여 새 행 추가"""
        if self.table is None or self.table.element is None:
            return

        last_row = self.table.row_count - 1
        new_row = self.table.row_count

        # 각 열별 셀 타입 분류
        input_cells_by_col: Dict[int, CellInfo] = {}
        gstub_cells_by_col: Dict[int, CellInfo] = {}
        stub_cells_by_col: Dict[int, CellInfo] = {}

        for (r, c), cell in self.table.cells.items():
            if cell.field_name:
                if cell.field_name.startswith('input_'):
                    if c not in input_cells_by_col:
                        input_cells_by_col[c] = cell
                elif cell.field_name.startswith('gstub_'):
                    # gstub는 end_row가 가장 큰 셀 저장
                    if c not in gstub_cells_by_col or cell.end_row > gstub_cells_by_col[c].end_row:
                        gstub_cells_by_col[c] = cell
                elif cell.field_name.startswith('stub_'):
                    if c not in stub_cells_by_col:
                        stub_cells_by_col[c] = cell

        # 각 열의 처리 방법 결정
        col_actions = {}  # col -> ('extend'|'new'|'data', cell, value)

        for col in range(self.table.col_count):
            # 열 타입 결정: gstub > stub > input > 기타
            if col in gstub_cells_by_col:
                ref_cell = gstub_cells_by_col[col]
                field_name = ref_cell.field_name
            elif col in stub_cells_by_col:
                ref_cell = stub_cells_by_col[col]
                field_name = ref_cell.field_name
            elif col in input_cells_by_col:
                ref_cell = input_cells_by_col[col]
                field_name = ref_cell.field_name
            else:
                ref_cell = self.table.get_cell(last_row, col)
                if ref_cell is None:
                    continue
                field_name = ref_cell.field_name

            if field_name and field_name.startswith('gstub_'):
                # gstub 처리
                gstub_cell = None
                for (r, c), cell in self.table.cells.items():
                    if c == col and cell.field_name == field_name:
                        if gstub_cell is None or cell.end_row > gstub_cell.end_row:
                            gstub_cell = cell

                if field_name in gstub_values:
                    new_value = gstub_values[field_name]
                    if gstub_cell and gstub_cell.text == new_value:
                        col_actions[col] = ('extend', gstub_cell, None)
                    else:
                        col_actions[col] = ('new', ref_cell, new_value)
                elif stub_values:
                    col_actions[col] = ('new', ref_cell, "")
                else:
                    col_actions[col] = ('data', ref_cell, "")

            elif field_name and field_name.startswith('stub_'):
                new_value = stub_values.get(field_name, ref_cell.text)
                col_actions[col] = ('new', ref_cell, new_value)

            elif field_name and field_name.startswith('input_'):
                new_value = input_values.get(field_name, "")
                col_actions[col] = ('data', ref_cell, new_value)

            elif field_name and field_name.startswith('header_'):
                continue

            elif field_name and field_name.startswith('data_'):
                col_actions[col] = ('data', ref_cell, "")

            else:
                last_cell = self.table.get_cell(last_row, col)
                if last_cell and last_cell.row < last_row:
                    col_actions[col] = ('extend', last_cell, None)
                else:
                    col_actions[col] = ('data', ref_cell, "")

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
        self.table.row_count += 1

        # tbl 요소의 rowCnt 속성 업데이트
        self.table.element.set('rowCnt', str(self.table.row_count))

    def _create_row_with_actions(
        self,
        row_idx: int,
        col_actions: Dict
    ):
        """액션에 따라 새 행 생성"""
        if self.table is None or self.table.element is None:
            return

        # 마지막 tr 찾기
        last_tr = None
        for child in self.table.element:
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
                    self.table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)

            elif action == 'data':
                # 데이터 셀 생성
                colspan = ref_cell.col_span if ref_cell else 1
                tc = self._create_cell_element(row_idx, col, value or "", rowspan=1, colspan=colspan)
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
                    self.table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)

        # 테이블에 새 행 추가
        self.table.element.append(new_tr)
