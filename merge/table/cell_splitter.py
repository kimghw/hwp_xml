# -*- coding: utf-8 -*-
"""
gstub 셀 나누기 모듈

gstub 범위 내에서 행을 삽입하고 rowspan을 확장하는 기능을 제공합니다.

주요 기능:
- gstub 범위 내 새 행 삽입
- 기존 행들을 밀어내고 rowspan 확장
- 새 행의 셀 생성 (input_, stub_, data_ 등)
"""

import copy
from typing import Dict, Optional, TYPE_CHECKING

from .models import CellInfo

if TYPE_CHECKING:
    from .models import TableInfo


class GstubCellSplitter:
    """gstub 셀 나누기 처리"""

    def __init__(self, table: "TableInfo"):
        """
        Args:
            table: 대상 테이블 정보
        """
        self.table = table

    def insert_row_in_gstub_range(
        self,
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
        extend_rowspan_callback,
        create_cell_callback
    ) -> bool:
        """
        gstub 범위 내에 새 행 삽입 (셀 나누기)

        gstub 셀의 rowspan을 확장하고 해당 범위 바로 다음에 새 행을 삽입합니다.
        새 행의 input 셀은 동일한 nc_name을 갖습니다.
        기존 행들은 rowAddr이 밀려납니다.

        Args:
            gstub_values: {field_name: value} gstub 필드 데이터
            stub_values: {field_name: value} stub 필드 데이터
            input_values: {field_name: value} input 필드 데이터
            extend_rowspan_callback: rowspan 확장 콜백 함수
            create_cell_callback: 셀 생성 콜백 함수

        Returns:
            True: 행 삽입 성공
            False: 삽입 실패 (gstub가 없거나 값이 다름)
        """
        if self.table is None or self.table.element is None:
            return False

        # 같은 gstub 값을 가진 gstub 셀 찾기
        target_gstub_cell = None
        gstub_field_name = None

        for field_name, expected_value in gstub_values.items():
            cells = self.table.get_cells_by_field(field_name)
            for cell in cells:
                if cell.text == expected_value:
                    # 가장 end_row가 큰 셀 선택
                    if target_gstub_cell is None or cell.end_row > target_gstub_cell.end_row:
                        target_gstub_cell = cell
                        gstub_field_name = field_name
                    break

        if target_gstub_cell is None:
            return False  # 같은 값의 gstub 없음 - 새 gstub 생성 필요

        # 새 행 인덱스는 gstub 범위 바로 다음 (end_row + 1)
        insert_row_idx = target_gstub_cell.end_row + 1

        # 삽입 위치 이후의 행들을 밀어냄
        self._shift_rows_down(insert_row_idx)

        # gstub 셀의 rowspan 확장 (셀 나누기)
        extend_rowspan_callback(target_gstub_cell)

        # 테이블 행 수 증가
        self.table.row_count += 1
        self.table.element.set('rowCnt', str(self.table.row_count))

        # 각 열별로 새 셀 생성 (gstub 열 제외)
        self._create_row_for_gstub_extension(
            insert_row_idx,
            target_gstub_cell.col,  # gstub 열
            gstub_field_name,
            stub_values,
            input_values,
            create_cell_callback
        )

        return True

    def _create_row_for_gstub_extension(
        self,
        row_idx: int,
        gstub_col: int,
        gstub_field_name: str,
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
        create_cell_callback
    ):
        """gstub 확장을 위한 새 행 생성"""
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

        # 기존 셀 제거
        for tc in list(new_tr):
            if tc.tag.endswith('}tc'):
                new_tr.remove(tc)

        # 각 열별 참조 셀 및 필드명 수집
        # 우선순위: gstub > input > stub > data > header
        col_to_ref_cell: Dict[int, CellInfo] = {}
        col_to_field_name: Dict[int, str] = {}

        def get_priority(field_name: str) -> int:
            if field_name.startswith('gstub_'):
                return 5
            elif field_name.startswith('input_'):
                return 4
            elif field_name.startswith('stub_'):
                return 3
            elif field_name.startswith('data_'):
                return 2
            elif field_name.startswith('header_'):
                return 1
            return 0

        for (r, c), cell in self.table.cells.items():
            if cell.field_name:
                current_priority = get_priority(col_to_field_name.get(c, ""))
                new_priority = get_priority(cell.field_name)
                if new_priority > current_priority:
                    col_to_field_name[c] = cell.field_name
                    col_to_ref_cell[c] = cell

        # 처리된 열 추적
        processed_cols = set()

        # 열 순서대로 셀 생성
        for col in range(self.table.col_count):
            if col in processed_cols:
                continue

            field_name = col_to_field_name.get(col, "")
            ref_cell = col_to_ref_cell.get(col)

            # gstub 열은 rowspan으로 커버되므로 셀 생성 안 함
            if col == gstub_col or (field_name and field_name.startswith('gstub_')):
                # gstub 열의 colspan 범위 처리
                if ref_cell:
                    for c in range(ref_cell.col, ref_cell.col + ref_cell.col_span):
                        processed_cols.add(c)
                else:
                    processed_cols.add(col)
                continue

            # header_ 열은 셀 생성 안 함 (rowspan으로 커버됨)
            if field_name and field_name.startswith('header_'):
                processed_cols.add(col)
                continue

            # data_ 열은 빈 셀 생성
            if field_name and field_name.startswith('data_'):
                tc = create_cell_callback(row_idx, col, "", rowspan=1, colspan=1)
                if tc is not None:
                    # 동일한 nc_name 설정
                    tc.set('name', field_name)
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        text="",
                        is_empty=True,
                        element=tc,
                        field_name=field_name
                    )
                    self.table.cells[(row_idx, col)] = cell
                processed_cols.add(col)
                continue

            # stub_ 열은 새 셀 생성 (값 포함)
            if field_name and field_name.startswith('stub_'):
                value = stub_values.get(field_name, "")
                colspan = ref_cell.col_span if ref_cell else 1
                tc = create_cell_callback(row_idx, col, value, rowspan=1, colspan=colspan)
                if tc is not None:
                    # 동일한 nc_name 설정
                    tc.set('name', field_name)
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        col_span=colspan,
                        end_col=col + colspan - 1,
                        text=value,
                        is_empty=not value,
                        element=tc,
                        field_name=field_name
                    )
                    self.table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)
                continue

            # input_ 열은 데이터 셀 생성 (동일한 nc_name)
            if field_name and field_name.startswith('input_'):
                value = input_values.get(field_name, "")
                colspan = ref_cell.col_span if ref_cell else 1
                tc = create_cell_callback(row_idx, col, value, rowspan=1, colspan=colspan)
                if tc is not None:
                    # 동일한 nc_name 설정
                    tc.set('name', field_name)
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        col_span=colspan,
                        end_col=col + colspan - 1,
                        text=value,
                        is_empty=not value,
                        element=tc,
                        field_name=field_name
                    )
                    self.table.cells[(row_idx, col)] = cell

                    # field_to_cell에도 추가
                    if field_name not in self.table.field_to_cell:
                        self.table.field_to_cell[field_name] = (row_idx, col)

                for c in range(col, col + colspan):
                    processed_cols.add(c)
                continue

            # 기타 열 - 빈 셀 생성
            if ref_cell:
                colspan = ref_cell.col_span
                tc = create_cell_callback(row_idx, col, "", rowspan=1, colspan=colspan)
                if tc is not None:
                    if field_name:
                        tc.set('name', field_name)
                    new_tr.append(tc)

                    cell = CellInfo(
                        row=row_idx,
                        col=col,
                        col_span=colspan,
                        end_col=col + colspan - 1,
                        text="",
                        is_empty=True,
                        element=tc,
                        field_name=field_name if field_name else None
                    )
                    self.table.cells[(row_idx, col)] = cell

                for c in range(col, col + colspan):
                    processed_cols.add(c)

        # 테이블에 새 행 삽입 (지정된 위치에)
        self._insert_tr_at_position(new_tr, row_idx)

    def _insert_tr_at_position(self, new_tr, row_idx: int):
        """tr 요소를 지정된 행 위치에 삽입"""
        if self.table is None or self.table.element is None:
            return

        # 현재 tr 요소들 수집
        tr_elements = []
        for child in self.table.element:
            if child.tag.endswith('}tr'):
                tr_elements.append(child)

        # row_idx 위치에 삽입
        if row_idx >= len(tr_elements):
            # 끝에 추가
            self.table.element.append(new_tr)
        else:
            # 지정 위치에 삽입
            # tbl 내에서 tr의 위치 찾기
            insert_idx = 0
            tr_count = 0
            for idx, child in enumerate(list(self.table.element)):
                if child.tag.endswith('}tr'):
                    if tr_count == row_idx:
                        insert_idx = idx
                        break
                    tr_count += 1
            self.table.element.insert(insert_idx, new_tr)

    def _shift_rows_down(self, from_row: int):
        """
        지정된 행부터 아래의 모든 행을 한 칸씩 밀어냄

        Args:
            from_row: 밀어낼 시작 행 인덱스
        """
        if self.table is None or self.table.element is None:
            return

        # 1. XML의 cellAddr 업데이트
        for child in self.table.element:
            if child.tag.endswith('}tr'):
                for tc in child:
                    if tc.tag.endswith('}tc'):
                        for tc_child in tc:
                            if tc_child.tag.endswith('}cellAddr'):
                                row_addr = int(tc_child.get('rowAddr', 0))
                                if row_addr >= from_row:
                                    tc_child.set('rowAddr', str(row_addr + 1))

        # 2. 메모리상 cells 딕셔너리 업데이트
        new_cells = {}
        for (r, c), cell in self.table.cells.items():
            if r >= from_row:
                # 행 인덱스를 1 증가
                cell.row = r + 1
                cell.end_row = cell.row + cell.row_span - 1
                new_cells[(r + 1, c)] = cell
            else:
                # 그대로 유지하되, rowspan이 from_row를 포함하면 end_row 업데이트
                if cell.end_row >= from_row:
                    # rowspan 셀은 end_row만 증가 (row_span은 _extend_rowspan에서 처리)
                    pass
                new_cells[(r, c)] = cell
        self.table.cells = new_cells

        # 3. field_to_cell 매핑 업데이트
        new_field_to_cell = {}
        for field_name, (r, c) in self.table.field_to_cell.items():
            if r >= from_row:
                new_field_to_cell[field_name] = (r + 1, c)
            else:
                new_field_to_cell[field_name] = (r, c)
        self.table.field_to_cell = new_field_to_cell
