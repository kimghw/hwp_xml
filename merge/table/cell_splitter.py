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
from typing import Dict, Optional, Tuple, TYPE_CHECKING

from .models import CellInfo

if TYPE_CHECKING:
    from .models import TableInfo


# 필드 접두사 우선순위
FIELD_PRIORITY = {"gstub_": 5, "input_": 4, "stub_": 3, "data_": 2, "header_": 1}


def get_field_prefix(field_name: Optional[str]) -> str:
    """필드명에서 접두사 추출"""
    if not field_name or "_" not in field_name:
        return ""
    return field_name.split("_")[0] + "_"


class GstubCellSplitter:
    """gstub 셀 나누기 처리"""

    def __init__(self, table: "TableInfo"):
        """
        Args:
            table: 대상 테이블 정보
        """
        self.table = table

    # ========== 공통 헬퍼 메서드 ==========

    def _get_last_tr(self):
        """마지막 tr 요소 반환"""
        if self.table is None or self.table.element is None:
            return None
        last_tr = None
        for child in self.table.element:
            if child.tag.endswith("}tr"):
                last_tr = child
        return last_tr

    def _create_empty_tr(self):
        """빈 tr 요소 생성"""
        last_tr = self._get_last_tr()
        if last_tr is None:
            return None
        new_tr = copy.deepcopy(last_tr)
        for tc in list(new_tr):
            if tc.tag.endswith("}tc"):
                new_tr.remove(tc)
        return new_tr

    def _collect_col_info(self) -> Dict[int, Tuple[str, CellInfo]]:
        """열별로 가장 우선순위 높은 셀 수집"""
        col_info: Dict[int, Tuple[str, CellInfo]] = {}

        for (r, c), cell in self.table.cells.items():
            prefix = get_field_prefix(cell.field_name)
            if not prefix:
                continue

            current = col_info.get(c)
            current_priority = FIELD_PRIORITY.get(current[0], 0) if current else 0
            new_priority = FIELD_PRIORITY.get(prefix, 0)

            if new_priority > current_priority:
                col_info[c] = (prefix, cell)

        return col_info

    def _create_cell_for_row(
        self,
        new_tr,
        row_idx: int,
        col: int,
        value: str,
        field_name: Optional[str],
        colspan: int,
        create_cell_callback,
    ) -> Optional[CellInfo]:
        """새 행에 셀 생성하고 테이블에 등록"""
        tc = create_cell_callback(row_idx, col, value, rowspan=1, colspan=colspan)
        if tc is None:
            return None

        if field_name:
            tc.set("name", field_name)
        new_tr.append(tc)

        cell = CellInfo(
            row=row_idx,
            col=col,
            col_span=colspan,
            end_col=col + colspan - 1,
            text=value,
            is_empty=not value,
            element=tc,
            field_name=field_name,
        )
        self.table.cells[(row_idx, col)] = cell
        return cell

    def insert_row_in_gstub_range(
        self,
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
        extend_rowspan_callback,
        create_cell_callback,
    ) -> bool:
        """
        gstub 범위 내에 새 행 삽입 (셀 나누기)

        모든 gstub 셀의 rowspan을 확장하고 해당 범위 바로 다음에 새 행을 삽입합니다.
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

        # 모든 gstub 열에서 매칭되는 셀 찾기
        matching_gstub_cells: Dict[int, CellInfo] = {}  # col -> cell
        gstub_cols: set = set()

        for field_name, expected_value in gstub_values.items():
            cells = self.table.get_cells_by_field(field_name)
            for cell in cells:
                if cell.text == expected_value:
                    # 해당 열에서 가장 end_row가 큰 셀 선택
                    if cell.col not in matching_gstub_cells or cell.end_row > matching_gstub_cells[cell.col].end_row:
                        matching_gstub_cells[cell.col] = cell
                        gstub_cols.add(cell.col)
                    break

        if not matching_gstub_cells:
            return False  # 같은 값의 gstub 없음 - 새 gstub 생성 필요

        # 모든 gstub의 공통 범위에서 삽입 위치 결정
        # 공통 범위 = 모든 gstub가 커버하는 행의 교집합
        # 삽입 위치 = 가장 작은 end_row + 1 (가장 좁은 gstub 범위 바로 다음)
        min_end_row = min(cell.end_row for cell in matching_gstub_cells.values())
        insert_row_idx = min_end_row + 1

        # 삽입 위치 이후의 행들을 밀어냄
        self._shift_rows_down(insert_row_idx)

        # 삽입 위치가 범위 안에 있는 gstub 셀만 rowspan 확장
        # (end_row >= insert_row_idx - 1인 셀들)
        for cell in matching_gstub_cells.values():
            # 모든 매칭 gstub는 확장 필요 (새 행이 범위 안에 들어가므로)
            extend_rowspan_callback(cell)

        # 테이블 행 수 증가
        self.table.row_count += 1
        self.table.element.set("rowCnt", str(self.table.row_count))

        # 각 열별로 새 셀 생성 (모든 gstub 열 제외)
        self._create_row_for_gstub_extension(
            insert_row_idx,
            gstub_cols,  # 모든 gstub 열
            stub_values,
            input_values,
            create_cell_callback,
        )

        return True

    def _create_row_for_gstub_extension(
        self,
        row_idx: int,
        gstub_cols: set,
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
        create_cell_callback,
    ):
        """gstub 확장을 위한 새 행 생성"""
        if self.table is None or self.table.element is None:
            return

        new_tr = self._create_empty_tr()
        if new_tr is None:
            return

        col_info = self._collect_col_info()
        processed_cols = set()

        for col in range(self.table.col_count):
            if col in processed_cols:
                continue

            info = col_info.get(col)
            prefix = info[0] if info else ""
            ref_cell = info[1] if info else None
            field_name = ref_cell.field_name if ref_cell else ""
            colspan = ref_cell.col_span if ref_cell else 1

            # 셀 생성 여부 및 값 결정
            should_create, value = self._get_cell_action(
                col, gstub_cols, prefix, field_name, stub_values, input_values
            )

            if should_create:
                cell = self._create_cell_for_row(
                    new_tr, row_idx, col, value, field_name, colspan, create_cell_callback
                )
                # input_ 필드는 field_to_cell에 추가
                if cell and prefix == "input_" and field_name not in self.table.field_to_cell:
                    self.table.field_to_cell[field_name] = (row_idx, col)

            # colspan 범위 처리됨으로 표시
            for c in range(col, col + colspan):
                processed_cols.add(c)

        self._insert_tr_at_position(new_tr, row_idx)

    def _get_cell_action(
        self,
        col: int,
        gstub_cols: set,
        prefix: str,
        field_name: str,
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
    ) -> Tuple[bool, str]:
        """
        열에 대한 셀 생성 여부와 값 결정

        Args:
            col: 현재 열 인덱스
            gstub_cols: 모든 gstub 열 인덱스 집합

        Returns:
            (should_create, value) 튜플
        """
        # gstub/header 열은 rowspan으로 커버되므로 셀 생성 안 함
        if col in gstub_cols or prefix == "gstub_" or prefix == "header_":
            return False, ""

        # 타입별 값 결정
        if prefix == "stub_":
            return True, stub_values.get(field_name, "")
        elif prefix == "input_":
            return True, input_values.get(field_name, "")
        else:
            # data_ 또는 기타 - 빈 셀
            return True, ""

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
                    cell.end_row += 1
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
