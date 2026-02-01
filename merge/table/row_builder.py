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
from typing import List, Dict, Optional, Tuple, TYPE_CHECKING

from .models import CellInfo, HeaderConfig

if TYPE_CHECKING:
    from .models import TableInfo


def get_field_prefix(field_name: Optional[str]) -> str:
    """필드명에서 접두사 추출 (header_, data_, input_ 등)"""
    if not field_name:
        return ""
    if "_" in field_name:
        return field_name.split("_")[0] + "_"
    return ""


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
        """빈 tr 요소 생성 (마지막 tr 복사 후 셀 제거)"""
        last_tr = self._get_last_tr()
        if last_tr is None:
            return None
        new_tr = copy.deepcopy(last_tr)
        for tc in list(new_tr):
            if tc.tag.endswith("}tc"):
                new_tr.remove(tc)
        return new_tr

    def _create_cell_and_info(
        self,
        row_idx: int,
        col: int,
        value: str,
        field_name: Optional[str] = None,
        rowspan: int = 1,
        colspan: int = 1,
    ) -> Tuple[Optional[ET.Element], Optional[CellInfo]]:
        """셀 XML 요소와 CellInfo 생성"""
        tc = self._create_cell_element(row_idx, col, value, rowspan=rowspan, colspan=colspan)
        if tc is None:
            return None, None

        cell = CellInfo(
            row=row_idx,
            col=col,
            row_span=rowspan,
            col_span=colspan,
            end_row=row_idx + rowspan - 1,
            end_col=col + colspan - 1,
            text=value,
            is_empty=not value,
            element=tc,
            field_name=field_name,
        )
        return tc, cell

    def _append_cell_to_tr(
        self,
        new_tr: ET.Element,
        row_idx: int,
        col: int,
        value: str,
        field_name: Optional[str] = None,
        rowspan: int = 1,
        colspan: int = 1,
    ) -> Optional[CellInfo]:
        """tr에 새 셀 추가하고 테이블에 등록"""
        tc, cell = self._create_cell_and_info(
            row_idx, col, value, field_name, rowspan, colspan
        )
        if tc is None or cell is None:
            return None

        new_tr.append(tc)
        self.table.cells[(row_idx, col)] = cell
        return cell

    def _get_field_col_mapping(self) -> Dict[str, int]:
        """필드명 -> 열 매핑 반환"""
        return {fn: c for fn, (_, c) in self.table.field_to_cell.items()}

    def _collect_cells_by_prefix(self) -> Dict[int, Tuple[str, CellInfo]]:
        """
        열별로 가장 우선순위 높은 셀 수집

        Returns:
            {col: (prefix, cell)} 딕셔너리
        """
        priority = {"gstub_": 5, "input_": 4, "stub_": 3, "data_": 2, "header_": 1}
        col_info: Dict[int, Tuple[str, CellInfo]] = {}

        for (r, c), cell in self.table.cells.items():
            prefix = get_field_prefix(cell.field_name)
            if not prefix:
                continue

            current = col_info.get(c)
            if current is None or priority.get(prefix, 0) > priority.get(current[0], 0):
                col_info[c] = (prefix, cell)
            # gstub는 end_row가 가장 큰 셀 선택
            elif prefix == "gstub_" and current[0] == "gstub_":
                if cell.end_row > current[1].end_row:
                    col_info[c] = (prefix, cell)

        return col_info

    def _finalize_new_row(self, new_tr: ET.Element):
        """새 행 추가 마무리 (row_count 증가, XML 속성 갱신)"""
        if self.table is None or self.table.element is None:
            return
        self.table.element.append(new_tr)
        self.table.row_count += 1
        self.table.element.set("rowCnt", str(self.table.row_count))

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

        # 4. XML 요소의 rowCnt 속성 갱신
        if self.table.element is not None:
            self.table.element.set('rowCnt', str(self.table.row_count))

    def _find_field_name_for_col(self, col: int, prefixes: tuple) -> str:
        """열에 해당하는 필드명 찾기 (지정된 접두사 우선)"""
        for fn, (_, fc) in self.table.field_to_cell.items():
            if fc == col and fn.startswith(prefixes):
                return fn
        return ""

    def _create_new_row_with_headers(
        self,
        row_idx: int,
        data: Dict[str, str],
        header_config: List[HeaderConfig],
    ):
        """헤더 설정에 따라 새 행 XML 생성"""
        if self.table is None or self.table.element is None:
            return

        new_tr = self._create_empty_tr()
        if new_tr is None:
            return

        # 필드명 -> 열 -> 값 매핑
        field_to_col = self._get_field_col_mapping()
        cols_with_data = {field_to_col[fn]: v for fn, v in data.items() if fn in field_to_col}

        processed_cols = set()
        for hc in sorted(header_config, key=lambda x: x.col):
            if hc.col in processed_cols:
                continue

            if hc.action == "extend":
                pass  # rowspan 확장된 셀 - 새 행에 셀 없음
            elif hc.action == "new":
                field_name = self._find_field_name_for_col(hc.col, ("gstub_", "stub_", "header_"))
                self._append_cell_to_tr(
                    new_tr, row_idx, hc.col, hc.text, field_name, hc.rowspan, hc.col_span
                )
            elif hc.action == "data":
                value = cols_with_data.get(hc.col, "")
                field_name = self._find_field_name_for_col(hc.col, ("input_", "data_"))
                self._append_cell_to_tr(
                    new_tr, row_idx, hc.col, value, field_name, colspan=hc.col_span
                )

            for c in range(hc.col, hc.col + hc.col_span):
                processed_cols.add(c)

        self.table.element.append(new_tr)

    def _add_row_with_data(self, data: Dict[str, str]):
        """새 행 추가 (중첩 rowspan 고려)"""
        if self.table is None or self.table.element is None:
            return

        last_row_idx = self.table.row_count - 1
        col_status = self._analyze_col_status(last_row_idx)

        # 필드명 -> 열 -> 값 매핑
        cols_with_data = {}
        for field_name, value in data.items():
            if field_name in self.table.field_to_cell:
                _, col = self.table.field_to_cell[field_name]
                cols_with_data[col] = value

        # 데이터 없는 rowspan 셀 확장
        self._extend_empty_rowspan_cols(col_status, cols_with_data)

        # 새 행 XML 생성
        new_row_idx = self.table.row_count
        self._create_new_row(new_row_idx, cols_with_data, col_status)
        self.table.row_count += 1

        if self.table.element is not None:
            self.table.element.set("rowCnt", str(self.table.row_count))

    def _analyze_col_status(self, last_row_idx: int) -> Dict[int, Tuple[str, Optional[CellInfo]]]:
        """각 열의 rowspan 상태 분석"""
        col_status = {}
        col = 0

        while col < self.table.col_count:
            cell = self.table.get_cell(last_row_idx, col)

            if cell:
                for c in range(cell.col, cell.col + cell.col_span):
                    if c == cell.col:
                        if cell.row < last_row_idx:
                            col_status[c] = ("extend_rowspan", cell)
                        else:
                            col_status[c] = ("new_cell", cell)
                    else:
                        col_status[c] = ("skip", cell)
                col = cell.col + cell.col_span
            else:
                col_status[col] = ("new_cell", None)
                col += 1

        return col_status

    def _extend_empty_rowspan_cols(
        self,
        col_status: Dict[int, Tuple[str, Optional[CellInfo]]],
        cols_with_data: Dict[int, str],
    ):
        """데이터 없는 rowspan 셀 확장"""
        extended_cells = set()

        for col in range(self.table.col_count):
            status, ref_cell = col_status.get(col, ("new_cell", None))
            if status == "skip" or col in cols_with_data:
                continue
            if status == "extend_rowspan" and ref_cell:
                cell_key = (ref_cell.row, ref_cell.col)
                if cell_key not in extended_cells:
                    self._extend_rowspan(ref_cell)
                    extended_cells.add(cell_key)

    def _create_new_row(
        self,
        row_idx: int,
        cols_with_data: Dict[int, str],
        col_status: Dict,
    ):
        """새 행 XML 생성"""
        if self.table is None or self.table.element is None:
            return

        new_tr = self._create_empty_tr()
        if new_tr is None:
            return

        # new_cell 상태인 열에만 셀 생성
        for col in range(self.table.col_count):
            status, _ = col_status.get(col, ("new_cell", None))
            if status == "new_cell":
                value = cols_with_data.get(col, "")
                field_name = self._find_field_name_for_col(col, ("input_", "data_"))
                self._append_cell_to_tr(new_tr, row_idx, col, value, field_name)

        self.table.element.append(new_tr)

    def add_row_with_stub(
        self,
        data: Dict[str, str],
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
    ):
        """stub/gstub 고려하여 새 행 추가"""
        if self.table is None or self.table.element is None:
            return

        last_row = self.table.row_count - 1
        col_info = self._collect_cells_by_prefix()

        # 각 열의 처리 방법 결정
        col_actions = self._determine_col_actions(
            col_info, last_row, gstub_values, stub_values, input_values
        )

        # rowspan 확장 처리
        self._extend_rowspans(col_actions)

        # 새 행 생성
        new_row = self.table.row_count
        self._create_row_with_actions(new_row, col_actions)
        self.table.row_count += 1
        self.table.element.set("rowCnt", str(self.table.row_count))

    def _determine_col_actions(
        self,
        col_info: Dict[int, Tuple[str, CellInfo]],
        last_row: int,
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
    ) -> Dict[int, Tuple[str, CellInfo, Optional[str]]]:
        """각 열의 처리 방법 결정 (extend/new/data)"""
        col_actions = {}

        for col in range(self.table.col_count):
            info = col_info.get(col)
            if info:
                prefix, ref_cell = info
                field_name = ref_cell.field_name
            else:
                ref_cell = self.table.get_cell(last_row, col)
                if ref_cell is None:
                    continue
                prefix = get_field_prefix(ref_cell.field_name)
                field_name = ref_cell.field_name

            action = self._get_action_for_col(
                prefix, field_name, ref_cell, last_row,
                gstub_values, stub_values, input_values
            )
            if action:
                col_actions[col] = action

        return col_actions

    def _get_action_for_col(
        self,
        prefix: str,
        field_name: Optional[str],
        ref_cell: CellInfo,
        last_row: int,
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
        input_values: Dict[str, str],
    ) -> Optional[Tuple[str, CellInfo, Optional[str]]]:
        """열 타입에 따른 액션 결정"""
        if prefix == "gstub_":
            return self._get_gstub_action(field_name, ref_cell, gstub_values, stub_values)
        elif prefix == "stub_":
            value = stub_values.get(field_name, ref_cell.text)
            return ("new", ref_cell, value)
        elif prefix == "input_":
            value = input_values.get(field_name, "")
            return ("data", ref_cell, value)
        elif prefix == "header_":
            return None  # header는 스킵
        elif prefix == "data_":
            return ("data", ref_cell, "")
        else:
            # 기타 열
            last_cell = self.table.get_cell(last_row, ref_cell.col)
            if last_cell and last_cell.row < last_row:
                return ("extend", last_cell, None)
            return ("data", ref_cell, "")

    def _get_gstub_action(
        self,
        field_name: str,
        ref_cell: CellInfo,
        gstub_values: Dict[str, str],
        stub_values: Dict[str, str],
    ) -> Tuple[str, CellInfo, Optional[str]]:
        """gstub 열의 액션 결정"""
        # 가장 end_row가 큰 gstub 셀 찾기
        gstub_cell = None
        for (r, c), cell in self.table.cells.items():
            if c == ref_cell.col and cell.field_name == field_name:
                if gstub_cell is None or cell.end_row > gstub_cell.end_row:
                    gstub_cell = cell

        if field_name in gstub_values:
            new_value = gstub_values[field_name]
            if gstub_cell and gstub_cell.text == new_value:
                return ("extend", gstub_cell, None)
            return ("new", ref_cell, new_value)
        elif stub_values:
            return ("new", ref_cell, "")
        return ("data", ref_cell, "")

    def _extend_rowspans(self, col_actions: Dict[int, Tuple[str, CellInfo, Optional[str]]]):
        """col_actions에서 extend 액션인 셀들의 rowspan 확장"""
        extended_cells = set()
        for col, (action, cell, _) in col_actions.items():
            if action == "extend" and cell:
                cell_key = (cell.row, cell.col)
                if cell_key not in extended_cells:
                    self._extend_rowspan(cell)
                    extended_cells.add(cell_key)

    def _create_row_with_actions(self, row_idx: int, col_actions: Dict):
        """액션에 따라 새 행 생성"""
        if self.table is None or self.table.element is None:
            return

        new_tr = self._create_empty_tr()
        if new_tr is None:
            return

        processed_cols = set()

        for col in sorted(col_actions.keys()):
            if col in processed_cols:
                continue

            action, ref_cell, value = col_actions[col]
            colspan = ref_cell.col_span if ref_cell else 1
            field_name = ref_cell.field_name if ref_cell else None

            if action == "extend":
                # rowspan 확장된 열 - 셀 생성 안 함
                pass
            else:
                # new 또는 data - 셀 생성
                self._append_cell_to_tr(
                    new_tr, row_idx, col, value or "", field_name, colspan=colspan
                )

            # colspan 범위 처리됨으로 표시
            for c in range(col, col + colspan):
                processed_cols.add(c)

        self.table.element.append(new_tr)
