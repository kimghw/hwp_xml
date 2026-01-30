# -*- coding: utf-8 -*-
"""
nc_name 자동 생성 모듈

테이블 셀의 위치, 텍스트, 배경색, rowspan 등을 분석하여
자동으로 필드명(nc_name)을 설정합니다.

접두사 규칙:
- header_: 최상단 행 + 배경색 있음, 또는 헤더와 연결되고 배경색 동일
- add_: 최상단 행 + 텍스트 30자 이상 + 배경색 없음
- stub_: 텍스트 있음 + 오른쪽에 빈 셀 존재 + rowspan 없음
- gstub_: 텍스트 있음 + 오른쪽에 빈 셀 존재 + rowspan 있음
- input_: 빈 셀 (기본값)

동일 필드명 그룹화:
- 빈 셀(input_)의 경우, 위로 이동 시 첫 번째 만나는 셀의 값이 동일하고
- 좌측으로 이동 시 텍스트 있는 셀이 없으면
- 같은 nc_name 사용
"""

import uuid
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field

# 로거 설정
_logger = None

def get_logger():
    """로거 반환 (싱글톤)"""
    global _logger
    if _logger is None:
        _logger = logging.getLogger('field_name_generator')
        _logger.setLevel(logging.DEBUG)
        # 기존 핸들러 제거
        _logger.handlers.clear()
        # 파일 핸들러 추가
        log_path = Path(__file__).parent / 'field_name_log.txt'
        fh = logging.FileHandler(log_path, mode='w', encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(message)s')
        fh.setFormatter(formatter)
        _logger.addHandler(fh)
    return _logger


@dataclass
class CellForNaming:
    """필드명 생성을 위한 셀 정보"""
    row: int = 0
    col: int = 0
    row_span: int = 1
    col_span: int = 1
    end_row: int = 0
    end_col: int = 0
    text: str = ""
    bg_color: str = ""
    nc_name: str = ""

    @property
    def is_empty(self) -> bool:
        return not self.text.strip()

    @property
    def has_bg_color(self) -> bool:
        """배경색이 있는지 확인 (흰색 계열 제외)"""
        if not self.bg_color:
            return False
        color = self.bg_color.lower().strip()
        if color in ('', 'none'):
            return False
        # 흰색 계열 제외 (#ffffff, #fffff*, #ffffe* 등)
        if color.startswith('#'):
            color_hex = color[1:]
            if len(color_hex) == 6:
                # RGB 값 추출
                try:
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    # 흰색에 가까운 색상 제외 (R, G, B 모두 220 이상)
                    if r >= 220 and g >= 220 and b >= 220:
                        return False
                except ValueError:
                    pass
        return True


class FieldNameGenerator:
    """
    테이블 셀에 nc_name 자동 생성

    사용 예:
        generator = FieldNameGenerator()
        cells = generator.generate(table_cells)
    """

    _table_counter = 0  # 테이블 카운터

    def __init__(self, text_length_threshold: int = 30, use_random_names: bool = True):
        """
        Args:
            text_length_threshold: add_ 접두사 적용을 위한 최소 텍스트 길이
            use_random_names: True면 텍스트 기반 대신 랜덤 ID 사용
        """
        self.text_length_threshold = text_length_threshold
        self.use_random_names = use_random_names
        self.log = get_logger()

    def generate(self, cells: List[CellForNaming]) -> List[CellForNaming]:
        """
        셀 목록에 nc_name 자동 생성

        Args:
            cells: CellForNaming 목록

        Returns:
            nc_name이 설정된 셀 목록
        """
        if not cells:
            return cells

        FieldNameGenerator._table_counter += 1
        self.log.info(f"\n{'='*80}")
        self.log.info(f"테이블 {FieldNameGenerator._table_counter} 처리 시작")
        self.log.info(f"{'='*80}")

        # 셀을 그리드로 변환
        grid = self._build_grid(cells)
        row_count = max(c.row for c in cells) + 1
        col_count = max(c.col for c in cells) + 1

        self.log.info(f"테이블 크기: {row_count}행 x {col_count}열, 셀 수: {len(cells)}")

        # 1단계: header_ 식별
        self.log.info(f"\n[1단계] header_ 식별")
        self._identify_headers(cells, grid, row_count, col_count)

        # 2단계: add_ 식별 (최상단 + 30자 이상 + 배경색 없음)
        self.log.info(f"\n[2단계] add_ 식별")
        self._identify_add_cells(cells, grid, row_count, col_count)

        # 3단계: stub_/gstub_ 식별
        self.log.info(f"\n[3단계] stub_/gstub_ 식별")
        self._identify_stubs(cells, grid, row_count, col_count)

        # 4단계: input_ 식별 및 그룹화
        self.log.info(f"\n[4단계] input_/data_ 식별 및 그룹화")
        self._identify_inputs(cells, grid, row_count, col_count)

        return cells

    def _build_grid(self, cells: List[CellForNaming]) -> Dict[Tuple[int, int], CellForNaming]:
        """셀을 (row, col) -> CellForNaming 그리드로 변환"""
        grid = {}
        for cell in cells:
            # 셀이 차지하는 모든 위치에 매핑
            for r in range(cell.row, cell.end_row + 1):
                for c in range(cell.col, cell.end_col + 1):
                    grid[(r, c)] = cell
        return grid

    def _get_cell(self, grid: Dict[Tuple[int, int], CellForNaming],
                  row: int, col: int) -> Optional[CellForNaming]:
        """그리드에서 셀 가져오기"""
        return grid.get((row, col))

    def _generate_id(self) -> str:
        """랜덤 ID 생성"""
        return str(uuid.uuid4())[:8]

    def _is_full_row_with_bg(self, grid: Dict[Tuple[int, int], CellForNaming],
                             row: int, col_count: int) -> bool:
        """행 전체가 배경색이 있는지 확인"""
        for col in range(col_count):
            cell = self._get_cell(grid, row, col)
            if not cell:
                self.log.debug(f"    col {col}: 셀 없음 → False")
                return False
            # rowspan으로 위 셀과 연결된 경우도 허용
            if not cell.has_bg_color and cell.row == row:
                self.log.debug(f"    col {col}: bg_color='{cell.bg_color}', has_bg={cell.has_bg_color}, cell.row={cell.row} → False")
                return False
            self.log.debug(f"    col {col}: bg_color='{cell.bg_color}', has_bg={cell.has_bg_color}, cell.row={cell.row} → OK")
        return True

    def _identify_headers(self, cells: List[CellForNaming],
                          grid: Dict[Tuple[int, int], CellForNaming],
                          row_count: int, col_count: int):
        """header_ 식별: 행 전체가 배경색 있음, 연속된 행도 모두 배경색 있으면 header"""

        header_rows = set()

        # 첫 번째 행부터 확인: 행 전체가 배경색이 있어야 함
        for row in range(row_count):
            if self._is_full_row_with_bg(grid, row, col_count):
                header_rows.add(row)
                self.log.info(f"  행 {row}: 전체 배경색 있음 → header 행")
            else:
                self.log.info(f"  행 {row}: 배경색 없는 셀 존재 → header 중단")
                break  # 연속되지 않으면 중단

        # 헤더 셀에 nc_name 설정 (헤더 행에 있는 모든 셀)
        for cell in cells:
            if cell.row in header_rows and not cell.nc_name:
                cell.nc_name = f"header_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:20]}' → {cell.nc_name} [조건: 행 전체 배경색]")

    def _identify_add_cells(self, cells: List[CellForNaming],
                            grid: Dict[Tuple[int, int], CellForNaming],
                            row_count: int, col_count: int):
        """add_ 식별: (최상단 + 텍스트 30자 이상 + 배경색 없음) 또는 (1x1 단일 셀 + 배경색 없음)"""

        # 1x1 단일 셀 테이블이고 배경색 없으면 add_
        if row_count == 1 and col_count == 1 and len(cells) == 1:
            cell = cells[0]
            if not cell.nc_name and not cell.has_bg_color:
                cell.nc_name = f"add_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:30]}' → {cell.nc_name} [조건: 1x1단일셀+배경색없음]")
                return

        # 헤더가 아닌 첫 번째 행 찾기
        first_data_row = 0
        for row in range(row_count):
            has_header = False
            for col in range(col_count):
                cell = self._get_cell(grid, row, col)
                if cell and cell.nc_name and cell.nc_name.startswith('header_'):
                    has_header = True
                    break
            if not has_header:
                first_data_row = row
                break

        self.log.info(f"  첫 번째 데이터 행: {first_data_row}")

        for cell in cells:
            if cell.nc_name:  # 이미 설정됨
                continue

            # 최상단 데이터 행에 있고, 텍스트 30자 이상, 배경색 없음
            if (cell.row == first_data_row and
                len(cell.text.strip()) >= self.text_length_threshold and
                not cell.has_bg_color):
                cell.nc_name = f"add_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:30]}...' → {cell.nc_name} [조건: 최상단+30자이상+배경색없음]")

    def _identify_stubs(self, cells: List[CellForNaming],
                        grid: Dict[Tuple[int, int], CellForNaming],
                        row_count: int, col_count: int):
        """stub_/gstub_ 식별: 텍스트 있음 + 오른쪽에 빈 셀 존재"""

        for cell in cells:
            if cell.nc_name:  # 이미 설정됨
                continue

            if not cell.text.strip():  # 텍스트 없음
                continue

            # 오른쪽에 빈 셀이 있는지 확인
            has_empty_right = False
            empty_right_col = -1
            for right_col in range(cell.end_col + 1, col_count):
                right_cell = self._get_cell(grid, cell.row, right_col)
                if right_cell and right_cell.is_empty:
                    has_empty_right = True
                    empty_right_col = right_col
                    break

            if not has_empty_right:
                continue

            if cell.row_span > 1:
                # gstub_: rowspan 있음 (병합된 셀)
                cell.nc_name = f"gstub_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:20]}' → {cell.nc_name} [조건: 텍스트+오른쪽빈셀(col{empty_right_col})+rowspan={cell.row_span}]")
            else:
                # stub_: rowspan 없음
                cell.nc_name = f"stub_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:20]}' → {cell.nc_name} [조건: 텍스트+오른쪽빈셀(col{empty_right_col})+rowspan=1]")

    def _identify_inputs(self, cells: List[CellForNaming],
                         grid: Dict[Tuple[int, int], CellForNaming],
                         row_count: int, col_count: int):
        """input_ 식별: 빈 셀, 동일 조건이면 같은 이름"""

        # 그룹화를 위한 매핑: (header_text, col, col_span, stub_pattern) -> nc_name
        group_map: Dict[Tuple, str] = {}

        for cell in cells:
            if cell.nc_name:  # 이미 설정됨
                continue

            if not cell.is_empty:  # 빈 셀만 처리
                # 빈 셀이 아니고 아직 nc_name이 없으면 기본 이름 설정
                cell.nc_name = f"data_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:20]}' → {cell.nc_name} [조건: 텍스트있음+stub조건미충족]")
                continue

            # 위로 이동하여 첫 번째 텍스트 있는 셀 찾기 (stub/gstub 제외)
            # stub/gstub는 행 헤더이므로 열 헤더로 사용할 수 없음
            header_cell = None
            skipped_stubs = []
            for check_row in range(cell.row - 1, -1, -1):
                above_cell = self._get_cell(grid, check_row, cell.col)
                if above_cell and above_cell.text.strip():
                    # stub/gstub는 헤더로 사용하지 않음
                    if above_cell.nc_name and (above_cell.nc_name.startswith('stub_') or
                                                above_cell.nc_name.startswith('gstub_')):
                        skipped_stubs.append(f"({check_row},{cell.col})={above_cell.nc_name}")
                        continue  # stub/gstub는 건너뛰고 계속 위로 탐색
                    header_cell = above_cell
                    break

            # 좌측 stub/gstub 패턴 수집 + 일반 텍스트 셀 확인
            # stub/gstub가 다르면 다른 그룹이 됨
            has_text_left = False
            left_text_cell = None
            stub_pattern = []  # 좌측 stub/gstub의 nc_name 목록
            for check_col in range(cell.col - 1, -1, -1):
                left_cell = self._get_cell(grid, cell.row, check_col)
                if left_cell and left_cell.text.strip():
                    if left_cell.nc_name and (left_cell.nc_name.startswith('stub_') or
                                               left_cell.nc_name.startswith('gstub_')):
                        # stub/gstub는 패턴에 추가
                        stub_pattern.append(left_cell.nc_name)
                    else:
                        # 일반 텍스트 셀 발견
                        has_text_left = True
                        left_text_cell = left_cell
                        break

            # 그룹화 조건: 위 셀 텍스트 동일 + 좌측에 텍스트 없음 + colspan 동일 + stub 패턴 동일
            if header_cell and not has_text_left:
                # group_key에 stub_pattern 포함 (stub/gstub가 다르면 다른 그룹)
                group_key = (header_cell.text.strip(), cell.col, cell.col_span, tuple(stub_pattern))

                if group_key in group_map:
                    cell.nc_name = group_map[group_key]
                    self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [그룹화: 헤더='{header_cell.text[:15]}', col={cell.col}, colspan={cell.col_span}, stub패턴={stub_pattern}]")
                else:
                    # 새 그룹 생성 - 랜덤 ID 사용
                    nc_name = f"input_{self._generate_id()}"
                    group_map[group_key] = nc_name
                    cell.nc_name = nc_name
                    self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [새그룹: 헤더='{header_cell.text[:15]}', col={cell.col}, colspan={cell.col_span}, stub패턴={stub_pattern}]")
            else:
                # 그룹화 조건 미충족 - 고유 이름 생성
                cell.nc_name = f"input_{self._generate_id()}"
                reason = []
                if not header_cell:
                    reason.append("헤더없음")
                    if skipped_stubs:
                        reason.append(f"건너뛴stub={skipped_stubs}")
                if has_text_left:
                    reason.append(f"좌측텍스트='{left_text_cell.text[:10] if left_text_cell else ''}'")
                self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [그룹화불가: {', '.join(reason)}]")

    def _sanitize_name(self, text: str, max_length: int = 20) -> str:
        """텍스트를 필드명으로 변환 (특수문자 제거, 길이 제한)"""
        if not text:
            return ""

        # 공백과 특수문자를 언더스코어로 변환
        result = []
        for char in text.strip()[:max_length]:
            if char.isalnum() or char in ('_', '-'):
                result.append(char)
            elif char in (' ', '\t', '\n'):
                if result and result[-1] != '_':
                    result.append('_')

        name = ''.join(result).strip('_')
        return name if name else "unnamed"


def generate_field_names(cells: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    딕셔너리 형태의 셀 목록에 nc_name 생성

    Args:
        cells: 셀 정보 딕셔너리 목록
            필수 키: row, col, text
            선택 키: row_span, col_span, bg_color

    Returns:
        nc_name이 추가된 셀 목록
    """
    # 딕셔너리를 CellForNaming으로 변환
    cell_objects = []
    for c in cells:
        cell = CellForNaming(
            row=c.get('row', 0),
            col=c.get('col', 0),
            row_span=c.get('row_span', 1),
            col_span=c.get('col_span', 1),
            end_row=c.get('row', 0) + c.get('row_span', 1) - 1,
            end_col=c.get('col', 0) + c.get('col_span', 1) - 1,
            text=c.get('text', ''),
            bg_color=c.get('bg_color', ''),
        )
        cell_objects.append(cell)

    # 필드명 생성
    generator = FieldNameGenerator()
    generator.generate(cell_objects)

    # 결과를 원본 딕셔너리에 반영
    for i, cell in enumerate(cell_objects):
        cells[i]['nc_name'] = cell.nc_name

    return cells


if __name__ == "__main__":
    # 테스트
    test_cells = [
        # 헤더 행 (배경색 있음)
        {'row': 0, 'col': 0, 'text': '구분', 'bg_color': '#CCCCCC', 'row_span': 1, 'col_span': 1},
        {'row': 0, 'col': 1, 'text': '항목1', 'bg_color': '#CCCCCC', 'row_span': 1, 'col_span': 1},
        {'row': 0, 'col': 2, 'text': '항목2', 'bg_color': '#CCCCCC', 'row_span': 1, 'col_span': 1},
        # 데이터 행 1
        {'row': 1, 'col': 0, 'text': '분류A', 'bg_color': '', 'row_span': 2, 'col_span': 1},  # gstub
        {'row': 1, 'col': 1, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input
        {'row': 1, 'col': 2, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input
        # 데이터 행 2
        {'row': 2, 'col': 1, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input (같은 헤더)
        {'row': 2, 'col': 2, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input (같은 헤더)
        # 데이터 행 3
        {'row': 3, 'col': 0, 'text': '분류B', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # stub
        {'row': 3, 'col': 1, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input
        {'row': 3, 'col': 2, 'text': '', 'bg_color': '', 'row_span': 1, 'col_span': 1},  # input
    ]

    result = generate_field_names(test_cells)

    print("필드명 생성 결과:")
    print("-" * 60)
    for cell in result:
        print(f"({cell['row']}, {cell['col']}) text='{cell['text'][:10]}...' -> {cell['nc_name']}")
