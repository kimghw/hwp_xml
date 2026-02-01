# -*- coding: utf-8 -*-
"""
자동 필드명 생성 및 삽입 모듈

테이블 구조를 분석하여 자동으로 필드명을 생성하고 HWPX에 저장합니다.

접두사 규칙:
- header_: 행 전체가 배경색 있음
- add_: 최상단 데이터 행 + 30자 이상 텍스트
- stub_: 텍스트 있음 + 오른쪽 빈 셀 (rowspan=1)
- gstub_: 텍스트 있음 + 오른쪽 빈 셀 (rowspan>1)
- input_: 빈 셀
- data_: 텍스트 있음 + stub 조건 미충족
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, List, Dict, Tuple, Optional, Any
import tempfile
import shutil
import os
import uuid
import logging
from dataclasses import dataclass

from ..table.parser import TableParser, NAMESPACES
from ..table.models import TableInfo


# 로거 설정
_logger = None

def get_logger():
    """로거 반환 (싱글톤)"""
    global _logger
    if _logger is None:
        _logger = logging.getLogger('field_name_generator')
        _logger.setLevel(logging.DEBUG)
        _logger.handlers.clear()
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
        if color.startswith('#'):
            color_hex = color[1:]
            if len(color_hex) == 6:
                try:
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    if r >= 220 and g >= 220 and b >= 220:
                        return False
                except ValueError:
                    pass
        return True


class FieldNameGenerator:
    """테이블 셀에 nc_name 자동 생성"""

    _table_counter = 0

    def __init__(self, text_length_threshold: int = 30, use_random_names: bool = True):
        self.text_length_threshold = text_length_threshold
        self.use_random_names = use_random_names
        self.log = get_logger()

    def generate(self, cells: List[CellForNaming]) -> List[CellForNaming]:
        """셀 목록에 nc_name 자동 생성"""
        if not cells:
            return cells

        FieldNameGenerator._table_counter += 1
        self.log.info(f"\n{'='*80}")
        self.log.info(f"테이블 {FieldNameGenerator._table_counter} 처리 시작")
        self.log.info(f"{'='*80}")

        grid = self._build_grid(cells)
        row_count = max(c.row for c in cells) + 1
        col_count = max(c.col for c in cells) + 1

        self.log.info(f"테이블 크기: {row_count}행 x {col_count}열, 셀 수: {len(cells)}")

        # 4단계 처리
        self.log.info(f"\n[1단계] header_ 식별")
        self._identify_headers(cells, grid, row_count, col_count)

        self.log.info(f"\n[2단계] add_ 식별")
        self._identify_add_cells(cells, grid, row_count, col_count)

        self.log.info(f"\n[3단계] stub_/gstub_ 식별")
        self._identify_stubs(cells, grid, row_count, col_count)

        self.log.info(f"\n[4단계] input_/data_ 식별 및 그룹화")
        self._identify_inputs(cells, grid, row_count, col_count)

        return cells

    def _build_grid(self, cells: List[CellForNaming]) -> Dict[Tuple[int, int], CellForNaming]:
        """셀을 (row, col) -> CellForNaming 그리드로 변환"""
        grid = {}
        for cell in cells:
            for r in range(cell.row, cell.end_row + 1):
                for c in range(cell.col, cell.end_col + 1):
                    grid[(r, c)] = cell
        return grid

    def _get_cell(self, grid: Dict[Tuple[int, int], CellForNaming],
                  row: int, col: int) -> Optional[CellForNaming]:
        return grid.get((row, col))

    def _generate_id(self) -> str:
        return str(uuid.uuid4())[:8]

    def _is_full_row_with_bg(self, grid: Dict[Tuple[int, int], CellForNaming],
                             row: int, col_count: int) -> bool:
        """행 전체가 배경색이 있는지 확인"""
        for col in range(col_count):
            cell = self._get_cell(grid, row, col)
            if not cell:
                return False
            if not cell.has_bg_color and cell.row == row:
                return False
        return True

    def _identify_headers(self, cells: List[CellForNaming],
                          grid: Dict[Tuple[int, int], CellForNaming],
                          row_count: int, col_count: int):
        """header_ 식별"""
        header_rows = set()

        for row in range(row_count):
            if self._is_full_row_with_bg(grid, row, col_count):
                header_rows.add(row)
                self.log.info(f"  행 {row}: 전체 배경색 있음 → header 행")
            else:
                self.log.info(f"  행 {row}: 배경색 없는 셀 존재 → header 중단")
                break

        for cell in cells:
            if cell.row in header_rows and not cell.nc_name:
                cell.nc_name = f"header_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) text='{cell.text[:20]}' → {cell.nc_name}")

    def _identify_add_cells(self, cells: List[CellForNaming],
                            grid: Dict[Tuple[int, int], CellForNaming],
                            row_count: int, col_count: int):
        """add_ 식별"""
        if row_count == 1 and col_count == 1 and len(cells) == 1:
            cell = cells[0]
            if not cell.nc_name and not cell.has_bg_color:
                cell.nc_name = f"add_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [1x1 단일셀]")
                return

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

        for cell in cells:
            if cell.nc_name:
                continue

            if (cell.row == first_data_row and
                len(cell.text.strip()) >= self.text_length_threshold and
                not cell.has_bg_color):
                cell.nc_name = f"add_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [30자 이상]")

    def _identify_stubs(self, cells: List[CellForNaming],
                        grid: Dict[Tuple[int, int], CellForNaming],
                        row_count: int, col_count: int):
        """stub_/gstub_ 식별"""
        for cell in cells:
            if cell.nc_name:
                continue

            if not cell.text.strip():
                continue

            has_empty_right = False
            for right_col in range(cell.end_col + 1, col_count):
                right_cell = self._get_cell(grid, cell.row, right_col)
                if right_cell and right_cell.is_empty:
                    has_empty_right = True
                    break

            if not has_empty_right:
                continue

            if cell.row_span > 1:
                cell.nc_name = f"gstub_{self._generate_id()}"
            else:
                cell.nc_name = f"stub_{self._generate_id()}"
            self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name}")

    def _identify_inputs(self, cells: List[CellForNaming],
                         grid: Dict[Tuple[int, int], CellForNaming],
                         row_count: int, col_count: int):
        """input_/data_ 식별"""
        group_map: Dict[Tuple, str] = {}

        for cell in cells:
            if cell.nc_name:
                continue

            if not cell.is_empty:
                cell.nc_name = f"data_{self._generate_id()}"
                self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name} [data]")
                continue

            # 위 셀에서 헤더 찾기
            header_cell = None
            for check_row in range(cell.row - 1, -1, -1):
                above_cell = self._get_cell(grid, check_row, cell.col)
                if above_cell and above_cell.text.strip():
                    if above_cell.nc_name and (above_cell.nc_name.startswith('stub_') or
                                                above_cell.nc_name.startswith('gstub_')):
                        continue
                    header_cell = above_cell
                    break

            # 좌측 stub 패턴
            has_text_left = False
            stub_pattern = []
            for check_col in range(cell.col - 1, -1, -1):
                left_cell = self._get_cell(grid, cell.row, check_col)
                if left_cell and left_cell.text.strip():
                    if left_cell.nc_name and (left_cell.nc_name.startswith('stub_') or
                                               left_cell.nc_name.startswith('gstub_')):
                        stub_pattern.append(left_cell.nc_name)
                    else:
                        has_text_left = True
                        break

            if header_cell and not has_text_left:
                group_key = (header_cell.text.strip(), cell.col, cell.col_span, tuple(stub_pattern))

                if group_key in group_map:
                    cell.nc_name = group_map[group_key]
                else:
                    nc_name = f"input_{self._generate_id()}"
                    group_map[group_key] = nc_name
                    cell.nc_name = nc_name
            else:
                cell.nc_name = f"input_{self._generate_id()}"

            self.log.info(f"  ({cell.row},{cell.col}) → {cell.nc_name}")


class AutoFieldInserter:
    """자동 필드명을 HWPX 셀에 삽입"""

    def __init__(self, regenerate: bool = False):
        self.regenerate = regenerate
        self.parser = TableParser(auto_field_names=True, regenerate=regenerate)
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def insert_fields(self, hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> List[TableInfo]:
        """HWPX 파일에 자동 필드명 삽입"""
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        tables = self.parser.parse_tables(hwpx_path)
        print(f"테이블 {len(tables)}개 파싱 완료")

        field_mapping = {}
        for table_idx, table in enumerate(tables):
            for (row, col), cell in table.cells.items():
                if cell.field_name:
                    field_mapping[(table_idx, row, col)] = cell.field_name

        temp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            contents_dir = os.path.join(temp_dir, 'Contents')
            section_files = sorted([
                f for f in os.listdir(contents_dir)
                if f.startswith('section') and f.endswith('.xml')
            ])

            table_global_idx = 0
            for section_file in section_files:
                section_path = os.path.join(contents_dir, section_file)
                tree = ET.parse(section_path)
                root = tree.getroot()

                table_global_idx = self._process_section(root, field_mapping, table_global_idx)

                tree.write(section_path, encoding='utf-8', xml_declaration=True)

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        if arcname == 'mimetype':
                            zf.write(file_path, arcname, compress_type=zipfile.ZIP_STORED)
                        else:
                            zf.write(file_path, arcname)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        print(f"필드명 삽입 완료: {output_path}")
        return tables

    def _process_section(self, element, field_mapping: dict, table_idx: int) -> int:
        for child in element:
            if child.tag.endswith('}tbl'):
                table_idx = self._process_table(child, field_mapping, table_idx)
            elif not child.tag.endswith('}tc'):
                table_idx = self._process_section(child, field_mapping, table_idx)
        return table_idx

    def _process_table(self, tbl_elem, field_mapping: dict, table_idx: int) -> int:
        current_table_idx = table_idx
        cell_count = 0
        cell_elements = []

        for tr in tbl_elem:
            if not tr.tag.endswith('}tr'):
                continue

            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue

                cell_elements.append(tc)

                row, col = 0, 0
                for child in tc:
                    if child.tag.endswith('}cellAddr'):
                        row = int(child.get('rowAddr', 0))
                        col = int(child.get('colAddr', 0))
                        break

                key = (current_table_idx, row, col)
                if key in field_mapping:
                    field_name = field_mapping[key]
                    tc.set('name', field_name)
                    cell_count += 1

        if cell_count > 0:
            print(f"  테이블 {current_table_idx}: {cell_count}개 셀에 필드명 설정")

        next_table_idx = current_table_idx + 1
        for tc in cell_elements:
            next_table_idx = self._find_nested_tables(tc, field_mapping, next_table_idx)

        return next_table_idx

    def _find_nested_tables(self, element, field_mapping: dict, table_idx: int) -> int:
        for child in element:
            if child.tag.endswith('}tbl'):
                table_idx = self._process_table(child, field_mapping, table_idx)
            else:
                table_idx = self._find_nested_tables(child, field_mapping, table_idx)
        return table_idx


def insert_auto_fields(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None, regenerate: bool = False):
    """HWPX 파일에 자동 필드명 삽입"""
    inserter = AutoFieldInserter(regenerate=regenerate)
    return inserter.insert_fields(hwpx_path, output_path)


def generate_field_names(cells: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """딕셔너리 형태의 셀 목록에 nc_name 생성"""
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

    generator = FieldNameGenerator()
    generator.generate(cell_objects)

    for i, cell in enumerate(cell_objects):
        cells[i]['nc_name'] = cell.nc_name

    return cells


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법: python -m merge.field.auto_field <input.hwpx> [output.hwpx] [--regenerate]")
        sys.exit(1)

    regenerate = '--regenerate' in sys.argv
    args = [a for a in sys.argv[1:] if a != '--regenerate']

    input_path = args[0]
    output_path = args[1] if len(args) > 1 else None

    tables = insert_auto_fields(input_path, output_path, regenerate=regenerate)

    print(f"\n처리 완료: {len(tables)}개 테이블")
    for i, table in enumerate(tables):
        print(f"  테이블 {i}: {table.row_count}x{table.col_count}, 필드 {len(table.field_to_cell)}개")
