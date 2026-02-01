# -*- coding: utf-8 -*-
"""
HWPX 테이블 파싱 모듈

개요:
- TableParser: HWPX 파일에서 테이블 파싱
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Union
from pathlib import Path
from io import BytesIO

from .models import CellInfo, TableInfo


# XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class TableParser:
    """HWPX 테이블 파싱"""

    def __init__(self, auto_field_names: bool = False, regenerate: bool = False):
        """
        Args:
            auto_field_names: True면 필드명이 없는 셀에 자동으로 nc_name 생성
            regenerate: True면 기존 필드명 무시하고 새로 생성
        """
        self._ns = ""  # 네임스페이스 접두사
        self._auto_field_names = auto_field_names
        self._regenerate = regenerate
        self._border_fills: Dict[str, Dict] = {}  # borderFillIDRef -> 배경색 등

    def parse_tables(self, hwpx_path: Union[str, Path]) -> List[TableInfo]:
        """HWPX 파일에서 모든 테이블 파싱"""
        hwpx_path = Path(hwpx_path)
        tables = []
        self._border_fills.clear()

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # header.xml에서 borderFill 정보 로드
            if 'Contents/header.xml' in zf.namelist():
                header_content = zf.read('Contents/header.xml')
                self._parse_header(header_content)

            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                content = zf.read(section_file)
                section_tables = self._parse_section(content)
                tables.extend(section_tables)

        return tables

    def _parse_header(self, xml_content: bytes):
        """header.xml에서 borderFill 정보 파싱"""
        root = ET.parse(BytesIO(xml_content)).getroot()

        for elem in root.iter():
            if elem.tag.endswith('}borderFill'):
                bf_id = elem.get('id', '')
                if bf_id:
                    self._border_fills[bf_id] = {'bg_color': ''}
                    for child in elem:
                        if child.tag.endswith('}fillBrush'):
                            for brush_child in child:
                                if brush_child.tag.endswith('}winBrush'):
                                    self._border_fills[bf_id]['bg_color'] = brush_child.get('faceColor', '')

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

        # 자동 필드명 생성
        if self._auto_field_names:
            self._generate_field_names(table)

        # 필드명 매핑 생성
        for (row, col), cell in table.cells.items():
            if cell.field_name:
                table.field_to_cell[cell.field_name] = (row, col)

        return table

    def _generate_field_names(self, table: TableInfo):
        """필드명이 없는 셀에 자동으로 nc_name 생성"""
        # 지연 import (순환 참조 방지)
        from merge.field.auto_insert_field_template import FieldNameGenerator, CellForNaming

        # CellInfo -> CellForNaming 변환 (모든 셀 포함, 그룹화를 위해)
        cells_for_naming = []
        cell_mapping = {}  # CellForNaming -> CellInfo
        cells_with_existing_name = set()  # 이미 필드명이 있는 셀

        for (row, col), cell in table.cells.items():
            naming_cell = CellForNaming(
                row=cell.row,
                col=cell.col,
                row_span=cell.row_span,
                col_span=cell.col_span,
                end_row=cell.end_row,
                end_col=cell.end_col,
                text=cell.text,
                bg_color=cell.bg_color,
                nc_name=cell.field_name,  # 기존 필드명 유지
            )
            cells_for_naming.append(naming_cell)
            cell_mapping[id(naming_cell)] = cell

            if cell.field_name:
                cells_with_existing_name.add(id(naming_cell))

        if not cells_for_naming:
            return

        # 자동 필드명 생성
        generator = FieldNameGenerator()
        generator.generate(cells_for_naming)

        # 결과를 원본 CellInfo에 반영 (기존 필드명이 없는 셀만)
        for naming_cell in cells_for_naming:
            if id(naming_cell) in cells_with_existing_name:
                continue  # 기존 필드명 유지
            original_cell = cell_mapping[id(naming_cell)]
            if naming_cell.nc_name:
                original_cell.field_name = naming_cell.nc_name

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

            # 배경색 추출 (borderFillIDRef)
            border_fill_id = tc_elem.get('borderFillIDRef', '')
            if border_fill_id and border_fill_id in self._border_fills:
                cell.bg_color = self._border_fills[border_fill_id].get('bg_color', '')

            # tc 태그의 name 속성에서 필드명 추출 (regenerate가 아닐 때만)
            if not self._regenerate:
                tc_name = tc_elem.get('name', '')
                if tc_name:
                    cell.field_name = tc_name

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

                elif tag == 'cellSz':
                    # 셀 크기 파싱
                    cell.width = int(child.get('width', 0))
                    cell.height = int(child.get('height', 0))

                elif tag == 'subList':
                    # 텍스트 추출
                    text = self._extract_text(child)
                    cell.text = text
                    cell.is_empty = not text.strip()

                    # 필드명 추출 (regenerate가 아닐 때만)
                    if not self._regenerate:
                        field_name = self._extract_field_name(child)
                        if field_name:
                            cell.field_name = field_name

            cell.end_row = cell.row + cell.row_span - 1
            cell.end_col = cell.col + cell.col_span - 1

            # 열 너비/행 높이 기록 (colspan/rowspan이 1인 셀만)
            if cell.col_span == 1 and cell.width > 0:
                if cell.col not in table.col_widths:
                    table.col_widths[cell.col] = cell.width
            if cell.row_span == 1 and cell.height > 0:
                if cell.row not in table.row_heights:
                    table.row_heights[cell.row] = cell.height

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
