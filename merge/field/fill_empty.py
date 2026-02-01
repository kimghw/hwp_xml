# -*- coding: utf-8 -*-
"""
빈 셀 필드명 채우기 모듈

행 추가 후 필드명이 없는 셀에 위 셀의 필드명을 복사합니다.

조건:
1. 행의 모든 셀이 필드명(nc_name)이 없음
2. 행의 모든 셀 너비가 바로 위 셀과 동일
3. 위 조건 충족 시 바로 위 셀의 필드명을 복사
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, List, Dict, Tuple, Optional
import tempfile
import shutil
import os

from ..table.parser import NAMESPACES


class EmptyFieldFiller:
    """빈 셀에 위 셀 필드명 복사"""

    def __init__(self):
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def fill_fields(self, hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Dict[str, int]:
        """
        HWPX 파일에서 필드명 없는 셀에 위 셀 필드명 복사

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_path: 출력 HWPX 파일 경로 (None이면 원본 덮어쓰기)

        Returns:
            처리 결과 {'tables': 테이블 수, 'rows_filled': 채워진 행 수, 'cells_filled': 채워진 셀 수}
        """
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        result = {'tables': 0, 'rows_filled': 0, 'cells_filled': 0}

        temp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            contents_dir = os.path.join(temp_dir, 'Contents')
            section_files = sorted([
                f for f in os.listdir(contents_dir)
                if f.startswith('section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                section_path = os.path.join(contents_dir, section_file)
                tree = ET.parse(section_path)
                root = tree.getroot()

                section_result = self._process_section(root)
                result['tables'] += section_result['tables']
                result['rows_filled'] += section_result['rows_filled']
                result['cells_filled'] += section_result['cells_filled']

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

        print(f"빈 필드 채우기 완료: {output_path}")
        print(f"  테이블 {result['tables']}개, {result['rows_filled']}행 {result['cells_filled']}셀 채움")
        return result

    def _process_section(self, element) -> Dict[str, int]:
        """section XML에서 테이블 처리"""
        result = {'tables': 0, 'rows_filled': 0, 'cells_filled': 0}

        for child in element:
            if child.tag.endswith('}tbl'):
                table_result = self._process_table(child)
                result['tables'] += 1
                result['rows_filled'] += table_result['rows_filled']
                result['cells_filled'] += table_result['cells_filled']
            elif not child.tag.endswith('}tc'):
                sub_result = self._process_section(child)
                result['tables'] += sub_result['tables']
                result['rows_filled'] += sub_result['rows_filled']
                result['cells_filled'] += sub_result['cells_filled']

        return result

    def _process_table(self, tbl_elem) -> Dict[str, int]:
        """테이블의 빈 행에 위 셀 필드명 복사"""
        result = {'rows_filled': 0, 'cells_filled': 0}

        # 행 목록 수집
        rows = []
        for child in tbl_elem:
            if child.tag.endswith('}tr'):
                rows.append(child)

        if len(rows) < 2:
            return result

        # 각 행의 셀 정보 추출
        row_cells = []
        for tr in rows:
            cells = self._extract_row_cells(tr)
            row_cells.append(cells)

        # 두 번째 행부터 처리
        for row_idx in range(1, len(rows)):
            current_cells = row_cells[row_idx]
            above_cells = row_cells[row_idx - 1]

            # 조건 확인:
            # 1. 현재 행의 모든 셀에 필드명 없음
            # 2. 셀 개수가 동일
            # 3. 각 셀의 너비(colspan)가 동일
            if not current_cells or not above_cells:
                continue

            all_empty = all(not cell.get('name') for cell in current_cells)
            same_count = len(current_cells) == len(above_cells)
            same_widths = same_count and all(
                current_cells[i].get('colspan', 1) == above_cells[i].get('colspan', 1)
                for i in range(len(current_cells))
            )

            if all_empty and same_widths:
                # 위 셀의 필드명 복사
                filled = 0
                for i, cell in enumerate(current_cells):
                    above_name = above_cells[i].get('name')
                    if above_name:
                        cell['element'].set('name', above_name)
                        filled += 1

                if filled > 0:
                    result['rows_filled'] += 1
                    result['cells_filled'] += filled

        # 중첩 테이블도 처리
        for tr in rows:
            for tc in tr:
                if tc.tag.endswith('}tc'):
                    nested_result = self._find_nested_tables(tc)
                    result['rows_filled'] += nested_result['rows_filled']
                    result['cells_filled'] += nested_result['cells_filled']

        return result

    def _extract_row_cells(self, tr_elem) -> List[Dict]:
        """행에서 셀 정보 추출"""
        cells = []
        for tc in tr_elem:
            if not tc.tag.endswith('}tc'):
                continue

            cell_info = {
                'element': tc,
                'name': tc.get('name'),
                'colspan': 1,
                'rowspan': 1,
            }

            # cellSpan에서 colspan, rowspan 추출
            for child in tc:
                if child.tag.endswith('}cellSpan'):
                    cell_info['colspan'] = int(child.get('colSpan', 1))
                    cell_info['rowspan'] = int(child.get('rowSpan', 1))
                    break

            cells.append(cell_info)

        return cells

    def _find_nested_tables(self, element) -> Dict[str, int]:
        """중첩 테이블 찾아 처리"""
        result = {'rows_filled': 0, 'cells_filled': 0}

        for child in element:
            if child.tag.endswith('}tbl'):
                table_result = self._process_table(child)
                result['rows_filled'] += table_result['rows_filled']
                result['cells_filled'] += table_result['cells_filled']
            else:
                nested_result = self._find_nested_tables(child)
                result['rows_filled'] += nested_result['rows_filled']
                result['cells_filled'] += nested_result['cells_filled']

        return result


def fill_empty_fields(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Dict[str, int]:
    """
    HWPX 파일에서 필드명 없는 셀에 위 셀 필드명 복사

    Args:
        hwpx_path: 입력 HWPX 파일 경로
        output_path: 출력 HWPX 파일 경로 (None이면 원본 덮어쓰기)

    Returns:
        처리 결과 {'tables': 테이블 수, 'rows_filled': 채워진 행 수, 'cells_filled': 채워진 셀 수}
    """
    filler = EmptyFieldFiller()
    return filler.fill_fields(hwpx_path, output_path)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법: python -m merge.field.fill_empty <input.hwpx> [output.hwpx]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    result = fill_empty_fields(input_path, output_path)
    print(f"\n처리 완료: {result}")
