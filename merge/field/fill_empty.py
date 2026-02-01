# -*- coding: utf-8 -*-
"""
빈 셀 필드명 채우기 모듈

행 추가 후 필드명이 없는 셀에 필드명을 자동으로 설정합니다.

Case 1: 위 행 복사
- 필드명이 없고, 바로 위 행이 모두 data_, input_ 인 경우
- 컬럼 수와 너비가 동일한 경우
- 위 셀의 nc_name을 복사

Case 2: gstub 처리
- 필드명이 없고, rowspan > 1 (위아래 병합)인 경우
- gstub_{랜덤}으로 설정
- 우측 빈 셀: 첫 번째 행은 input_{랜덤}, 두 번째 행부터는 첫 번째 행의 nc_name 복사

Case 3: 연속 gstub 처리
- gstub가 2개 이상 연속된 경우
- 모든 gstub가 동일한 행에 걸쳐있으면 빈 셀도 동일한 nc_name
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, List, Dict, Tuple, Optional
import tempfile
import shutil
import os
import uuid

from ..table.parser import NAMESPACES


def _generate_id() -> str:
    """랜덤 ID 생성"""
    return str(uuid.uuid4())[:8]


class EmptyFieldFiller:
    """빈 셀에 필드명 자동 채우기"""

    def __init__(self):
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def fill_fields(self, hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Dict[str, int]:
        """
        HWPX 파일에서 필드명 없는 셀에 필드명 자동 설정

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
        """테이블의 빈 셀에 필드명 설정"""
        result = {'rows_filled': 0, 'cells_filled': 0}

        # 행 목록 수집
        rows = []
        for child in tbl_elem:
            if child.tag.endswith('}tr'):
                rows.append(child)

        if len(rows) < 1:
            return result

        # 각 행의 셀 정보 추출
        row_cells = []
        for tr in rows:
            cells = self._extract_row_cells(tr)
            row_cells.append(cells)

        # 그리드 형태로 셀 매핑 (rowspan 고려)
        grid = self._build_grid(row_cells)

        # Case 2 & 3: gstub 처리 (rowspan > 1인 셀)
        gstub_result = self._process_gstub_cells(row_cells, grid)
        result['cells_filled'] += gstub_result['cells_filled']

        # Case 1: 위 행 복사 (data_, input_ 패턴)
        copy_result = self._process_copy_above(row_cells, grid)
        result['rows_filled'] += copy_result['rows_filled']
        result['cells_filled'] += copy_result['cells_filled']

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
                'col': len(cells),
                'text': '',
            }

            # cellSpan, cellAddr, text 추출
            for child in tc:
                if child.tag.endswith('}cellSpan'):
                    cell_info['colspan'] = int(child.get('colSpan', 1))
                    cell_info['rowspan'] = int(child.get('rowSpan', 1))
                elif child.tag.endswith('}cellAddr'):
                    cell_info['col'] = int(child.get('colAddr', cell_info['col']))
                elif child.tag.endswith('}subList'):
                    cell_info['text'] = self._extract_text(child)

            cells.append(cell_info)

        return cells

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

    def _build_grid(self, row_cells: List[List[Dict]]) -> Dict[Tuple[int, int], Dict]:
        """(row, col) -> cell 그리드 생성 (rowspan/colspan 고려)"""
        grid = {}
        for row_idx, cells in enumerate(row_cells):
            for cell in cells:
                col = cell['col']
                for r in range(row_idx, row_idx + cell['rowspan']):
                    for c in range(col, col + cell['colspan']):
                        grid[(r, c)] = cell
        return grid

    def _process_gstub_cells(self, row_cells: List[List[Dict]], grid: Dict) -> Dict[str, int]:
        """
        Case 2 & 3: gstub 처리
        - rowspan > 1이고 필드명 없는 셀 → gstub_{랜덤}
        - 우측 빈 셀: 첫 번째 행은 input_{랜덤}, 이후 행은 첫 번째 행 복사
        - 2개 이상 gstub가 있을 때: 모든 gstub의 rowspan이 동일해야만 같은 nc_name 공유
        """
        result = {'cells_filled': 0}

        if not row_cells:
            return result

        col_count = max(
            (cell['col'] + cell['colspan'] for cells in row_cells for cell in cells),
            default=0
        )

        for row_idx, cells in enumerate(row_cells):
            for cell in cells:
                # rowspan > 1이고 필드명 없는 셀 → gstub
                if cell['rowspan'] > 1 and not cell['name'] and cell['text'].strip():
                    cell['name'] = f"gstub_{_generate_id()}"
                    cell['element'].set('name', cell['name'])
                    result['cells_filled'] += 1

        # 각 행에서 gstub 우측의 빈 셀 처리
        for row_idx, cells in enumerate(row_cells):
            # 이 행에 영향을 주는 gstub 목록 수집 (현재 행이 gstub 범위 내인 것들)
            affecting_gstubs = []
            for r_idx, r_cells in enumerate(row_cells):
                for cell in r_cells:
                    if cell['name'] and cell['name'].startswith('gstub_'):
                        # 이 gstub가 현재 row_idx에 영향을 주는지 확인
                        if r_idx <= row_idx < r_idx + cell['rowspan']:
                            affecting_gstubs.append({
                                'cell': cell,
                                'start_row': r_idx,
                                'end_col': cell['col'] + cell['colspan']
                            })

            if not affecting_gstubs:
                continue

            # gstub 우측의 빈 셀들 처리
            # 가장 오른쪽 gstub의 끝 컬럼부터 처리
            rightmost_gstub_end = max(g['end_col'] for g in affecting_gstubs)

            for cell in cells:
                if cell['col'] < rightmost_gstub_end:
                    continue
                if cell['name']:
                    continue
                if cell['text'].strip():
                    continue

                # 빈 셀 발견
                # 2개 이상의 gstub가 있을 때, 모든 gstub의 rowspan이 동일해야만 같은 nc_name 공유
                if len(affecting_gstubs) >= 2:
                    rowspans = [g['cell']['rowspan'] for g in affecting_gstubs]
                    all_same_rowspan = len(set(rowspans)) == 1

                    if not all_same_rowspan:
                        # gstub들의 rowspan이 다르면 각 셀에 새 nc_name 부여
                        new_name = f"input_{_generate_id()}"
                        cell['name'] = new_name
                        cell['element'].set('name', new_name)
                        result['cells_filled'] += 1
                        continue

                # Case 3: 연속 gstub 패턴 확인 (모든 gstub가 동일 행 범위인지)
                gstub_pattern = tuple(sorted(
                    (g['cell']['name'], g['start_row'], g['cell']['rowspan'])
                    for g in affecting_gstubs
                ))

                # 첫 번째 행인지 확인 (gstub 시작 행)
                is_first_row = row_idx == min(g['start_row'] for g in affecting_gstubs)

                if is_first_row:
                    # 첫 번째 행: input_{랜덤} 생성
                    new_name = f"input_{_generate_id()}"
                    cell['name'] = new_name
                    cell['element'].set('name', new_name)
                    # 패턴 저장 (같은 gstub 패턴 + 컬럼 위치)
                    cell['_gstub_pattern'] = (gstub_pattern, cell['col'])
                    result['cells_filled'] += 1
                else:
                    # 두 번째 행 이후: 첫 번째 행의 같은 컬럼 셀 찾아서 복사
                    first_row_idx = min(g['start_row'] for g in affecting_gstubs)
                    first_row_cell = grid.get((first_row_idx, cell['col']))
                    if first_row_cell and first_row_cell.get('name'):
                        cell['name'] = first_row_cell['name']
                        cell['element'].set('name', first_row_cell['name'])
                        result['cells_filled'] += 1

        return result

    def _process_copy_above(self, row_cells: List[List[Dict]], grid: Dict) -> Dict[str, int]:
        """
        Case 1: 위 행 복사
        - 필드명 없는 행 전체가 빈 경우
        - 위 행이 모두 data_ 또는 input_인 경우
        - 컬럼 수와 너비가 동일한 경우
        """
        result = {'rows_filled': 0, 'cells_filled': 0}

        for row_idx in range(1, len(row_cells)):
            current_cells = row_cells[row_idx]
            above_cells = row_cells[row_idx - 1]

            if not current_cells or not above_cells:
                continue

            # 조건 1: 현재 행의 모든 셀에 필드명 없음
            all_empty = all(not cell.get('name') for cell in current_cells)
            if not all_empty:
                continue

            # 조건 2: 셀 개수가 동일
            if len(current_cells) != len(above_cells):
                continue

            # 조건 3: 각 셀의 너비(colspan)가 동일
            same_widths = all(
                current_cells[i].get('colspan', 1) == above_cells[i].get('colspan', 1)
                for i in range(len(current_cells))
            )
            if not same_widths:
                continue

            # 조건 4: 위 행이 모두 data_ 또는 input_ 접두사
            above_names = [cell.get('name', '') for cell in above_cells]
            all_data_or_input = all(
                name.startswith('data_') or name.startswith('input_')
                for name in above_names if name
            )
            if not all_data_or_input:
                continue

            # 위 셀의 필드명 복사
            filled = 0
            for i, cell in enumerate(current_cells):
                above_name = above_cells[i].get('name')
                if above_name:
                    cell['name'] = above_name
                    cell['element'].set('name', above_name)
                    filled += 1

            if filled > 0:
                result['rows_filled'] += 1
                result['cells_filled'] += filled

        return result

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
    HWPX 파일에서 필드명 없는 셀에 필드명 자동 설정

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
