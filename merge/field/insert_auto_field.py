# -*- coding: utf-8 -*-
"""
자동 생성된 필드명을 HWPX 셀에 삽입

테이블 셀에 자동 생성된 nc_name(필드명)을 tc 태그의 name 속성에 설정합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, List
import tempfile
import shutil
import os

from merge.table.parser import TableParser, NAMESPACES
from merge.table.models import TableInfo


class AutoFieldInserter:
    """자동 필드명을 HWPX 셀에 삽입"""

    def __init__(self, regenerate: bool = False):
        """
        Args:
            regenerate: True면 기존 필드명 무시하고 새로 생성
        """
        self.regenerate = regenerate
        self.parser = TableParser(auto_field_names=True, regenerate=regenerate)
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def insert_fields(self, hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> List[TableInfo]:
        """
        HWPX 파일에 자동 필드명 삽입

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_path: 출력 HWPX 파일 경로 (None이면 원본 덮어쓰기)

        Returns:
            처리된 테이블 정보 목록
        """
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        # 1. 테이블 파싱 및 자동 필드명 생성
        tables = self.parser.parse_tables(hwpx_path)
        print(f"테이블 {len(tables)}개 파싱 완료")

        # 2. 필드명 -> 셀 위치 매핑 생성
        field_mapping = {}  # (table_idx, row, col) -> field_name
        for table_idx, table in enumerate(tables):
            for (row, col), cell in table.cells.items():
                if cell.field_name:
                    field_mapping[(table_idx, row, col)] = cell.field_name

        # 3. HWPX 파일 수정
        temp_dir = tempfile.mkdtemp()
        try:
            # 압축 해제
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            # section 파일 수정
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

                # 테이블 찾아서 필드명 설정
                table_global_idx = self._process_section(root, field_mapping, table_global_idx)

                # 저장
                tree.write(section_path, encoding='utf-8', xml_declaration=True)

            # 다시 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        # mimetype은 압축 안 함
                        if arcname == 'mimetype':
                            zf.write(file_path, arcname, compress_type=zipfile.ZIP_STORED)
                        else:
                            zf.write(file_path, arcname)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        print(f"필드명 삽입 완료: {output_path}")
        return tables

    def _process_section(self, element, field_mapping: dict, table_idx: int) -> int:
        """section XML에서 테이블을 찾아 필드명 설정 (TableParser와 동일한 순서)"""
        for child in element:
            if child.tag.endswith('}tbl'):
                # 테이블 처리 (중첩 테이블은 _process_table 내에서 처리)
                table_idx = self._process_table(child, field_mapping, table_idx)
            elif not child.tag.endswith('}tc'):
                # tc가 아닌 element만 재귀 (tc 안의 tbl은 _process_table에서 처리)
                table_idx = self._process_section(child, field_mapping, table_idx)
        return table_idx

    def _process_table(self, tbl_elem, field_mapping: dict, table_idx: int) -> int:
        """테이블의 각 셀에 필드명 설정"""
        current_table_idx = table_idx
        cell_count = 0
        cell_elements = []  # 셀 element 수집 (중첩 테이블 처리용)

        for tr in tbl_elem:
            if not tr.tag.endswith('}tr'):
                continue

            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue

                cell_elements.append(tc)

                # 셀 위치 추출
                row, col = 0, 0
                for child in tc:
                    if child.tag.endswith('}cellAddr'):
                        row = int(child.get('rowAddr', 0))
                        col = int(child.get('colAddr', 0))
                        break

                # 필드명 찾기
                key = (current_table_idx, row, col)
                if key in field_mapping:
                    field_name = field_mapping[key]
                    tc.set('name', field_name)
                    cell_count += 1

        if cell_count > 0:
            print(f"  테이블 {current_table_idx}: {cell_count}개 셀에 필드명 설정")

        # 중첩 테이블 처리 (현재 테이블의 셀 element를 통해 재귀)
        next_table_idx = current_table_idx + 1
        for tc in cell_elements:
            next_table_idx = self._find_nested_tables(tc, field_mapping, next_table_idx)

        return next_table_idx

    def _find_nested_tables(self, element, field_mapping: dict, table_idx: int) -> int:
        """셀 element 내의 중첩 테이블 찾아 처리"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table_idx = self._process_table(child, field_mapping, table_idx)
            else:
                table_idx = self._find_nested_tables(child, field_mapping, table_idx)
        return table_idx


def insert_auto_fields(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None, regenerate: bool = False):
    """
    HWPX 파일에 자동 필드명 삽입

    Args:
        hwpx_path: 입력 HWPX 파일 경로
        output_path: 출력 HWPX 파일 경로 (None이면 원본 덮어쓰기)
        regenerate: True면 기존 필드명 무시하고 새로 생성
    """
    inserter = AutoFieldInserter(regenerate=regenerate)
    return inserter.insert_fields(hwpx_path, output_path)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법: python -m merge.insert_auto_field <input.hwpx> [output.hwpx] [--regenerate]")
        sys.exit(1)

    # --regenerate 옵션 확인
    regenerate = '--regenerate' in sys.argv
    args = [a for a in sys.argv[1:] if a != '--regenerate']

    input_path = args[0]
    output_path = args[1] if len(args) > 1 else None

    tables = insert_auto_fields(input_path, output_path, regenerate=regenerate)

    print(f"\n처리 완료: {len(tables)}개 테이블")
    for i, table in enumerate(tables):
        print(f"  테이블 {i}: {table.row_count}x{table.col_count}, 필드 {len(table.field_to_cell)}개")
