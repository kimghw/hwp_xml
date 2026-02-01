# -*- coding: utf-8 -*-
"""
필드 시각화 모듈

필드명이 없는 셀에 빨간 배경색을 넣거나,
필드명을 파란색 텍스트로 셀에 표시합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, List, Dict
import tempfile
import shutil
import os
import re

from ..table.parser import TableParser, NAMESPACES
from ..table.models import TableInfo


# 스타일 ID
RED_BORDER_FILL_ID = "9998"  # 빨간 배경색용
BLUE_CHAR_PR_ID = "9999"     # 파란 텍스트용


class FieldVisualizer:
    """필드 시각화 (빨간 배경, 파란 텍스트)"""

    def __init__(self, regenerate: bool = False):
        self.regenerate = regenerate
        self.parser = TableParser(auto_field_names=True, regenerate=regenerate)
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def highlight_empty_fields(self, hwpx_path: Union[str, Path],
                                output_path: Union[str, Path] = None) -> Dict[str, int]:
        """
        필드명이 없는 셀에 빨간 배경색 적용

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_path: 출력 HWPX 파일 경로

        Returns:
            처리 결과 {'tables': 테이블 수, 'cells_highlighted': 강조된 셀 수}
        """
        hwpx_path = Path(hwpx_path)
        if output_path:
            output_path = Path(output_path)
        else:
            output_dir = hwpx_path.parent / 'output'
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / f"{hwpx_path.stem}_visual{hwpx_path.suffix}"

        result = {'tables': 0, 'cells_highlighted': 0}

        temp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            # header.xml에 빨간 배경색 borderFill 추가
            self._add_red_border_fill(temp_dir)

            contents_dir = os.path.join(temp_dir, 'Contents')
            section_files = sorted([
                f for f in os.listdir(contents_dir)
                if f.startswith('section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                section_path = os.path.join(contents_dir, section_file)
                tree = ET.parse(section_path)
                root = tree.getroot()

                section_result = self._highlight_section(root)
                result['tables'] += section_result['tables']
                result['cells_highlighted'] += section_result['cells_highlighted']

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

        print(f"빨간 배경 강조 완료: {output_path}")
        print(f"  테이블 {result['tables']}개, {result['cells_highlighted']}셀 강조")
        return result

    def insert_field_text(self, hwpx_path: Union[str, Path],
                          output_path: Union[str, Path] = None) -> List[TableInfo]:
        """
        필드명을 셀 안에 파란색 텍스트로 삽입

        Args:
            hwpx_path: 입력 HWPX 파일 경로
            output_path: 출력 HWPX 파일 경로

        Returns:
            처리된 테이블 정보 목록
        """
        hwpx_path = Path(hwpx_path)
        if output_path:
            output_path = Path(output_path)
        else:
            output_dir = hwpx_path.parent / 'output'
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / f"{hwpx_path.stem}_visual{hwpx_path.suffix}"

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

            self._add_blue_char_style(temp_dir)

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

                table_global_idx = self._insert_text_section(root, field_mapping, table_global_idx)

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

        print(f"필드명 텍스트 삽입 완료: {output_path}")
        return tables

    def _add_red_border_fill(self, temp_dir: str):
        """header.xml에 빨간 배경색 borderFill 추가"""
        header_path = os.path.join(temp_dir, 'Contents', 'header.xml')

        with open(header_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # borderFills의 itemCnt 증가
        match = re.search(r'<hh:borderFills itemCnt="(\d+)">', content)
        if match:
            old_count = int(match.group(1))
            new_count = old_count + 1
            content = content.replace(
                f'<hh:borderFills itemCnt="{old_count}">',
                f'<hh:borderFills itemCnt="{new_count}">'
            )

        # 빨간 배경색 borderFill 추가
        red_border_fill = f'''<hh:borderFill id="{RED_BORDER_FILL_ID}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" crooked="0" isCounter="0"/><hh:backSlash type="NONE" crooked="0" isCounter="0"/><hh:leftBorder type="NONE" width="0.12mm" color="#000000"/><hh:rightBorder type="NONE" width="0.12mm" color="#000000"/><hh:topBorder type="NONE" width="0.12mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.12mm" color="#000000"/><hh:diagonal type="NONE" width="0.12mm" color="#000000"/><hh:fillBrush><hh:winBrush faceColor="#FF6B6B" hatchColor="#FF0000" alpha="0.3"/></hh:fillBrush></hh:borderFill>'''

        content = content.replace(
            '</hh:borderFills>',
            red_border_fill + '</hh:borderFills>'
        )

        with open(header_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def _add_blue_char_style(self, temp_dir: str):
        """header.xml에 파란색 charPr 스타일 추가"""
        header_path = os.path.join(temp_dir, 'Contents', 'header.xml')

        with open(header_path, 'r', encoding='utf-8') as f:
            content = f.read()

        match = re.search(r'<hh:charProperties itemCnt="(\d+)">', content)
        if match:
            old_count = int(match.group(1))
            new_count = old_count + 1
            content = content.replace(
                f'<hh:charProperties itemCnt="{old_count}">',
                f'<hh:charProperties itemCnt="{new_count}">'
            )

        blue_char_pr = f'''<hh:charPr id="{BLUE_CHAR_PR_ID}" height="800" textColor="#0000FF" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2"><hh:fontRef hangul="4" latin="5" hanja="5" japanese="5" other="5" symbol="5" user="5"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:strikeout type="NONE" shape="SOLID" color="#000000"/><hh:outline type="NONE"/><hh:shadow type="NONE" color="#b2b2b2" offsetX="10" offsetY="10"/><hh:underline type="NONE" shape="SOLID" color="#000000"/><hh:charShadow type="DISCRETE" x="7" y="7" color="#b2b2b2"/></hh:charPr>'''

        content = content.replace(
            '</hh:charProperties>',
            blue_char_pr + '</hh:charProperties>'
        )

        with open(header_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def _highlight_section(self, element) -> Dict[str, int]:
        """section XML에서 테이블의 빈 필드 셀 강조"""
        result = {'tables': 0, 'cells_highlighted': 0}

        for child in element:
            if child.tag.endswith('}tbl'):
                table_result = self._highlight_table(child)
                result['tables'] += 1
                result['cells_highlighted'] += table_result
            elif not child.tag.endswith('}tc'):
                sub_result = self._highlight_section(child)
                result['tables'] += sub_result['tables']
                result['cells_highlighted'] += sub_result['cells_highlighted']

        return result

    def _highlight_table(self, tbl_elem) -> int:
        """테이블의 필드명 없는 셀에 빨간 배경색 적용"""
        highlighted = 0

        for tr in tbl_elem:
            if not tr.tag.endswith('}tr'):
                continue

            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue

                # 필드명 확인
                if not tc.get('name'):
                    # borderFillIDRef를 빨간색으로 변경
                    for child in tc:
                        if child.tag.endswith('}cellMargin'):
                            child.set('borderFillIDRef', RED_BORDER_FILL_ID)
                            highlighted += 1
                            break

                # 중첩 테이블 처리
                highlighted += self._highlight_nested(tc)

        return highlighted

    def _highlight_nested(self, element) -> int:
        """중첩 테이블 처리"""
        highlighted = 0
        for child in element:
            if child.tag.endswith('}tbl'):
                highlighted += self._highlight_table(child)
            else:
                highlighted += self._highlight_nested(child)
        return highlighted

    def _insert_text_section(self, element, field_mapping: dict, table_idx: int) -> int:
        """section XML에서 테이블에 필드명 텍스트 삽입"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table_idx = self._insert_text_table(child, field_mapping, table_idx)
            elif not child.tag.endswith('}tc'):
                table_idx = self._insert_text_section(child, field_mapping, table_idx)
        return table_idx

    def _insert_text_table(self, tbl_elem, field_mapping: dict, table_idx: int) -> int:
        """테이블 셀에 필드명 텍스트 삽입"""
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
                    self._add_field_text_to_cell(tc, field_name)
                    cell_count += 1

        if cell_count > 0:
            print(f"  테이블 {current_table_idx}: {cell_count}개 셀에 필드명 삽입")

        next_table_idx = current_table_idx + 1
        for tc in cell_elements:
            next_table_idx = self._insert_text_nested(tc, field_mapping, next_table_idx)

        return next_table_idx

    def _insert_text_nested(self, element, field_mapping: dict, table_idx: int) -> int:
        """중첩 테이블에 텍스트 삽입"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table_idx = self._insert_text_table(child, field_mapping, table_idx)
            else:
                table_idx = self._insert_text_nested(child, field_mapping, table_idx)
        return table_idx

    def _add_field_text_to_cell(self, tc_elem, field_name: str):
        """셀에 파란색 필드명 텍스트 추가"""
        sublist = None
        for child in tc_elem:
            if child.tag.endswith('}subList'):
                sublist = child
                break

        if sublist is None:
            return

        p_elem = None
        for child in sublist:
            if child.tag.endswith('}p'):
                p_elem = child
                break

        if p_elem is None:
            return

        ns = ''
        if '}' in p_elem.tag:
            ns = p_elem.tag.split('}')[0] + '}'

        new_run = ET.Element(f'{ns}run')
        new_run.set('charPrIDRef', BLUE_CHAR_PR_ID)

        t_elem = ET.SubElement(new_run, f'{ns}t')
        t_elem.text = f' [{field_name}]'

        insert_idx = len(list(p_elem))
        for i, child in enumerate(p_elem):
            if child.tag.endswith('}linesegarray'):
                insert_idx = i
                break

        p_elem.insert(insert_idx, new_run)


def highlight_empty_fields(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Dict[str, int]:
    """필드명이 없는 셀에 빨간 배경색 적용"""
    visualizer = FieldVisualizer()
    return visualizer.highlight_empty_fields(hwpx_path, output_path)


def insert_field_text(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None, regenerate: bool = False):
    """필드명을 셀 안에 파란색 텍스트로 삽입"""
    visualizer = FieldVisualizer(regenerate=regenerate)
    return visualizer.insert_field_text(hwpx_path, output_path)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("사용법:")
        print("  python -m merge.field.visualizer highlight <input.hwpx> [output.hwpx]")
        print("  python -m merge.field.visualizer text <input.hwpx> [output.hwpx]")
        sys.exit(1)

    mode = sys.argv[1]
    input_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None

    if mode == 'highlight':
        highlight_empty_fields(input_path, output_path)
    elif mode == 'text':
        insert_field_text(input_path, output_path)
    else:
        print(f"알 수 없는 모드: {mode}")
        print("사용 가능한 모드: highlight, text")
        sys.exit(1)
