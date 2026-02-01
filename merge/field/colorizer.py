# -*- coding: utf-8 -*-
"""
필드명별 색상 설정 모듈

같은 필드명(nc.name)을 가진 셀들에 동일한 배경색을 설정합니다.

접두사별 기본 색상:
- header_: 회색 (#D1D1D1)
- add_: 연한 파랑 (#E6F3FF)
- stub_: 연한 노랑 (#FFFFD0)
- gstub_: 연한 주황 (#FFE4C4)
- input_: 연한 초록 (#E8FFE8)
- data_: 연한 보라 (#F0E6FF)
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Union, Dict, List, Tuple, Optional
import tempfile
import shutil
import os
import hashlib

from ..table.parser import NAMESPACES


# 접두사별 기본 색상
PREFIX_COLORS = {
    'header_': '#D1D1D1',  # 회색
    'add_': '#E6F3FF',     # 연한 파랑
    'stub_': '#FFFFD0',    # 연한 노랑
    'gstub_': '#FFE4C4',   # 연한 주황
    'input_': '#E8FFE8',   # 연한 초록
    'data_': '#F0E6FF',    # 연한 보라
}

# 같은 nc_name끼리 구분하기 위한 색상 팔레트
COLOR_PALETTE = [
    '#FFDAB9',  # peachpuff
    '#E6E6FA',  # lavender
    '#F0FFF0',  # honeydew
    '#FFF0F5',  # lavenderblush
    '#F5FFFA',  # mintcream
    '#FFF5EE',  # seashell
    '#F0FFFF',  # azure
    '#FFFAF0',  # floralwhite
    '#F5F5DC',  # beige
    '#FAFAD2',  # lightgoldenrodyellow
    '#E0FFFF',  # lightcyan
    '#FFE4E1',  # mistyrose
    '#DCDCDC',  # gainsboro
    '#B0E0E6',  # powderblue
    '#DDA0DD',  # plum
    '#98FB98',  # palegreen
    '#AFEEEE',  # paleturquoise
    '#DB7093',  # palevioletred
    '#FFEFD5',  # papayawhip
    '#FFB6C1',  # lightpink
]


for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


def _hash_to_color_index(name: str) -> int:
    """필드명을 색상 인덱스로 변환"""
    h = hashlib.md5(name.encode()).hexdigest()
    return int(h[:8], 16) % len(COLOR_PALETTE)


class FieldColorizer:
    """필드명별 색상 설정"""

    def __init__(self, use_prefix_colors: bool = True, use_unique_colors: bool = True):
        """
        Args:
            use_prefix_colors: True면 접두사별 기본 색상 사용
            use_unique_colors: True면 같은 nc_name끼리 동일한 고유 색상 사용
        """
        self.use_prefix_colors = use_prefix_colors
        self.use_unique_colors = use_unique_colors
        self.border_fills: Dict[str, int] = {}  # color -> id
        self.next_border_fill_id = 100  # 기존 ID와 충돌 방지

    def colorize(self, hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Path:
        """HWPX 파일의 셀에 필드명별 색상 설정"""
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        temp_dir = tempfile.mkdtemp()
        try:
            # HWPX 압축 해제
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            # 기존 borderFill ID 수집
            header_path = os.path.join(temp_dir, 'Contents', 'header.xml')
            self._collect_existing_border_fills(header_path)

            # section 파일들 처리 - 필드명 수집 (순서 유지)
            contents_dir = os.path.join(temp_dir, 'Contents')
            field_names = self._collect_field_names_ordered(contents_dir)
            print(f"필드명 {len(field_names)}개 수집")

            # 필드명별 색상 결정
            field_colors = self._assign_colors(field_names)
            print(f"색상 {len(set(field_colors.values()))}종 할당")

            # 새 borderFill 추가
            self._add_border_fills_to_header(header_path, field_colors)

            # section 파일들 처리 - borderFillIDRef 설정
            self._apply_colors_to_sections(contents_dir, field_colors)

            # HWPX 재압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        if arcname == 'mimetype':
                            zf.write(file_path, arcname, compress_type=zipfile.ZIP_STORED)
                        else:
                            zf.write(file_path, arcname)

            print(f"색상 설정 완료: {output_path}")
            return output_path

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _collect_existing_border_fills(self, header_path: str):
        """기존 borderFill ID 수집"""
        tree = ET.parse(header_path)
        root = tree.getroot()

        max_id = 0
        for elem in root.iter():
            if elem.tag.endswith('}borderFill'):
                bf_id = int(elem.get('id', 0))
                if bf_id > max_id:
                    max_id = bf_id

        self.next_border_fill_id = max_id + 1

    def _collect_field_names_ordered(self, contents_dir: str) -> List[str]:
        """section 파일들에서 필드명을 테이블 내 순서대로 수집"""
        field_names_ordered = []  # 순서 유지
        field_names_set = set()   # 중복 체크

        section_files = sorted([
            f for f in os.listdir(contents_dir)
            if f.startswith('section') and f.endswith('.xml')
        ])

        for section_file in section_files:
            section_path = os.path.join(contents_dir, section_file)
            tree = ET.parse(section_path)
            root = tree.getroot()

            for elem in root.iter():
                if elem.tag.endswith('}tc'):
                    name = elem.get('name', '')
                    if name and name not in field_names_set:
                        field_names_ordered.append(name)
                        field_names_set.add(name)

        return field_names_ordered

    def _assign_colors(self, field_names: List[str]) -> Dict[str, str]:
        """필드명별 색상 할당 (인접한 필드끼리 다른 색상)"""
        field_colors = {}
        used_colors = []  # 최근 사용된 색상 추적

        for name in field_names:
            # 인접한 셀과 다른 색상 선택
            available_colors = [c for c in COLOR_PALETTE if c not in used_colors[-5:]]
            if not available_colors:
                available_colors = COLOR_PALETTE.copy()

            # 해시 기반으로 available 중에서 선택
            h = hashlib.md5(name.encode()).hexdigest()
            idx = int(h[:8], 16) % len(available_colors)
            color = available_colors[idx]

            field_colors[name] = color
            used_colors.append(color)

        return field_colors

    def _add_border_fills_to_header(self, header_path: str, field_colors: Dict[str, str]):
        """header.xml에 새 borderFill 추가"""
        tree = ET.parse(header_path)
        root = tree.getroot()

        # borderFills 요소 찾기
        border_fills_elem = None
        for elem in root.iter():
            if elem.tag.endswith('}borderFills'):
                border_fills_elem = elem
                break

        if border_fills_elem is None:
            print("borderFills 요소를 찾을 수 없습니다.")
            return

        # 기존 borderFill의 네임스페이스 접두사 확인
        ns_head = ''
        ns_core = ''
        for child in border_fills_elem:
            if child.tag.endswith('}borderFill'):
                ns_head = child.tag.rsplit('}', 1)[0] + '}'
                # fillBrush의 네임스페이스 찾기
                for sub in child.iter():
                    if sub.tag.endswith('}fillBrush'):
                        ns_core = sub.tag.rsplit('}', 1)[0] + '}'
                        break
                break

        if not ns_head:
            ns_head = '{http://www.hancom.co.kr/hwpml/2011/head}'
        if not ns_core:
            ns_core = '{http://www.hancom.co.kr/hwpml/2011/core}'

        ns_prefix = ns_head

        # 색상별로 borderFill 생성
        unique_colors = set(field_colors.values())

        for color in unique_colors:
            if color in self.border_fills:
                continue

            bf_id = self.next_border_fill_id
            self.next_border_fill_id += 1
            self.border_fills[color] = bf_id

            # borderFill 요소 생성
            border_fill = ET.SubElement(border_fills_elem, f'{ns_prefix}borderFill')
            border_fill.set('id', str(bf_id))
            border_fill.set('threeD', '0')
            border_fill.set('shadow', '0')
            border_fill.set('centerLine', 'NONE')
            border_fill.set('breakCellSeparateLine', '0')

            # slash
            slash = ET.SubElement(border_fill, f'{ns_prefix}slash')
            slash.set('type', 'NONE')
            slash.set('Crooked', '0')
            slash.set('isCounter', '0')

            # backSlash
            back_slash = ET.SubElement(border_fill, f'{ns_prefix}backSlash')
            back_slash.set('type', 'NONE')
            back_slash.set('Crooked', '0')
            back_slash.set('isCounter', '0')

            # borders
            for border_name in ['leftBorder', 'rightBorder', 'topBorder', 'bottomBorder']:
                border = ET.SubElement(border_fill, f'{ns_prefix}{border_name}')
                border.set('type', 'SOLID')
                border.set('width', '0.12 mm')
                border.set('color', '#000000')

            # diagonal
            diagonal = ET.SubElement(border_fill, f'{ns_prefix}diagonal')
            diagonal.set('type', 'SOLID')
            diagonal.set('width', '0.1 mm')
            diagonal.set('color', '#000000')

            # fillBrush (core 네임스페이스 사용)
            fill_brush = ET.SubElement(border_fill, f'{ns_core}fillBrush')
            win_brush = ET.SubElement(fill_brush, f'{ns_core}winBrush')
            win_brush.set('faceColor', color)
            win_brush.set('hatchColor', '#000000')
            win_brush.set('alpha', '0')

        # borderFills itemCnt 업데이트
        border_fills_elem.set('itemCnt', str(len(list(border_fills_elem))))

        tree.write(header_path, encoding='utf-8', xml_declaration=True)
        print(f"  borderFill {len(unique_colors)}개 추가")

    def _apply_colors_to_sections(self, contents_dir: str, field_colors: Dict[str, str]):
        """section 파일들에 borderFillIDRef 설정"""
        section_files = sorted([
            f for f in os.listdir(contents_dir)
            if f.startswith('section') and f.endswith('.xml')
        ])

        total_cells = 0
        for section_file in section_files:
            section_path = os.path.join(contents_dir, section_file)
            tree = ET.parse(section_path)
            root = tree.getroot()

            cell_count = 0
            for elem in root.iter():
                if elem.tag.endswith('}tc'):
                    name = elem.get('name', '')
                    if name and name in field_colors:
                        color = field_colors[name]
                        if color in self.border_fills:
                            bf_id = self.border_fills[color]
                            elem.set('borderFillIDRef', str(bf_id))
                            cell_count += 1

            if cell_count > 0:
                tree.write(section_path, encoding='utf-8', xml_declaration=True)
                total_cells += cell_count

        print(f"  {total_cells}개 셀에 색상 적용")


def colorize_by_field(hwpx_path: Union[str, Path], output_path: Union[str, Path] = None) -> Path:
    """HWPX 파일의 셀에 필드명별 색상 설정"""
    colorizer = FieldColorizer()
    return colorizer.colorize(hwpx_path, output_path)


def auto_field_and_colorize(
    hwpx_path: Union[str, Path],
    output_path: Union[str, Path] = None,
    regenerate: bool = False
) -> Path:
    """
    자동 필드명 생성 + 색상 설정을 한 번에 수행

    Args:
        hwpx_path: 입력 HWPX 파일 경로
        output_path: 출력 HWPX 파일 경로 (없으면 입력 파일 덮어쓰기)
        regenerate: True면 기존 필드명 무시하고 재생성

    Returns:
        출력 파일 경로
    """
    from .auto_field import insert_auto_fields

    hwpx_path = Path(hwpx_path)
    output_path = Path(output_path) if output_path else hwpx_path

    # 1. 자동 필드명 생성
    print("=== 1단계: 자동 필드명 생성 ===")
    insert_auto_fields(hwpx_path, output_path, regenerate=regenerate)

    # 2. 색상 설정
    print("\n=== 2단계: 필드명별 색상 설정 ===")
    colorize_by_field(output_path, output_path)

    return output_path


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법: python -m merge.field.colorizer <input.hwpx> [output.hwpx] [--regenerate] [--auto]")
        print("  --auto: 자동 필드명 생성 + 색상 설정")
        print("  --regenerate: 기존 필드명 무시하고 재생성")
        sys.exit(1)

    auto_mode = '--auto' in sys.argv
    regenerate = '--regenerate' in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith('--')]

    input_path = args[0]
    output_path = args[1] if len(args) > 1 else None

    if auto_mode:
        auto_field_and_colorize(input_path, output_path, regenerate=regenerate)
    else:
        colorize_by_field(input_path, output_path)
