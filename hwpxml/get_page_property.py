# -*- coding: utf-8 -*-
"""HWPX 페이지 속성 추출 모듈 (XML 기반)"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Optional, Dict, List, Union
from pathlib import Path
from io import BytesIO


# ============================================================
# 단위 변환 상수
# ============================================================

class Unit:
    """단위 변환 상수"""
    # 1 inch = 72 pt = 2.54 cm = 7200 HWPUNIT
    # 1 pt = 100 HWPUNIT
    # 1 cm = 2834.6 HWPUNIT (7200 / 2.54)

    HWPUNIT_PER_PT = 100
    HWPUNIT_PER_INCH = 7200
    HWPUNIT_PER_CM = 7200 / 2.54  # ≈ 2834.6
    HWPUNIT_PER_MM = 7200 / 25.4  # ≈ 283.46

    PT_PER_INCH = 72
    PT_PER_CM = 72 / 2.54  # ≈ 28.35

    # 엑셀 열 너비: 문자 단위 -> 포인트 (대략적 변환)
    EXCEL_CHAR_TO_PT = 7

    @staticmethod
    def hwpunit_to_pt(hwpunit: int) -> float:
        """HWPUNIT -> 포인트"""
        return hwpunit / Unit.HWPUNIT_PER_PT

    @staticmethod
    def pt_to_hwpunit(pt: float) -> int:
        """포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    @staticmethod
    def hwpunit_to_cm(hwpunit: int) -> float:
        """HWPUNIT -> cm"""
        return hwpunit / Unit.HWPUNIT_PER_CM

    @staticmethod
    def cm_to_hwpunit(cm: float) -> int:
        """cm -> HWPUNIT"""
        return int(cm * Unit.HWPUNIT_PER_CM)

    @staticmethod
    def hwpunit_to_mm(hwpunit: int) -> float:
        """HWPUNIT -> mm"""
        return hwpunit / Unit.HWPUNIT_PER_MM

    @staticmethod
    def mm_to_hwpunit(mm: float) -> int:
        """mm -> HWPUNIT"""
        return int(mm * Unit.HWPUNIT_PER_MM)

    @staticmethod
    def excel_pt_to_hwpunit(pt: float) -> int:
        """엑셀 포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    @staticmethod
    def excel_char_to_hwpunit(chars: float) -> int:
        """엑셀 문자 단위(열 너비) -> HWPUNIT"""
        pt = chars * Unit.EXCEL_CHAR_TO_PT
        return int(pt * Unit.HWPUNIT_PER_PT)


# ============================================================
# 페이지 메타 데이터 클래스
# ============================================================

@dataclass
class PageMargin:
    """페이지 여백 (HWPUNIT)"""
    left: int = 0
    right: int = 0
    top: int = 0
    bottom: int = 0
    header: int = 0
    footer: int = 0
    gutter: int = 0  # 제본 여백

    def to_dict(self) -> Dict:
        return {
            'left': self.left,
            'right': self.right,
            'top': self.top,
            'bottom': self.bottom,
            'header': self.header,
            'footer': self.footer,
            'gutter': self.gutter,
            # cm 변환
            'left_cm': round(Unit.hwpunit_to_cm(self.left), 2),
            'right_cm': round(Unit.hwpunit_to_cm(self.right), 2),
            'top_cm': round(Unit.hwpunit_to_cm(self.top), 2),
            'bottom_cm': round(Unit.hwpunit_to_cm(self.bottom), 2),
            'header_cm': round(Unit.hwpunit_to_cm(self.header), 2),
            'footer_cm': round(Unit.hwpunit_to_cm(self.footer), 2),
        }


@dataclass
class PageSize:
    """페이지 크기 (HWPUNIT)"""
    width: int = 0
    height: int = 0
    orientation: str = 'portrait'  # portrait / landscape / widely

    # 표준 용지 크기 (HWPUNIT)
    PAPER_SIZES = {
        'A4': (59528, 84188),      # 210mm x 297mm
        'A3': (84188, 119055),     # 297mm x 420mm
        'Letter': (61200, 79200),  # 8.5" x 11"
        'Legal': (61200, 100800),  # 8.5" x 14"
    }

    def to_dict(self) -> Dict:
        return {
            'width': self.width,
            'height': self.height,
            'orientation': self.orientation,
            'width_cm': round(Unit.hwpunit_to_cm(self.width), 2),
            'height_cm': round(Unit.hwpunit_to_cm(self.height), 2),
            'width_mm': round(Unit.hwpunit_to_mm(self.width), 1),
            'height_mm': round(Unit.hwpunit_to_mm(self.height), 1),
        }


@dataclass
class PageProperty:
    """페이지 속성 정보"""
    page_size: PageSize = field(default_factory=PageSize)
    margin: PageMargin = field(default_factory=PageMargin)

    # 본문 영역 (페이지 - 여백)
    content_width: int = 0
    content_height: int = 0

    # 섹션 정보
    section_id: str = ""
    text_direction: str = "HORIZONTAL"

    # 단 설정
    columns: int = 1
    column_gap: int = 0

    def calculate_content_area(self):
        """본문 영역 계산"""
        self.content_width = (self.page_size.width
                              - self.margin.left
                              - self.margin.right
                              - self.margin.gutter)
        self.content_height = (self.page_size.height
                               - self.margin.top
                               - self.margin.bottom
                               - self.margin.header
                               - self.margin.footer)

    def to_dict(self) -> Dict:
        return {
            'page_size': self.page_size.to_dict(),
            'margin': self.margin.to_dict(),
            'content_width': self.content_width,
            'content_height': self.content_height,
            'content_width_cm': round(Unit.hwpunit_to_cm(self.content_width), 2),
            'content_height_cm': round(Unit.hwpunit_to_cm(self.content_height), 2),
            'section_id': self.section_id,
            'text_direction': self.text_direction,
            'columns': self.columns,
            'column_gap': self.column_gap,
        }


# ============================================================
# HWPX 페이지 속성 추출 클래스
# ============================================================

class GetPageProperty:
    """
    HWPX 파일에서 페이지 속성 추출

    사용 예:
        parser = GetPageProperty()
        pages = parser.from_hwpx("document.hwpx")
        for page in pages:
            print(page.to_dict())
    """

    def __init__(self):
        self.namespaces = {
            'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
            'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
            'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
            'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
        }

    def from_hwpx(self, hwpx_path: Union[str, Path]) -> List[PageProperty]:
        """
        HWPX 파일에서 페이지 속성 추출

        Args:
            hwpx_path: HWPX 파일 경로

        Returns:
            PageProperty 리스트 (섹션별)
        """
        hwpx_path = Path(hwpx_path)
        pages = []

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # 섹션 파일 목록 찾기
            section_files = [f for f in zf.namelist()
                           if f.startswith('Contents/section') and f.endswith('.xml')]
            section_files.sort()

            for section_file in section_files:
                xml_content = zf.read(section_file)
                page_props = self._parse_section(xml_content)
                pages.extend(page_props)

        return pages

    def from_xml_string(self, xml_string: Union[str, bytes]) -> List[PageProperty]:
        """XML 문자열에서 페이지 속성 추출"""
        if isinstance(xml_string, str):
            xml_string = xml_string.encode('utf-8')
        return self._parse_section(xml_string)

    def _parse_section(self, xml_content: bytes) -> List[PageProperty]:
        """섹션 XML에서 페이지 속성 파싱"""
        pages = []

        try:
            root = ET.parse(BytesIO(xml_content)).getroot()
        except ET.ParseError:
            return pages

        # secPr (섹션 속성) 찾기
        for elem in root.iter():
            if elem.tag.endswith('}secPr'):
                page_prop = self._parse_secPr(elem)
                if page_prop:
                    pages.append(page_prop)

        return pages

    def _parse_secPr(self, secPr_elem) -> Optional[PageProperty]:
        """secPr 요소에서 페이지 속성 파싱"""
        page = PageProperty()

        # secPr 속성
        page.section_id = secPr_elem.get('id', '')
        page.text_direction = secPr_elem.get('textDirection', 'HORIZONTAL')

        # 단 간격
        space_columns = secPr_elem.get('spaceColumns')
        if space_columns:
            page.column_gap = int(space_columns)

        # pagePr (페이지 설정) 찾기
        for child in secPr_elem:
            if child.tag.endswith('}pagePr'):
                self._parse_pagePr(child, page)

        page.calculate_content_area()
        return page

    def _parse_pagePr(self, pagePr_elem, page: PageProperty):
        """pagePr 요소에서 페이지 크기/여백 파싱"""
        # 페이지 크기
        page.page_size.width = int(pagePr_elem.get('width', 0))
        page.page_size.height = int(pagePr_elem.get('height', 0))

        # 용지 방향
        landscape = pagePr_elem.get('landscape', 'NORMAL')
        if landscape == 'WIDELY':
            page.page_size.orientation = 'landscape'
        elif landscape == 'ROTATE':
            page.page_size.orientation = 'landscape'
        else:
            page.page_size.orientation = 'portrait'

        # margin 찾기
        for child in pagePr_elem:
            if child.tag.endswith('}margin'):
                page.margin.left = int(child.get('left', 0))
                page.margin.right = int(child.get('right', 0))
                page.margin.top = int(child.get('top', 0))
                page.margin.bottom = int(child.get('bottom', 0))
                page.margin.header = int(child.get('header', 0))
                page.margin.footer = int(child.get('footer', 0))
                page.margin.gutter = int(child.get('gutter', 0))


# ============================================================
# 편의 함수
# ============================================================

def get_page_property(hwpx_path: Union[str, Path]) -> Optional[PageProperty]:
    """
    HWPX 파일에서 첫 번째 페이지 속성 반환

    Args:
        hwpx_path: HWPX 파일 경로

    Returns:
        PageProperty 객체 (없으면 None)
    """
    parser = GetPageProperty()
    pages = parser.from_hwpx(hwpx_path)
    return pages[0] if pages else None


def get_all_page_properties(hwpx_path: Union[str, Path]) -> List[PageProperty]:
    """
    HWPX 파일에서 모든 섹션의 페이지 속성 반환

    Args:
        hwpx_path: HWPX 파일 경로

    Returns:
        PageProperty 리스트
    """
    parser = GetPageProperty()
    return parser.from_hwpx(hwpx_path)


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) > 1:
        hwpx_path = sys.argv[1]
    else:
        hwpx_path = "/mnt/c/hwp_xml/table.hwpx"

    print(f"파일: {hwpx_path}")
    print("=" * 50)

    parser = GetPageProperty()
    pages = parser.from_hwpx(hwpx_path)

    print(f"섹션 수: {len(pages)}")

    for i, page in enumerate(pages):
        print(f"\n[섹션 {i}]")
        print(json.dumps(page.to_dict(), indent=2, ensure_ascii=False))

    print("\n=== 단위 변환 테스트 ===")
    print(f"100 HWPUNIT = {Unit.hwpunit_to_pt(100)} pt")
    print(f"100 HWPUNIT = {Unit.hwpunit_to_cm(100):.4f} cm")
    print(f"1 cm = {Unit.cm_to_hwpunit(1)} HWPUNIT")
    print(f"1 mm = {Unit.mm_to_hwpunit(1)} HWPUNIT")
