# -*- coding: utf-8 -*-
"""
HWPX 셀 상세 정보 추출 모듈

셀의 폰트, 정렬, 테두리, 배경색 등 상세 스타일 정보를 추출합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union, TYPE_CHECKING
from pathlib import Path
from io import BytesIO


@dataclass
class FontInfo:
    """폰트 정보"""
    name: str = ""
    size: int = 0  # HWPUNIT (1000 = 10pt)
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikeout: bool = False
    color: str = "#000000"

    def size_pt(self) -> float:
        return self.size / 100


@dataclass
class ParaInfo:
    """문단 정보"""
    para_id: str = ""  # p 요소의 id (HWPX에서는 동일할 수 있음)
    para_pr_id: str = ""  # paraPrIDRef (문단 스타일 참조 ID)
    style_id: str = ""  # styleIDRef (스타일 ID)
    align_h: str = "LEFT"  # JUSTIFY, LEFT, RIGHT, CENTER
    align_v: str = "BASELINE"
    line_spacing: int = 160  # percent

    # 텍스트
    text: str = ""

    # 줄 정보 (linesegarray에서 추출)
    line_count: int = 1
    height: int = 0  # HWPUNIT (문단 전체 높이)

    # 중첩 테이블 포함 여부
    has_nested_table: bool = False

    # 문단별 폰트 정보
    font: Optional['FontInfo'] = None


@dataclass
class BorderInfo:
    """테두리 정보"""
    left: str = "NONE"
    right: str = "NONE"
    top: str = "NONE"
    bottom: str = "NONE"
    bg_color: str = ""


@dataclass
class CellDetail:
    """셀 상세 정보"""
    # 테이블 정보
    table_id: str = ""

    # 위치 정보
    list_id: str = ""
    row: int = 0
    col: int = 0
    end_row: int = 0
    end_col: int = 0
    row_span: int = 1
    col_span: int = 1

    # 크기 (HWPUNIT)
    x: int = 0
    y: int = 0
    width: int = 0
    height: int = 0

    # 여백
    margin_left: int = 0
    margin_right: int = 0
    margin_top: int = 0
    margin_bottom: int = 0

    # 스타일 참조 ID
    char_pr_id: str = ""
    para_pr_id: str = ""
    border_fill_id: str = ""

    # 문단 정보 리스트 (하나의 셀에 여러 문단 가능)
    paragraphs: List[ParaInfo] = field(default_factory=list)

    # 폰트 정보
    font: FontInfo = field(default_factory=FontInfo)

    # 테두리/배경
    border: BorderInfo = field(default_factory=BorderInfo)

    # 텍스트
    text: str = ""

    # 필드 정보
    field_name: str = ""
    field_source: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            'table_id': self.table_id,
            'list_id': self.list_id,
            'row': self.row,
            'col': self.col,
            'end_row': self.end_row,
            'end_col': self.end_col,
            'row_span': self.row_span,
            'col_span': self.col_span,
            'x': self.x,
            'y': self.y,
            'width': self.width,
            'height': self.height,
            'width_pt': self.width / 100,
            'height_pt': self.height / 100,
            'margin_left': self.margin_left,
            'margin_right': self.margin_right,
            'margin_top': self.margin_top,
            'margin_bottom': self.margin_bottom,
            'font_name': self.font.name,
            'font_size': self.font.size,
            'font_size_pt': self.font.size_pt(),
            'bold': self.font.bold,
            'italic': self.font.italic,
            'underline': self.font.underline,
            'strikeout': self.font.strikeout,
            'font_color': self.font.color,
            'align_h': self.paragraphs[0].align_h if self.paragraphs else 'LEFT',
            'align_v': self.paragraphs[0].align_v if self.paragraphs else 'CENTER',
            'line_spacing': self.paragraphs[0].line_spacing if self.paragraphs else 160,
            'border_left': self.border.left,
            'border_right': self.border.right,
            'border_top': self.border.top,
            'border_bottom': self.border.bottom,
            'bg_color': self.border.bg_color,
            'text': self.text,
            'field_name': self.field_name,
            'field_source': self.field_source,
        }


class GetCellDetail:
    """
    HWPX 파일에서 셀 상세 정보 추출

    사용 예:
        extractor = GetCellDetail()
        cells = extractor.from_hwpx("document.hwpx")
        for cell in cells:
            print(cell.to_dict())
    """

    def __init__(self):
        # 스타일 정의 캐시
        self._char_props: Dict[str, Dict] = {}
        self._para_props: Dict[str, Dict] = {}
        self._border_fills: Dict[str, Dict] = {}
        self._fonts: Dict[str, str] = {}

    def _clear_caches(self):
        """스타일 캐시 초기화 (파일별 독립 처리를 위해)"""
        self._char_props.clear()
        self._para_props.clear()
        self._border_fills.clear()
        self._fonts.clear()

    def from_hwpx(self, hwpx_path: Union[str, Path]) -> List[CellDetail]:
        """HWPX 파일에서 모든 셀 상세 정보 추출"""
        self._clear_caches()
        hwpx_path = Path(hwpx_path)
        cells = []

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # 1. header.xml에서 스타일 정의 로드
            if 'Contents/header.xml' in zf.namelist():
                header_content = zf.read('Contents/header.xml')
                self._parse_header(header_content)

            # 2. section 파일들에서 셀 정보 추출
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                section_content = zf.read(section_file)
                section_cells = self._parse_section(section_content)
                cells.extend(section_cells)

        return cells

    def from_hwpx_by_table(self, hwpx_path: Union[str, Path]) -> List[List[CellDetail]]:
        """HWPX 파일에서 테이블별로 그룹화된 셀 상세 정보 추출"""
        self._clear_caches()
        hwpx_path = Path(hwpx_path)
        tables_cells = []

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # 1. header.xml에서 스타일 정의 로드
            if 'Contents/header.xml' in zf.namelist():
                header_content = zf.read('Contents/header.xml')
                self._parse_header(header_content)

            # 2. section 파일들에서 테이블별 셀 정보 추출
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                section_content = zf.read(section_file)
                section_tables = self._parse_section_by_table(section_content)
                tables_cells.extend(section_tables)

        return tables_cells

    def _parse_section_by_table(self, xml_content: bytes) -> List[List[CellDetail]]:
        """section XML에서 테이블별로 셀 정보 파싱 (중첩 테이블 순서 유지)"""
        tables_cells = []
        root = ET.parse(BytesIO(xml_content)).getroot()

        # 재귀적으로 테이블 찾기
        self._find_tables_recursive(root, tables_cells)

        return tables_cells

    def _find_tables_recursive(self, element, tables_cells: List[List[CellDetail]]):
        """재귀적으로 테이블을 찾아 순서대로 처리"""
        for child in element:
            if child.tag.endswith('}tbl'):
                # 테이블 ID 추출
                table_id = child.get('id', '')

                # 이 테이블의 직접 셀만 파싱 (중첩 테이블 제외)
                table_cells = self._parse_table_direct_cells(child, table_id)
                tables_cells.append(table_cells)

                # 셀 내부의 중첩 테이블 재귀 탐색
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue
                        # 셀 내부에서 중첩 테이블 찾기
                        self._find_tables_recursive(tc, tables_cells)
            else:
                # tbl이 아닌 요소 내부도 탐색
                self._find_tables_recursive(child, tables_cells)

    def _parse_table_direct_cells(self, tbl_element, table_id: str = "") -> List[CellDetail]:
        """테이블의 직접 셀만 파싱 (중첩 테이블 내부 셀 제외)"""
        cells = []

        for tr in tbl_element:
            if not tr.tag.endswith('}tr'):
                continue
            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue

                cell = CellDetail()
                cell.table_id = table_id
                cell.border_fill_id = tc.get('borderFillIDRef', '')

                # 테두리/배경 적용
                if cell.border_fill_id in self._border_fills:
                    bf = self._border_fills[cell.border_fill_id]
                    cell.border = BorderInfo(
                        left=bf['left'],
                        right=bf['right'],
                        top=bf['top'],
                        bottom=bf['bottom'],
                        bg_color=bf['bg_color'],
                    )

                # 셀 내부 요소 파싱
                for child in tc:
                    tag = child.tag.split('}')[-1]

                    if tag == 'subList':
                        cell.list_id = child.get('id', '')
                        # 문단들 파싱
                        self._parse_paragraphs(child, cell)

                    elif tag == 'cellAddr':
                        cell.col = int(child.get('colAddr', 0))
                        cell.row = int(child.get('rowAddr', 0))

                    elif tag == 'cellSpan':
                        cell.col_span = int(child.get('colSpan', 1))
                        cell.row_span = int(child.get('rowSpan', 1))
                        cell.end_col = cell.col + cell.col_span - 1
                        cell.end_row = cell.row + cell.row_span - 1

                    elif tag == 'cellSz':
                        cell.width = int(child.get('width', 0))
                        cell.height = int(child.get('height', 0))

                    elif tag == 'cellMargin':
                        cell.margin_left = int(child.get('left', 0))
                        cell.margin_right = int(child.get('right', 0))
                        cell.margin_top = int(child.get('top', 0))
                        cell.margin_bottom = int(child.get('bottom', 0))

                cells.append(cell)

        return cells

    def _parse_header(self, xml_content: bytes):
        """header.xml에서 스타일 정의 파싱"""
        root = ET.parse(BytesIO(xml_content)).getroot()

        # 폰트 정의 파싱 (HANGUL fontface만)
        for fontface in root.iter():
            if fontface.tag.endswith('}fontface') and fontface.get('lang') == 'HANGUL':
                for font in fontface:
                    if font.tag.endswith('}font'):
                        font_id = font.get('id', '')
                        face = font.get('face', '')
                        if font_id:
                            self._fonts[font_id] = face
                break  # HANGUL fontface만 처리

        # 문자 속성 파싱
        for elem in root.iter():
            if elem.tag.endswith('}charPr'):
                char_id = elem.get('id', '')
                if char_id:
                    self._char_props[char_id] = {
                        'height': int(elem.get('height', 1000)),
                        'textColor': elem.get('textColor', '#000000'),
                        'bold': elem.get('bold', '0') == '1',
                        'italic': elem.get('italic', '0') == '1',
                        'underline': elem.get('underline', 'NONE') != 'NONE',
                        'strikeout': elem.get('strikeout', 'NONE') != 'NONE',
                    }
                    # 폰트 참조 찾기
                    for child in elem:
                        if child.tag.endswith('}fontRef'):
                            hangul_ref = child.get('hangul', '0')
                            self._char_props[char_id]['font_ref'] = hangul_ref

        # 문단 속성 파싱
        for elem in root.iter():
            if elem.tag.endswith('}paraPr'):
                para_id = elem.get('id', '')
                if para_id:
                    self._para_props[para_id] = {
                        'align_h': 'LEFT',
                        'align_v': 'BASELINE',
                        'line_spacing': 160,
                    }
                    for child in elem:
                        if child.tag.endswith('}align'):
                            self._para_props[para_id]['align_h'] = child.get('horizontal', 'LEFT')
                            self._para_props[para_id]['align_v'] = child.get('vertical', 'BASELINE')
                        elif child.tag.endswith('}lineSpacing'):
                            self._para_props[para_id]['line_spacing'] = int(child.get('value', 160))

        # 테두리/배경 파싱
        for elem in root.iter():
            if elem.tag.endswith('}borderFill'):
                bf_id = elem.get('id', '')
                if bf_id:
                    self._border_fills[bf_id] = {
                        'left': 'NONE',
                        'right': 'NONE',
                        'top': 'NONE',
                        'bottom': 'NONE',
                        'bg_color': '',
                    }
                    for child in elem:
                        if child.tag.endswith('}leftBorder'):
                            self._border_fills[bf_id]['left'] = child.get('type', 'NONE')
                        elif child.tag.endswith('}rightBorder'):
                            self._border_fills[bf_id]['right'] = child.get('type', 'NONE')
                        elif child.tag.endswith('}topBorder'):
                            self._border_fills[bf_id]['top'] = child.get('type', 'NONE')
                        elif child.tag.endswith('}bottomBorder'):
                            self._border_fills[bf_id]['bottom'] = child.get('type', 'NONE')
                        elif child.tag.endswith('}fillBrush'):
                            for brush_child in child:
                                if brush_child.tag.endswith('}winBrush'):
                                    self._border_fills[bf_id]['bg_color'] = brush_child.get('faceColor', '')

    def _parse_section(self, xml_content: bytes) -> List[CellDetail]:
        """section XML에서 셀 정보 파싱"""
        cells = []
        root = ET.parse(BytesIO(xml_content)).getroot()

        # 테이블 내 셀 찾기
        for tc_elem in root.iter():
            if not tc_elem.tag.endswith('}tc'):
                continue

            cell = CellDetail()
            cell.border_fill_id = tc_elem.get('borderFillIDRef', '')

            # 테두리/배경 적용
            if cell.border_fill_id in self._border_fills:
                bf = self._border_fills[cell.border_fill_id]
                cell.border = BorderInfo(
                    left=bf['left'],
                    right=bf['right'],
                    top=bf['top'],
                    bottom=bf['bottom'],
                    bg_color=bf['bg_color'],
                )

            # 셀 내부 요소 파싱
            for child in tc_elem:
                tag = child.tag.split('}')[-1]

                if tag == 'subList':
                    cell.list_id = child.get('id', '')
                    # 문단들 파싱
                    self._parse_paragraphs(child, cell)

                elif tag == 'cellAddr':
                    cell.col = int(child.get('colAddr', 0))
                    cell.row = int(child.get('rowAddr', 0))

                elif tag == 'cellSpan':
                    cell.col_span = int(child.get('colSpan', 1))
                    cell.row_span = int(child.get('rowSpan', 1))
                    cell.end_col = cell.col + cell.col_span - 1
                    cell.end_row = cell.row + cell.row_span - 1

                elif tag == 'cellSz':
                    cell.width = int(child.get('width', 0))
                    cell.height = int(child.get('height', 0))

                elif tag == 'cellMargin':
                    cell.margin_left = int(child.get('left', 0))
                    cell.margin_right = int(child.get('right', 0))
                    cell.margin_top = int(child.get('top', 0))
                    cell.margin_bottom = int(child.get('bottom', 0))

            cells.append(cell)

        return cells

    def _parse_paragraphs(self, sublist_elem, cell: CellDetail):
        """subList 내의 문단들 파싱 (텍스트, 줄수, 높이, 폰트 포함)"""
        all_texts = []

        for p_elem in sublist_elem:
            if not p_elem.tag.endswith('}p'):
                continue

            para_info = ParaInfo()
            para_info.para_id = p_elem.get('id', '')
            para_pr_id = p_elem.get('paraPrIDRef', '')
            para_info.para_pr_id = para_pr_id
            para_info.style_id = p_elem.get('styleIDRef', '')

            # 문단 속성 적용
            if para_pr_id in self._para_props:
                pp = self._para_props[para_pr_id]
                para_info.align_h = pp['align_h']
                para_info.align_v = pp['align_v']
                para_info.line_spacing = pp['line_spacing']

            # 문단별 텍스트
            para_texts = []
            para_font_set = False  # 문단 폰트가 설정되었는지

            # run 요소에서 텍스트와 문자 속성 추출
            for child in p_elem:
                tag = child.tag.split('}')[-1]

                if tag == 'run':
                    char_pr_id = child.get('charPrIDRef', '')

                    # 문단의 첫 번째 run에서 폰트 정보 설정
                    if char_pr_id and not para_font_set:
                        para_font_set = True
                        if char_pr_id in self._char_props:
                            cp = self._char_props[char_pr_id]
                            para_info.font = FontInfo(
                                size=cp.get('height', 1000),
                                bold=cp.get('bold', False),
                                italic=cp.get('italic', False),
                                underline=cp.get('underline', False),
                                strikeout=cp.get('strikeout', False),
                                color=cp.get('textColor', '#000000'),
                            )
                            # 폰트 이름
                            font_ref = cp.get('font_ref', '0')
                            if font_ref in self._fonts:
                                para_info.font.name = self._fonts[font_ref]

                    # 셀의 기본 폰트 (첫 번째 문단의 첫 번째 run)
                    if char_pr_id and not cell.char_pr_id:
                        cell.char_pr_id = char_pr_id
                        if char_pr_id in self._char_props:
                            cp = self._char_props[char_pr_id]
                            cell.font = FontInfo(
                                size=cp.get('height', 1000),
                                bold=cp.get('bold', False),
                                italic=cp.get('italic', False),
                                underline=cp.get('underline', False),
                                strikeout=cp.get('strikeout', False),
                                color=cp.get('textColor', '#000000'),
                            )
                            font_ref = cp.get('font_ref', '0')
                            if font_ref in self._fonts:
                                cell.font.name = self._fonts[font_ref]

                    # 텍스트 추출
                    for t_elem in child:
                        if t_elem.tag.endswith('}t') and t_elem.text:
                            para_texts.append(t_elem.text)

                elif tag == 'linesegarray':
                    # lineseg에서 줄 수와 높이 추출
                    linesegs = [ls for ls in child if ls.tag.endswith('}lineseg')]
                    para_info.line_count = len(linesegs) if linesegs else 1

                    if linesegs:
                        # 문단 높이 = (마지막줄 vertpos + vertsize) - 첫줄 vertpos
                        first_ls = linesegs[0]
                        last_ls = linesegs[-1]
                        first_vertpos = int(first_ls.get('vertpos', 0))
                        last_vertpos = int(last_ls.get('vertpos', 0))
                        last_vertsize = int(last_ls.get('vertsize', 0))
                        para_info.height = (last_vertpos + last_vertsize) - first_vertpos

                elif tag == 'ctrl':
                    # ctrl 내에 테이블이 있는지 확인
                    for ctrl_child in child:
                        if ctrl_child.tag.endswith('}tbl'):
                            para_info.has_nested_table = True
                            break

            para_info.text = ''.join(para_texts)
            all_texts.extend(para_texts)
            cell.paragraphs.append(para_info)

        cell.text = ''.join(all_texts)


# 편의 함수
def get_cell_details(hwpx_path: Union[str, Path]) -> List[CellDetail]:
    """HWPX 파일에서 셀 상세 정보 추출"""
    extractor = GetCellDetail()
    return extractor.from_hwpx(hwpx_path)


if __name__ == "__main__":
    import sys
    import json

    hwpx_path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/c/hwp_xml/table.hwpx"

    print(f"파일: {hwpx_path}")
    print("=" * 50)

    cells = get_cell_details(hwpx_path)
    print(f"셀 개수: {len(cells)}")

    if cells:
        print("\n처음 3개 셀:")
        for i, cell in enumerate(cells[:3]):
            print(f"\n[셀 {i}]")
            d = cell.to_dict()
            for k, v in d.items():
                print(f"  {k}: {v}")
