# -*- coding: utf-8 -*-
"""
선택 영역의 헤더 셀(배경색 있는 셀)을 기준으로 필드 이름 설정

기능:
1. HWPX XML에서 테이블의 선택 영역 분석
2. 선택 영역의 첫 행/열에서 배경색 있는 셀을 헤더로 식별
3. 상단 헤더 + 좌측 헤더 텍스트를 조합하여 필드 이름 생성
4. tc.name 속성에 필드 이름 설정

사용:
    from hwpxml.set_field_by_header import set_field_by_header

    # 선택 영역 지정 (row_start, row_end, col_start, col_end)
    count = set_field_by_header(
        hwpx_path="document.hwpx",
        table_index=0,
        selection=(1, 5, 1, 3)  # 2행~6행, 2열~4열 (0-indexed)
    )
"""

import zipfile
import tempfile
import shutil
import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Union
from pathlib import Path
from io import BytesIO


@dataclass
class CellInfo:
    """셀 정보"""
    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    text: str = ""
    bg_color: str = ""  # 배경색 (빈 문자열이면 배경색 없음)
    border_fill_id: str = ""
    tc_element: Optional[ET.Element] = None  # tc 요소 참조


@dataclass
class HeaderInfo:
    """헤더 정보"""
    row_headers: List[List[CellInfo]] = field(default_factory=list)  # 좌측 헤더 (여러 열 가능)
    col_headers: List[List[CellInfo]] = field(default_factory=list)  # 상단 헤더 (여러 행 가능)
    row_header_cols: int = 0  # 좌측 헤더 열 수
    col_header_rows: int = 0  # 상단 헤더 행 수


class SetFieldByHeader:
    """선택 영역의 헤더 기반 필드 이름 설정"""

    def __init__(self):
        self._border_fills: Dict[str, str] = {}  # id -> bg_color

    def _clear_caches(self):
        """캐시 초기화"""
        self._border_fills.clear()

    def _parse_header_xml(self, xml_content: bytes):
        """header.xml에서 배경색 정보 파싱"""
        root = ET.parse(BytesIO(xml_content)).getroot()

        for elem in root.iter():
            if elem.tag.endswith('}borderFill'):
                bf_id = elem.get('id', '')
                if bf_id:
                    bg_color = ""
                    for child in elem:
                        if child.tag.endswith('}fillBrush'):
                            for brush in child:
                                if brush.tag.endswith('}winBrush'):
                                    bg_color = brush.get('faceColor', '')
                                    break
                    self._border_fills[bf_id] = bg_color

    def _get_cell_text(self, tc: ET.Element) -> str:
        """tc 요소에서 텍스트 추출"""
        texts = []
        for elem in tc.iter():
            if elem.tag.endswith('}t') and elem.text:
                texts.append(elem.text)
        return ''.join(texts)

    def _parse_table_cells(self, tbl: ET.Element) -> Dict[Tuple[int, int], CellInfo]:
        """테이블의 모든 셀 정보 파싱"""
        cells = {}

        for tr in tbl:
            if not tr.tag.endswith('}tr'):
                continue

            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue

                cell = CellInfo(row=0, col=0)
                cell.tc_element = tc
                cell.border_fill_id = tc.get('borderFillIDRef', '')

                # 배경색 조회
                if cell.border_fill_id in self._border_fills:
                    cell.bg_color = self._border_fills[cell.border_fill_id]

                # 셀 위치/병합 정보
                for child in tc:
                    tag = child.tag.split('}')[-1]
                    if tag == 'cellAddr':
                        cell.row = int(child.get('rowAddr', 0))
                        cell.col = int(child.get('colAddr', 0))
                    elif tag == 'cellSpan':
                        cell.row_span = int(child.get('rowSpan', 1))
                        cell.col_span = int(child.get('colSpan', 1))

                # 텍스트 추출
                cell.text = self._get_cell_text(tc)

                cells[(cell.row, cell.col)] = cell

        return cells

    def _identify_headers(
        self,
        cells: Dict[Tuple[int, int], CellInfo],
        selection: Tuple[int, int, int, int]
    ) -> HeaderInfo:
        """
        선택 영역에서 헤더 식별

        Args:
            cells: 셀 정보 딕셔너리 {(row, col): CellInfo}
            selection: (row_start, row_end, col_start, col_end)

        Returns:
            HeaderInfo
        """
        row_start, row_end, col_start, col_end = selection
        header_info = HeaderInfo()

        # 1. 상단 헤더 식별 (선택 영역의 첫 행들 중 배경색 있는 행)
        col_header_rows = 0
        for row in range(row_start, row_end + 1):
            # 이 행의 모든 셀이 배경색이 있는지 확인
            row_has_bg = True
            for col in range(col_start, col_end + 1):
                cell = cells.get((row, col))
                if not cell or not cell.bg_color:
                    row_has_bg = False
                    break

            if row_has_bg:
                col_header_rows += 1
            else:
                break

        header_info.col_header_rows = col_header_rows

        # 2. 좌측 헤더 식별 (선택 영역의 첫 열들 중 배경색 있는 열)
        row_header_cols = 0
        for col in range(col_start, col_end + 1):
            # 이 열의 모든 셀이 배경색이 있는지 확인 (상단 헤더 제외)
            col_has_bg = True
            data_row_start = row_start + col_header_rows

            for row in range(data_row_start, row_end + 1):
                cell = cells.get((row, col))
                if not cell or not cell.bg_color:
                    col_has_bg = False
                    break

            if col_has_bg:
                row_header_cols += 1
            else:
                break

        header_info.row_header_cols = row_header_cols

        # 3. 상단 헤더 셀 수집 (데이터 열만)
        data_col_start = col_start + row_header_cols
        for row in range(row_start, row_start + col_header_rows):
            row_cells = []
            for col in range(data_col_start, col_end + 1):
                cell = cells.get((row, col))
                if cell:
                    row_cells.append(cell)
            header_info.col_headers.append(row_cells)

        # 4. 좌측 헤더 셀 수집 (데이터 행만)
        data_row_start = row_start + col_header_rows
        for row in range(data_row_start, row_end + 1):
            col_cells = []
            for col in range(col_start, col_start + row_header_cols):
                cell = cells.get((row, col))
                if cell:
                    col_cells.append(cell)
            header_info.row_headers.append(col_cells)

        return header_info

    def _generate_field_name(
        self,
        row_idx: int,  # 데이터 영역 내 행 인덱스 (0부터)
        col_idx: int,  # 데이터 영역 내 열 인덱스 (0부터)
        header_info: HeaderInfo
    ) -> str:
        """
        헤더 텍스트를 조합하여 필드 이름 생성

        Args:
            row_idx: 데이터 영역 내 행 인덱스
            col_idx: 데이터 영역 내 열 인덱스
            header_info: 헤더 정보

        Returns:
            필드 이름 (예: "2024년 1분기 매출")
        """
        parts = []

        # 상단 헤더 텍스트 (여러 행이면 띄어쓰기로 연결)
        col_header_texts = []
        for row_cells in header_info.col_headers:
            if col_idx < len(row_cells):
                cell = row_cells[col_idx]
                if cell.text.strip():
                    col_header_texts.append(cell.text.strip())

        if col_header_texts:
            parts.append(' '.join(col_header_texts))

        # 좌측 헤더 텍스트 (여러 열이면 띄어쓰기로 연결)
        row_header_texts = []
        if row_idx < len(header_info.row_headers):
            for cell in header_info.row_headers[row_idx]:
                if cell.text.strip():
                    row_header_texts.append(cell.text.strip())

        if row_header_texts:
            parts.append(' '.join(row_header_texts))

        # 상단 + 좌측 조합
        return ' '.join(parts) if parts else ""

    def set_field_names(
        self,
        hwpx_path: Union[str, Path],
        table_index: int,
        selection: Tuple[int, int, int, int],
        output_path: Optional[Union[str, Path]] = None
    ) -> int:
        """
        선택 영역의 헤더 기반으로 필드 이름 설정

        Args:
            hwpx_path: HWPX 파일 경로
            table_index: 테이블 인덱스 (0부터)
            selection: (row_start, row_end, col_start, col_end) 선택 영역
            output_path: 출력 파일 경로 (없으면 원본 파일 수정)

        Returns:
            설정된 필드 수
        """
        self._clear_caches()
        hwpx_path = Path(hwpx_path)

        if output_path is None:
            output_path = hwpx_path
        else:
            output_path = Path(output_path)

        temp_dir = tempfile.mkdtemp()
        count = 0

        try:
            # 1. HWPX 압축 해제
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                zf.extractall(temp_dir)

            # 2. header.xml에서 배경색 정보 로드
            header_path = os.path.join(temp_dir, 'Contents', 'header.xml')
            if os.path.exists(header_path):
                with open(header_path, 'rb') as f:
                    self._parse_header_xml(f.read())

            # 3. section 파일 처리
            contents_dir = os.path.join(temp_dir, 'Contents')
            section_files = sorted([
                f for f in os.listdir(contents_dir)
                if f.startswith('section') and f.endswith('.xml')
            ])

            current_table_idx = 0
            target_table = None

            for section_file in section_files:
                section_path = os.path.join(contents_dir, section_file)
                tree = ET.parse(section_path)
                root = tree.getroot()

                # 테이블 찾기
                for elem in root.iter():
                    if elem.tag.endswith('}tbl'):
                        if current_table_idx == table_index:
                            target_table = elem

                            # 셀 정보 파싱
                            cells = self._parse_table_cells(target_table)

                            # 헤더 식별
                            header_info = self._identify_headers(cells, selection)

                            print(f"테이블 {table_index}:")
                            print(f"  상단 헤더: {header_info.col_header_rows}행")
                            print(f"  좌측 헤더: {header_info.row_header_cols}열")

                            # 데이터 영역 필드 이름 설정
                            row_start, row_end, col_start, col_end = selection
                            data_row_start = row_start + header_info.col_header_rows
                            data_col_start = col_start + header_info.row_header_cols

                            for row in range(data_row_start, row_end + 1):
                                for col in range(data_col_start, col_end + 1):
                                    cell = cells.get((row, col))
                                    if cell and cell.tc_element is not None:
                                        # 데이터 영역 내 상대 인덱스
                                        rel_row = row - data_row_start
                                        rel_col = col - data_col_start

                                        field_name = self._generate_field_name(
                                            rel_row, rel_col, header_info
                                        )

                                        if field_name:
                                            cell.tc_element.set('name', field_name)
                                            count += 1
                                            print(f"  [{row},{col}] = {field_name}")

                            break
                        current_table_idx += 1

                if target_table is not None:
                    # 수정된 XML 저장
                    tree.write(section_path, encoding='utf-8', xml_declaration=True)
                    break

            # 4. HWPX 다시 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zf.write(file_path, arcname)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return count


def set_field_by_header(
    hwpx_path: Union[str, Path],
    table_index: int,
    selection: Tuple[int, int, int, int],
    output_path: Optional[Union[str, Path]] = None
) -> int:
    """
    선택 영역의 헤더 기반으로 필드 이름 설정 (편의 함수)

    Args:
        hwpx_path: HWPX 파일 경로
        table_index: 테이블 인덱스 (0부터)
        selection: (row_start, row_end, col_start, col_end) 선택 영역
        output_path: 출력 파일 경로 (없으면 원본 파일 수정)

    Returns:
        설정된 필드 수
    """
    setter = SetFieldByHeader()
    return setter.set_field_names(hwpx_path, table_index, selection, output_path)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 6:
        print("사용법: python set_field_by_header.py <hwpx_path> <table_index> <row_start> <row_end> <col_start> <col_end>")
        print()
        print("예시: python set_field_by_header.py doc.hwpx 0 0 5 0 3")
        print("  - 테이블 0의 0~5행, 0~3열 영역에서 헤더 기반 필드 설정")
        sys.exit(1)

    hwpx_path = sys.argv[1]
    table_index = int(sys.argv[2])
    row_start = int(sys.argv[3])
    row_end = int(sys.argv[4])
    col_start = int(sys.argv[5])
    col_end = int(sys.argv[6])

    print("=" * 60)
    print("헤더 기반 필드 이름 설정")
    print("=" * 60)
    print(f"파일: {hwpx_path}")
    print(f"테이블: {table_index}")
    print(f"선택 영역: ({row_start}, {row_end}) x ({col_start}, {col_end})")
    print()

    count = set_field_by_header(
        hwpx_path,
        table_index,
        (row_start, row_end, col_start, col_end)
    )

    print()
    print(f"설정된 필드: {count}개")
