# -*- coding: utf-8 -*-
"""HWPX 테이블 → Excel 변환 모듈"""

import sys
from pathlib import Path
from typing import Optional, Union, List

# 프로젝트 루트 경로 설정
_project_root = Path(__file__).parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

from hwpxml.get_table_property import GetTableProperty, TableProperty
from hwpxml.get_page_property import GetPageProperty, PageProperty, Unit


class HwpxToExcel:
    """
    HWPX 테이블을 Excel로 변환

    사용 예:
        converter = HwpxToExcel()
        converter.convert("input.hwpx", "output.xlsx")
    """

    # 엑셀 열 너비 변환 계수
    # 엑셀 열 너비 1 문자 ≈ 8.43 픽셀 (Calibri 11pt 기준)
    # 1 HWPUNIT = 1/100 pt, 1 pt ≈ 1.333 픽셀 (96 DPI)
    # HWPUNIT → 픽셀: hwpunit / 100 * 1.333 = hwpunit / 75
    # 픽셀 → 엑셀 문자: pixel / 8.43
    # 따라서 HWPUNIT → 엑셀 열 너비: hwpunit / 75 / 8.43 ≈ hwpunit / 632
    HWPUNIT_TO_EXCEL_WIDTH = 632

    # 엑셀 행 높이는 포인트 단위
    # HWPUNIT → pt: hwpunit / 100
    HWPUNIT_TO_PT = 100

    def __init__(self):
        self.table_parser = GetTableProperty()
        self.page_parser = GetPageProperty()

    def convert(
        self,
        hwpx_path: Union[str, Path],
        output_path: Optional[Union[str, Path]] = None,
        table_index: int = 0,
        include_cell_info: bool = False,
        hide_para_rows: bool = True
    ) -> Path:
        """
        HWPX 파일의 테이블을 Excel로 변환

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 출력 Excel 경로 (없으면 자동 생성)
            table_index: 변환할 테이블 인덱스 (기본 0)
            include_cell_info: 셀 상세 정보 시트 포함 여부
            hide_para_rows: para_id 행 숨김 여부 (include_cell_info=True일 때)

        Returns:
            생성된 Excel 파일 경로
        """
        hwpx_path = Path(hwpx_path)

        if output_path is None:
            output_path = hwpx_path.with_suffix('.xlsx')
        else:
            output_path = Path(output_path)

        # HWPX에서 데이터 추출
        tables = self.table_parser.from_hwpx(hwpx_path)
        pages = self.page_parser.from_hwpx(hwpx_path)

        if not tables:
            raise ValueError(f"테이블을 찾을 수 없습니다: {hwpx_path}")

        table = tables[table_index] if table_index < len(tables) else tables[0]
        page = pages[0] if pages else None

        # Excel 워크북 생성
        wb = Workbook()
        ws = wb.active
        ws.title = "본문"

        # 1. 페이지 설정 적용
        if page:
            self._apply_page_settings(ws, page)

        # 2. 열 너비 설정
        self._apply_column_widths(ws, table)

        # 3. 행 높이 설정
        self._apply_row_heights(ws, table)

        # 4. 셀 병합 처리
        self._apply_cell_merges(ws, table)

        # 5. 셀 정보 시트 추가 (옵션)
        if include_cell_info:
            try:
                from .cell_info_sheet import add_cell_info_sheet
            except ImportError:
                from cell_info_sheet import add_cell_info_sheet
            # 페이지, 테이블 ID 추출
            page_id = page.section_id if page and page.section_id else "section_0"
            table_id = str(table.id) if table and table.id else f"table_{table_index}"
            add_cell_info_sheet(wb, hwpx_path, "본문_메타", hide_para_rows, page_id, table_id)

        # 저장
        wb.save(output_path)
        return output_path

    def convert_all(
        self,
        hwpx_path: Union[str, Path],
        output_path: Optional[Union[str, Path]] = None,
        include_cell_info: bool = False,
        hide_para_rows: bool = True
    ) -> Path:
        """
        HWPX 파일의 모든 테이블을 Excel 시트로 변환

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 출력 Excel 경로 (없으면 자동 생성)
            include_cell_info: 셀 상세 정보 시트 포함 여부
            hide_para_rows: para_id 행 숨김 여부

        Returns:
            생성된 Excel 파일 경로
        """
        hwpx_path = Path(hwpx_path)

        if output_path is None:
            output_path = hwpx_path.with_suffix('.xlsx')
        else:
            output_path = Path(output_path)

        # HWPX에서 데이터 추출
        tables = self.table_parser.from_hwpx(hwpx_path)
        pages = self.page_parser.from_hwpx(hwpx_path)

        if not tables:
            raise ValueError(f"테이블을 찾을 수 없습니다: {hwpx_path}")

        page = pages[0] if pages else None

        # Excel 워크북 생성
        wb = Workbook()

        # 기본 시트 제거
        default_sheet = wb.active

        for idx, table in enumerate(tables):
            # 시트 생성
            ws = wb.create_sheet(title=f"tbl_{idx}")

            # 1. 페이지 설정 적용
            if page:
                self._apply_page_settings(ws, page)

            # 2. 열 너비 설정
            self._apply_column_widths(ws, table)

            # 3. 행 높이 설정
            self._apply_row_heights(ws, table)

            # 4. 셀 병합 처리
            self._apply_cell_merges(ws, table)

        # 기본 시트 삭제
        wb.remove(default_sheet)

        # 저장
        wb.save(output_path)
        print(f"  {len(tables)}개 테이블 → {len(tables)}개 시트 생성")
        return output_path

    # 용지 크기 매핑 (mm → Excel paperSize 코드)
    PAPER_SIZES = {
        (210, 297): 9,    # A4
        (297, 420): 8,    # A3
        (148, 210): 11,   # A5
        (182, 257): 13,   # B5
        (250, 354): 12,   # B4
        (216, 279): 1,    # Letter (8.5 x 11 in)
        (216, 356): 5,    # Legal (8.5 x 14 in)
    }

    def _apply_page_settings(self, ws: Worksheet, page: PageProperty):
        """페이지 설정 적용 (HWPX 값 동적 매핑)"""
        # 용지 방향
        if page.page_size.orientation == 'landscape':
            ws.page_setup.orientation = 'landscape'
        else:
            ws.page_setup.orientation = 'portrait'

        # 용지 크기 (HWPX에서 가져온 값으로 매핑)
        width_mm = round(Unit.hwpunit_to_mm(page.page_size.width))
        height_mm = round(Unit.hwpunit_to_mm(page.page_size.height))

        # 용지 크기 매핑 (가로/세로 모두 확인)
        paper_size = None
        for (w, h), code in self.PAPER_SIZES.items():
            if (abs(width_mm - w) < 5 and abs(height_mm - h) < 5) or \
               (abs(width_mm - h) < 5 and abs(height_mm - w) < 5):
                paper_size = code
                break

        if paper_size:
            ws.page_setup.paperSize = paper_size
        else:
            # 매핑 안되면 사용자 지정 크기 (인치 단위)
            ws.page_setup.paperWidth = f"{page.page_size.width / Unit.HWPUNIT_PER_INCH:.2f}in"
            ws.page_setup.paperHeight = f"{page.page_size.height / Unit.HWPUNIT_PER_INCH:.2f}in"

        # 여백 (HWPX에서 가져온 값을 인치로 변환)
        ws.page_margins = PageMargins(
            left=page.margin.left / Unit.HWPUNIT_PER_INCH,
            right=page.margin.right / Unit.HWPUNIT_PER_INCH,
            top=page.margin.top / Unit.HWPUNIT_PER_INCH,
            bottom=page.margin.bottom / Unit.HWPUNIT_PER_INCH,
            header=page.margin.header / Unit.HWPUNIT_PER_INCH,
            footer=page.margin.footer / Unit.HWPUNIT_PER_INCH,
        )

    def _apply_column_widths(self, ws: Worksheet, table: TableProperty):
        """열 너비 설정"""
        # 각 열의 너비 계산 (첫 번째 행 기준)
        col_widths = self._get_column_widths(table)

        for col_idx, width in enumerate(col_widths, start=1):
            col_letter = get_column_letter(col_idx)
            # HWPUNIT → 엑셀 열 너비
            excel_width = width / self.HWPUNIT_TO_EXCEL_WIDTH
            # 최소 너비 보장
            excel_width = max(excel_width, 1)
            ws.column_dimensions[col_letter].width = excel_width

    def _apply_row_heights(self, ws: Worksheet, table: TableProperty):
        """행 높이 설정"""
        row_heights = self._get_row_heights(table)

        for row_idx, height in enumerate(row_heights, start=1):
            # HWPUNIT → 포인트
            excel_height = height / self.HWPUNIT_TO_PT
            # 최소 높이 보장
            excel_height = max(excel_height, 10)
            ws.row_dimensions[row_idx].height = excel_height

    def _apply_cell_merges(self, ws: Worksheet, table: TableProperty):
        """셀 병합 처리"""
        for row in table.cells:
            for cell in row:
                if cell.col_span > 1 or cell.row_span > 1:
                    start_row = cell.row_index + 1  # 1-based
                    start_col = cell.col_index + 1
                    end_row = start_row + cell.row_span - 1
                    end_col = start_col + cell.col_span - 1

                    start_cell = f"{get_column_letter(start_col)}{start_row}"
                    end_cell = f"{get_column_letter(end_col)}{end_row}"

                    try:
                        ws.merge_cells(f"{start_cell}:{end_cell}")
                    except ValueError:
                        # 이미 병합된 셀이면 무시
                        pass

    def _get_column_widths(self, table: TableProperty) -> List[int]:
        """각 열의 너비 추출 (셀 너비를 span으로 나눠서 분배)"""
        if table.col_count == 0:
            return []

        col_widths = [0] * table.col_count

        # 모든 셀에서 너비 추출하여 각 열에 분배
        for row in table.cells:
            for cell in row:
                if cell.width and cell.width > 0:
                    # 셀 너비를 span 수로 나눠서 각 열에 분배
                    per_col_width = cell.width // cell.col_span
                    for i in range(cell.col_span):
                        col_idx = cell.col_index + i
                        if col_idx < table.col_count:
                            # 아직 설정 안된 열만 설정
                            if col_widths[col_idx] == 0:
                                col_widths[col_idx] = per_col_width

        # 여전히 0인 열은 테이블 너비 기준 균등 분배
        if table.width:
            default_width = table.width // table.col_count
            for i in range(len(col_widths)):
                if col_widths[i] == 0:
                    col_widths[i] = default_width

        return col_widths

    def _get_row_heights(self, table: TableProperty) -> List[int]:
        """각 행의 높이 추출 (셀 높이를 span으로 나눠서 분배)"""
        if table.row_count == 0:
            return []

        row_heights = [0] * table.row_count

        # 모든 셀에서 높이 추출하여 각 행에 분배
        for row in table.cells:
            for cell in row:
                if cell.height and cell.height > 0:
                    # 셀 높이를 span 수로 나눠서 각 행에 분배
                    per_row_height = cell.height // cell.row_span
                    for i in range(cell.row_span):
                        row_idx = cell.row_index + i
                        if row_idx < table.row_count:
                            # 아직 설정 안된 행만 설정
                            if row_heights[row_idx] == 0:
                                row_heights[row_idx] = per_row_height

        # 여전히 0인 행은 테이블 높이 기준 균등 분배
        if table.height:
            default_height = table.height // table.row_count
            for i in range(len(row_heights)):
                if row_heights[i] == 0:
                    row_heights[i] = default_height

        return row_heights


# ============================================================
# 편의 함수
# ============================================================

def convert_hwpx_to_excel(
    hwpx_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    table_index: int = 0,
    include_cell_info: bool = False,
    hide_para_rows: bool = True
) -> Path:
    """
    HWPX 테이블을 Excel로 변환

    Args:
        hwpx_path: HWPX 파일 경로
        output_path: 출력 Excel 경로
        table_index: 테이블 인덱스
        include_cell_info: 셀 상세 정보 시트 포함 여부
        hide_para_rows: para_id 행 숨김 여부

    Returns:
        생성된 Excel 파일 경로
    """
    converter = HwpxToExcel()
    return converter.convert(
        hwpx_path, output_path, table_index,
        include_cell_info, hide_para_rows
    )


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import json
    import platform

    # OS에 따라 경로 설정
    if platform.system() == "Windows":
        hwpx_path = r"C:\hwp_xml\table.hwpx"
        output_path = r"C:\hwp_xml\excel\output4.xlsx"
    else:
        hwpx_path = "/mnt/c/hwp_xml/table.hwpx"
        output_path = "/mnt/c/hwp_xml/excel/output.xlsx"

    print(f"입력: {hwpx_path}")
    print(f"출력: {output_path}")
    print("=" * 50)

    converter = HwpxToExcel()

    # 테이블 정보 확인
    tables = converter.table_parser.from_hwpx(hwpx_path)
    if tables:
        table = tables[0]
        print(f"테이블: {table.row_count}행 x {table.col_count}열")
        print(f"크기: {table.width} x {table.height} HWPUNIT")
        print(f"      {Unit.hwpunit_to_mm(table.width):.1f}mm x {Unit.hwpunit_to_mm(table.height):.1f}mm")

    # 페이지 정보 확인
    pages = converter.page_parser.from_hwpx(hwpx_path)
    if pages:
        page = pages[0]
        print(f"\n페이지: {page.page_size.orientation}")
        print(f"크기: {Unit.hwpunit_to_mm(page.page_size.width):.1f}mm x {Unit.hwpunit_to_mm(page.page_size.height):.1f}mm")
        print(f"여백: 좌{Unit.hwpunit_to_mm(page.margin.left):.1f}mm 우{Unit.hwpunit_to_mm(page.margin.right):.1f}mm")

    # 변환 실행
    print("\n변환 중...")
    result = converter.convert(hwpx_path, output_path)
    print(f"완료: {result}")
