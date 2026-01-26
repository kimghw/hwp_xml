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
from openpyxl.worksheet.page import PageMargins, PrintPageSetup
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.worksheet.properties import PageSetupProperties
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from hwpxml.get_table_property import GetTableProperty, TableProperty
from hwpxml.get_page_property import GetPageProperty, PageProperty, Unit
from hwpxml.get_cell_detail import GetCellDetail, CellDetail


@dataclass
class TableHierarchy:
    """테이블 계층 정보"""
    tbl_idx: int
    table_id: str
    parent_tbl_idx: int = -1  # -1이면 최상위 테이블
    parent_row: int = -1
    parent_col: int = -1
    row_count: int = 0
    col_count: int = 0


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

    # HWP 테두리 타입 → openpyxl 스타일 매핑
    BORDER_STYLE_MAP = {
        'NONE': None,
        'SOLID': 'thin',
        'DASH': 'dashed',
        'DOT': 'dotted',
        'DASH_DOT': 'dashDot',
        'DASH_DOT_DOT': 'dashDotDot',
        'DOUBLE': 'double',
        'WAVE': 'thin',  # wave는 thin으로 대체
        'THICK': 'medium',
        'THICK_DOUBLE': 'double',
        'THICK_DASH': 'mediumDashed',
        'THICK_DASH_DOT': 'mediumDashDot',
        'THICK_DASH_DOT_DOT': 'mediumDashDotDot',
    }

    def __init__(self):
        self.table_parser = GetTableProperty()
        self.page_parser = GetPageProperty()
        self.cell_detail_parser = GetCellDetail()
        self.table_hierarchy: List[TableHierarchy] = []

    def _parse_table_hierarchy(self, hwpx_path: Union[str, Path]) -> List[TableHierarchy]:
        """HWPX에서 테이블 계층 구조 파악"""
        hwpx_path = Path(hwpx_path)
        hierarchy = []
        tbl_idx = 0

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # section 파일 목록
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                xml_content = zf.read(section_file)
                root = ET.fromstring(xml_content)

                # 재귀적으로 테이블 파싱
                tbl_idx = self._parse_tables_recursive(
                    root, hierarchy, tbl_idx, parent_tbl_idx=-1, parent_row=-1, parent_col=-1
                )

        return hierarchy

    def _parse_tables_recursive(
        self, element: ET.Element, hierarchy: List[TableHierarchy],
        tbl_idx: int, parent_tbl_idx: int, parent_row: int, parent_col: int
    ) -> int:
        """재귀적으로 테이블 파싱하여 계층 구조 수집"""
        for child in element:
            if child.tag.endswith('}tbl'):
                # 테이블 정보 추출
                table_id = child.get('id', '')
                row_cnt = int(child.get('rowCnt', 0))
                col_cnt = int(child.get('colCnt', 0))

                hierarchy.append(TableHierarchy(
                    tbl_idx=tbl_idx,
                    table_id=table_id,
                    parent_tbl_idx=parent_tbl_idx,
                    parent_row=parent_row,
                    parent_col=parent_col,
                    row_count=row_cnt,
                    col_count=col_cnt
                ))

                current_tbl_idx = tbl_idx
                tbl_idx += 1

                # 하위 셀에서 중첩 테이블 찾기
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue
                        # 셀 위치 추출
                        cell_row, cell_col = 0, 0
                        for cell_child in tc:
                            if cell_child.tag.endswith('}cellAddr'):
                                cell_row = int(cell_child.get('rowAddr', 0))
                                cell_col = int(cell_child.get('colAddr', 0))
                                break
                        # 셀 내 중첩 테이블 재귀 탐색
                        tbl_idx = self._parse_tables_recursive(
                            tc, hierarchy, tbl_idx, current_tbl_idx, cell_row, cell_col
                        )
            else:
                # 다른 요소 내부도 탐색 (tbl이 아닌 경우)
                tbl_idx = self._parse_tables_recursive(
                    child, hierarchy, tbl_idx, parent_tbl_idx, parent_row, parent_col
                )

        return tbl_idx

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
        hide_para_rows: bool = True,
        split_by_para: bool = False
    ) -> Path:
        """
        HWPX 파일의 모든 테이블을 Excel 시트로 변환
        - 최상위 테이블: 개별 시트로 생성
        - 중첩 테이블: tbl_sub 시트에 정리 + 부모 셀에 하이퍼링크/메모

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 출력 Excel 경로 (없으면 자동 생성)
            include_cell_info: 셀 상세 정보 시트 포함 여부
            hide_para_rows: para_id 행 숨김 여부
            split_by_para: 셀 내 문단별로 행 분할 (True면 문단 수만큼 행 생성)

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
        # 테이블별로 그룹화된 셀 디테일 가져오기
        table_cell_details = self.cell_detail_parser.from_hwpx_by_table(hwpx_path)

        if not tables:
            raise ValueError(f"테이블을 찾을 수 없습니다: {hwpx_path}")

        # 테이블 계층 구조 파악
        self.table_hierarchy = self._parse_table_hierarchy(hwpx_path)

        # 최상위 테이블과 중첩 테이블 분리
        top_level_tables = [h for h in self.table_hierarchy if h.parent_tbl_idx == -1]
        nested_tables = [h for h in self.table_hierarchy if h.parent_tbl_idx != -1]

        page = pages[0] if pages else None

        # Excel 워크북 생성
        wb = Workbook()
        default_sheet = wb.active

        # 1. 모든 테이블 시트 생성 (최상위 + 중첩 모두)
        # 각 테이블의 row_offset_map 저장 (split_by_para용)
        table_row_offset_maps = {}

        for idx, table in enumerate(tables):
            ws = wb.create_sheet(title=f"tbl_{idx}")

            if page:
                self._apply_page_settings(ws, page)

            self._apply_column_widths(ws, table)

            # split_by_para 옵션에 따라 처리
            if split_by_para and idx < len(table_cell_details):
                # 문단별 행 분할 (row_offset_map 반환)
                row_offset_map = self._apply_table_with_para_split(ws, table, table_cell_details[idx])
                table_row_offset_maps[idx] = row_offset_map
            else:
                # 기존 방식
                self._apply_row_heights(ws, table)
                self._apply_cell_merges(ws, table)

                # 셀 스타일 적용
                if idx < len(table_cell_details):
                    self._apply_cell_styles(ws, table_cell_details[idx], idx)

        # 2. 부모 테이블 셀에 중첩 테이블 하이퍼링크/메모 추가
        # 같은 셀에 여러 nested 테이블이 있을 수 있으므로 그룹화
        nested_by_cell = {}  # (parent_idx, row, col) -> [nested1, nested2, ...]
        for nested in nested_tables:
            parent_idx = nested.parent_tbl_idx
            if parent_idx < 0 or parent_idx >= len(tables):
                continue

            orig_row = nested.parent_row
            col = nested.parent_col + 1  # 1-based

            # split_by_para일 때 row_offset_map 사용
            if parent_idx in table_row_offset_maps:
                row_offset_map = table_row_offset_maps[parent_idx]
                row = row_offset_map.get(orig_row, orig_row) + 1  # 1-based
            else:
                row = orig_row + 1  # 1-based

            key = (parent_idx, row, col)
            if key not in nested_by_cell:
                nested_by_cell[key] = []
            nested_by_cell[key].append(nested)

        # 그룹화된 nested 테이블 처리
        for (parent_idx, row, col), nested_list in nested_by_cell.items():
            parent_sheet_name = f"tbl_{parent_idx}"
            if parent_sheet_name not in wb.sheetnames:
                continue

            parent_ws = wb[parent_sheet_name]

            if row > 0 and col > 0:
                cell = parent_ws.cell(row=row, column=col)

                # 여러 nested 테이블이 있을 경우 표시
                if len(nested_list) == 1:
                    nested = nested_list[0]
                    cell.hyperlink = f"#tbl_{nested.tbl_idx}!A1"
                    cell.value = f"→tbl_{nested.tbl_idx}"
                    comment_text = (
                        f"중첩 테이블: tbl_{nested.tbl_idx}\n"
                        f"크기: {nested.row_count}행 x {nested.col_count}열"
                    )
                else:
                    # 여러 개: 첫 번째 테이블로 링크, 나머지는 메모에 표시
                    tbl_names = [f"tbl_{n.tbl_idx}" for n in nested_list]
                    cell.hyperlink = f"#tbl_{nested_list[0].tbl_idx}!A1"
                    cell.value = f"→{','.join(tbl_names)}"
                    comment_lines = []
                    for n in nested_list:
                        comment_lines.append(f"tbl_{n.tbl_idx}: {n.row_count}행 x {n.col_count}열")
                    comment_text = "중첩 테이블:\n" + "\n".join(comment_lines)

                cell.style = "Hyperlink"
                cell.comment = Comment(comment_text, "System")

        # 3. tbl_sub 시트 생성 (중첩 테이블이 있는 경우)
        if nested_tables:
            ws_sub = wb.create_sheet(title="tbl_sub")

            # 헤더
            headers = ["tbl_idx", "table_id", "parent_tbl", "parent_row", "parent_col", "rows", "cols", "link"]
            for col_idx, header in enumerate(headers, 1):
                ws_sub.cell(row=1, column=col_idx, value=header)

            # 데이터
            for row_idx, nested in enumerate(nested_tables, 2):
                ws_sub.cell(row=row_idx, column=1, value=nested.tbl_idx)
                ws_sub.cell(row=row_idx, column=2, value=nested.table_id)
                ws_sub.cell(row=row_idx, column=3, value=f"tbl_{nested.parent_tbl_idx}")
                ws_sub.cell(row=row_idx, column=4, value=nested.parent_row)
                ws_sub.cell(row=row_idx, column=5, value=nested.parent_col)
                ws_sub.cell(row=row_idx, column=6, value=nested.row_count)
                ws_sub.cell(row=row_idx, column=7, value=nested.col_count)

                # 링크 셀
                link_cell = ws_sub.cell(row=row_idx, column=8)
                link_cell.hyperlink = f"#tbl_{nested.tbl_idx}!A1"
                link_cell.value = f"→tbl_{nested.tbl_idx}"
                link_cell.style = "Hyperlink"

            # 열 너비 조정
            ws_sub.column_dimensions['A'].width = 8
            ws_sub.column_dimensions['B'].width = 15
            ws_sub.column_dimensions['C'].width = 12
            ws_sub.column_dimensions['D'].width = 10
            ws_sub.column_dimensions['E'].width = 10
            ws_sub.column_dimensions['F'].width = 6
            ws_sub.column_dimensions['G'].width = 6
            ws_sub.column_dimensions['H'].width = 12

        # 기본 시트 삭제
        wb.remove(default_sheet)

        # 저장
        wb.save(output_path)

        top_count = len(top_level_tables)
        nested_count = len(nested_tables)
        print(f"  {len(tables)}개 테이블 (최상위: {top_count}, 중첩: {nested_count})")
        return output_path

    def _apply_cell_styles_with_offset(
        self, ws: Worksheet, cell_details: List[CellDetail], row_offset: int
    ):
        """셀 스타일 적용 (행 오프셋 적용)"""
        for cell_detail in cell_details:
            row = cell_detail.row + row_offset
            col = cell_detail.col + 1

            try:
                excel_cell = ws.cell(row=row, column=col)
            except:
                continue

            # 병합된 셀의 마스터가 아닌 경우 스킵
            if hasattr(excel_cell, 'is_merged') or type(excel_cell).__name__ == 'MergedCell':
                cell_coord = f"{get_column_letter(col)}{row}"
                is_master = True
                for merged_range in ws.merged_cells.ranges:
                    if cell_coord in merged_range and cell_coord != str(merged_range).split(':')[0]:
                        is_master = False
                        break
                if not is_master:
                    continue

            # 텍스트 설정
            try:
                if not excel_cell.hyperlink and cell_detail.text:
                    excel_cell.value = cell_detail.text
            except AttributeError:
                pass

            # 테두리 설정
            border = Border(
                left=self._get_border_side(cell_detail.border.left),
                right=self._get_border_side(cell_detail.border.right),
                top=self._get_border_side(cell_detail.border.top),
                bottom=self._get_border_side(cell_detail.border.bottom),
            )
            excel_cell.border = border

            # 배경색 설정
            bg_color = self._hwp_color_to_rgb(cell_detail.border.bg_color)
            if bg_color and bg_color != 'FFFFFF':
                excel_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

            # 폰트 설정
            font_color = self._hwp_color_to_rgb(cell_detail.font.color)
            excel_cell.font = Font(
                name=cell_detail.font.name if cell_detail.font.name else None,
                size=cell_detail.font.size_pt() if cell_detail.font.size > 0 else None,
                bold=cell_detail.font.bold,
                italic=cell_detail.font.italic,
                underline='single' if cell_detail.font.underline else None,
                strike=cell_detail.font.strikeout,
                color=font_color if font_color else None,
            )

            # 정렬 설정
            h_align_map = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}
            v_align_map = {'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom', 'BASELINE': 'center'}

            h_align = 'left'
            v_align = 'center'
            if cell_detail.paragraphs:
                h_align = h_align_map.get(cell_detail.paragraphs[0].align_h, 'left')
                v_align = v_align_map.get(cell_detail.paragraphs[0].align_v, 'center')

            excel_cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=True)

            # 병합된 셀의 나머지 영역에도 테두리 적용
            if cell_detail.row_span > 1 or cell_detail.col_span > 1:
                for r in range(cell_detail.row_span):
                    for c in range(cell_detail.col_span):
                        if r == 0 and c == 0:
                            continue
                        try:
                            merged_cell = ws.cell(row=row + r, column=col + c)
                            merged_border = Border(
                                left=self._get_border_side(cell_detail.border.left) if c == 0 else Side(),
                                right=self._get_border_side(cell_detail.border.right) if c == cell_detail.col_span - 1 else Side(),
                                top=self._get_border_side(cell_detail.border.top) if r == 0 else Side(),
                                bottom=self._get_border_side(cell_detail.border.bottom) if r == cell_detail.row_span - 1 else Side(),
                            )
                            merged_cell.border = merged_border
                        except:
                            pass

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

    def _apply_table_with_para_split(
        self, ws: Worksheet, table: TableProperty, cell_details: List[CellDetail]
    ) -> dict:
        """
        문단별 행 분할하여 테이블 생성

        원본 테이블의 각 행을 문단 수에 따라 여러 행으로 분할합니다.
        - 셀에 문단이 3개 있으면 해당 행이 3행으로 분할
        - 같은 행의 다른 셀(문단 1개)은 분할된 행에 맞춰 세로 병합
        - 원본 셀 높이를 분할된 행에 균등 분배 (높이 유지)

        Returns:
            row_offset_map: 원본 행 인덱스 → 새 행 오프셋 매핑
        """
        # 셀 디테일을 (row, col) → CellDetail 매핑으로 변환
        cell_map = {}
        for cd in cell_details:
            cell_map[(cd.row, cd.col)] = cd

        # 1. 각 행별 최대 문단 수 및 문단별 높이 계산
        row_max_paras = {}
        row_para_heights = {}  # row_idx -> [높이1, 높이2, ...] (각 문단별 최대 높이)

        for row_idx in range(table.row_count):
            max_paras = 1
            para_heights = []

            for col_idx in range(table.col_count):
                if (row_idx, col_idx) in cell_map:
                    cd = cell_map[(row_idx, col_idx)]
                    para_count = len(cd.paragraphs) if cd.paragraphs else 1
                    max_paras = max(max_paras, para_count)

                    # 각 문단별 높이 수집
                    for p_idx, para in enumerate(cd.paragraphs or []):
                        while len(para_heights) <= p_idx:
                            para_heights.append(0)
                        if para.height > para_heights[p_idx]:
                            para_heights[p_idx] = para.height

            row_max_paras[row_idx] = max_paras
            row_para_heights[row_idx] = para_heights if para_heights else [0]

        # 2. 원본 행 → 새 행 오프셋 매핑 계산
        row_offset_map = {}
        cumulative = 0
        for row_idx in range(table.row_count):
            row_offset_map[row_idx] = cumulative
            cumulative += row_max_paras.get(row_idx, 1)

        # 3. 분할된 행 높이 설정 (각 문단의 실제 높이 사용)
        for row_idx in range(table.row_count):
            para_count = row_max_paras.get(row_idx, 1)
            para_heights = row_para_heights.get(row_idx, [])
            new_start_row = row_offset_map[row_idx] + 1  # 1-based

            for p_idx in range(para_count):
                # 해당 문단의 높이 사용 (없으면 0)
                height = para_heights[p_idx] if p_idx < len(para_heights) else 0
                if height > 0:
                    height_pt = height / self.HWPUNIT_TO_PT
                    ws.row_dimensions[new_start_row + p_idx].height = height_pt

        # 4. 셀 배치
        for cd in cell_details:
            orig_row = cd.row
            orig_col = cd.col
            new_start_row = row_offset_map[orig_row] + 1  # 1-based
            new_col = orig_col + 1  # 1-based

            para_count = len(cd.paragraphs) if cd.paragraphs else 1

            # 이 셀이 차지하는 새 행 수 계산 (row_span 고려)
            new_row_span = 0
            for r in range(cd.row_span):
                new_row_span += row_max_paras.get(orig_row + r, 1)

            # 문단별로 값 설정
            for para_idx in range(para_count):
                new_row = new_start_row + para_idx

                try:
                    excel_cell = ws.cell(row=new_row, column=new_col)

                    if cd.paragraphs and para_idx < len(cd.paragraphs):
                        para = cd.paragraphs[para_idx]
                        excel_cell.value = para.text
                    elif para_idx == 0:
                        excel_cell.value = cd.text

                    # 스타일 적용
                    self._apply_cell_style_single(excel_cell, cd, para_idx, para_count)

                except Exception:
                    pass

            # 병합 처리
            merge_end_row = new_start_row + new_row_span - 1
            merge_start_col = new_col
            merge_end_col = new_col + cd.col_span - 1

            needs_row_merge = new_row_span > para_count  # 세로 병합 필요
            needs_col_merge = cd.col_span > 1  # 가로 병합 필요

            if needs_row_merge or needs_col_merge:
                if para_count <= 1:
                    # 문단이 0-1개: 전체 범위 병합
                    if new_row_span > 1 or cd.col_span > 1:
                        try:
                            ws.merge_cells(
                                start_row=new_start_row,
                                start_column=merge_start_col,
                                end_row=merge_end_row,
                                end_column=merge_end_col
                            )
                        except ValueError:
                            pass
                else:
                    # 문단이 2개 이상
                    if needs_row_merge:
                        # 세로 병합 필요: 마지막 문단 행부터 끝까지 병합 (가로도 포함)
                        # 먼저 마지막 문단 이전 행들은 가로만 병합
                        if needs_col_merge:
                            for p_idx in range(para_count - 1):  # 마지막 제외
                                try:
                                    ws.merge_cells(
                                        start_row=new_start_row + p_idx,
                                        start_column=merge_start_col,
                                        end_row=new_start_row + p_idx,
                                        end_column=merge_end_col
                                    )
                                except ValueError:
                                    pass
                        # 마지막 문단 행부터 끝까지 세로+가로 병합
                        try:
                            ws.merge_cells(
                                start_row=new_start_row + para_count - 1,
                                start_column=merge_start_col,
                                end_row=merge_end_row,
                                end_column=merge_end_col
                            )
                        except ValueError:
                            pass
                    else:
                        # 세로 병합 불필요: 각 문단 행 가로만 병합
                        if needs_col_merge:
                            for p_idx in range(para_count):
                                try:
                                    ws.merge_cells(
                                        start_row=new_start_row + p_idx,
                                        start_column=merge_start_col,
                                        end_row=new_start_row + p_idx,
                                        end_column=merge_end_col
                                    )
                                except ValueError:
                                    pass

        return row_offset_map

    def _apply_cell_style_single(self, excel_cell, cd: CellDetail, para_idx: int, total_paras: int):
        """단일 셀에 스타일 적용"""
        # 테두리
        border = Border(
            left=self._get_border_side(cd.border.left),
            right=self._get_border_side(cd.border.right),
            top=self._get_border_side(cd.border.top) if para_idx == 0 else Side(),
            bottom=self._get_border_side(cd.border.bottom) if para_idx == total_paras - 1 else Side(),
        )
        excel_cell.border = border

        # 배경색
        bg_color = self._hwp_color_to_rgb(cd.border.bg_color)
        if bg_color and bg_color != 'FFFFFF':
            excel_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

        # 폰트 (문단별 폰트 사용, 없으면 셀 기본 폰트)
        para_font = None
        if para_idx < len(cd.paragraphs) and cd.paragraphs[para_idx].font:
            para_font = cd.paragraphs[para_idx].font
        else:
            para_font = cd.font

        font_color = self._hwp_color_to_rgb(para_font.color)
        excel_cell.font = Font(
            name=para_font.name if para_font.name else None,
            size=para_font.size_pt() if para_font.size > 0 else None,
            bold=para_font.bold,
            italic=para_font.italic,
            underline='single' if para_font.underline else None,
            strike=para_font.strikeout,
            color=font_color if font_color else None,
        )

        # 정렬
        h_align_map = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}
        v_align_map = {'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom', 'BASELINE': 'center'}
        h_align = 'left'
        v_align = 'center'
        if cd.paragraphs:
            h_align = h_align_map.get(cd.paragraphs[0].align_h, 'left')
            v_align = v_align_map.get(cd.paragraphs[0].align_v, 'center')
        excel_cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=True)

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

    def _group_cells_by_table(self, tables: List[TableProperty], all_cells: List[CellDetail]) -> List[List[CellDetail]]:
        """셀 디테일을 테이블별로 그룹화"""
        # 각 테이블의 셀 개수 계산
        result = []
        cell_idx = 0

        for table in tables:
            table_cells = []
            expected_cells = sum(len(row) for row in table.cells)

            for _ in range(expected_cells):
                if cell_idx < len(all_cells):
                    table_cells.append(all_cells[cell_idx])
                    cell_idx += 1

            result.append(table_cells)

        return result

    def _hwp_color_to_rgb(self, color_str: str) -> str:
        """HWP 색상 문자열을 RGB hex로 변환"""
        if not color_str:
            return None
        # #RRGGBB 형식이면 그대로 반환 (# 제거)
        if color_str.startswith('#'):
            return color_str[1:].upper()
        # RGB(r,g,b) 형식 처리
        if color_str.startswith('RGB('):
            try:
                rgb = color_str[4:-1].split(',')
                r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
                return f"{r:02X}{g:02X}{b:02X}"
            except:
                return None
        # 숫자형 색상 (BGR 또는 RGB)
        try:
            val = int(color_str)
            r = val & 0xFF
            g = (val >> 8) & 0xFF
            b = (val >> 16) & 0xFF
            return f"{r:02X}{g:02X}{b:02X}"
        except:
            return None

    def _get_border_side(self, border_type: str) -> Side:
        """HWP 테두리 타입을 openpyxl Side로 변환"""
        style = self.BORDER_STYLE_MAP.get(border_type, None)
        if style:
            return Side(style=style, color='000000')
        return Side(style=None)

    def _apply_cell_styles(self, ws: Worksheet, cell_details: List[CellDetail], table_idx: int = 0):
        """셀 스타일 적용 (테두리, 배경색, 텍스트, 폰트)"""
        # cell_details에서 해당 테이블의 셀만 필터링하기 위해 순서대로 매핑
        # cell_details는 모든 테이블의 셀이 순서대로 들어있음
        # 테이블 인덱스별로 분리 필요

        for cell_detail in cell_details:
            row = cell_detail.row + 1  # 1-based
            col = cell_detail.col + 1

            try:
                excel_cell = ws.cell(row=row, column=col)
            except:
                continue

            # 병합된 셀의 마스터가 아닌 경우 스킵
            # openpyxl에서 MergedCell은 읽기 전용
            if hasattr(excel_cell, 'is_merged') or type(excel_cell).__name__ == 'MergedCell':
                # 병합 영역의 첫 번째 셀인지 확인
                cell_coord = f"{get_column_letter(col)}{row}"
                is_master = True
                for merged_range in ws.merged_cells.ranges:
                    if cell_coord in merged_range and cell_coord != str(merged_range).split(':')[0]:
                        is_master = False
                        break
                if not is_master:
                    continue

            # 1. 텍스트 설정 (하이퍼링크가 아닌 경우에만)
            try:
                if not excel_cell.hyperlink and cell_detail.text:
                    excel_cell.value = cell_detail.text
            except AttributeError:
                # MergedCell인 경우 무시
                pass

            # 2. 테두리 설정
            border = Border(
                left=self._get_border_side(cell_detail.border.left),
                right=self._get_border_side(cell_detail.border.right),
                top=self._get_border_side(cell_detail.border.top),
                bottom=self._get_border_side(cell_detail.border.bottom),
            )
            excel_cell.border = border

            # 3. 배경색 설정
            bg_color = self._hwp_color_to_rgb(cell_detail.border.bg_color)
            if bg_color and bg_color != 'FFFFFF':
                excel_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')

            # 4. 폰트 설정
            font_color = self._hwp_color_to_rgb(cell_detail.font.color)
            excel_cell.font = Font(
                name=cell_detail.font.name if cell_detail.font.name else None,
                size=cell_detail.font.size_pt() if cell_detail.font.size > 0 else None,
                bold=cell_detail.font.bold,
                italic=cell_detail.font.italic,
                underline='single' if cell_detail.font.underline else None,
                strike=cell_detail.font.strikeout,
                color=font_color if font_color else None,
            )

            # 5. 정렬 설정
            h_align_map = {'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'justify'}
            v_align_map = {'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom', 'BASELINE': 'center'}

            h_align = 'left'
            v_align = 'center'
            if cell_detail.paragraphs:
                h_align = h_align_map.get(cell_detail.paragraphs[0].align_h, 'left')
                v_align = v_align_map.get(cell_detail.paragraphs[0].align_v, 'center')

            excel_cell.alignment = Alignment(horizontal=h_align, vertical=v_align, wrap_text=True)

            # 6. 병합된 셀의 나머지 영역에도 테두리 적용
            if cell_detail.row_span > 1 or cell_detail.col_span > 1:
                for r in range(cell_detail.row_span):
                    for c in range(cell_detail.col_span):
                        if r == 0 and c == 0:
                            continue  # 첫 셀은 이미 처리됨
                        try:
                            merged_cell = ws.cell(row=row + r, column=col + c)
                            # 병합된 영역의 테두리만 설정
                            merged_border = Border(
                                left=self._get_border_side(cell_detail.border.left) if c == 0 else Side(),
                                right=self._get_border_side(cell_detail.border.right) if c == cell_detail.col_span - 1 else Side(),
                                top=self._get_border_side(cell_detail.border.top) if r == 0 else Side(),
                                bottom=self._get_border_side(cell_detail.border.bottom) if r == cell_detail.row_span - 1 else Side(),
                            )
                            merged_cell.border = merged_border
                        except:
                            pass


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
