# -*- coding: utf-8 -*-
"""
테이블 행 데이터 추출 모듈

테이블 XML 요소에서 행별로 필드명-값 데이터를 추출합니다.

주요 기능:
- 셀에서 필드명(nc.name)과 텍스트 추출
- gstub/stub rowspan 범위 전파
"""

from typing import Dict, List, Set, Optional


class RowExtractor:
    """테이블 행 데이터 추출기"""

    def extract_raw(self, tbl_elem) -> Dict[int, Dict[str, str]]:
        """
        테이블에서 원시 행 데이터 추출 (필터링 없음)

        Args:
            tbl_elem: 테이블 XML 요소

        Returns:
            행별 데이터 딕셔너리 {row_idx: {field_name: text}}
        """
        row_data: Dict[int, Dict[str, str]] = {}
        gstub_cells = []

        for tc in tbl_elem.iter():
            if not tc.tag.endswith('}tc'):
                continue

            field_name = tc.get('name', '')
            if not field_name:
                continue

            # 셀 주소와 span 정보 추출
            row_idx = 0
            row_span = 1
            for child in tc:
                if child.tag.endswith('}cellAddr'):
                    row_idx = int(child.get('rowAddr', 0))
                elif child.tag.endswith('}cellSpan'):
                    row_span = int(child.get('rowSpan', 1))

            # 텍스트 추출
            text = self._extract_text(tc)

            # gstub/stub 셀은 rowspan 정보와 함께 저장
            if field_name.startswith('gstub_') or field_name.startswith('stub_'):
                end_row = row_idx + row_span - 1
                gstub_cells.append((row_idx, end_row, field_name, text))

            # 행 데이터에 저장
            if row_idx not in row_data:
                row_data[row_idx] = {}
            row_data[row_idx][field_name] = text

        # gstub/stub 값을 rowspan 범위의 모든 행에 전파
        self._propagate_gstub_values(row_data, gstub_cells)

        return row_data

    def _extract_text(self, tc) -> str:
        """셀에서 텍스트 추출 (여러 문단은 줄바꿈으로 구분)"""
        paragraphs_text = []
        for sublist in tc:
            if sublist.tag.endswith('}subList'):
                for p in sublist:
                    if p.tag.endswith('}p'):
                        p_text = ""
                        for run in p:
                            if run.tag.endswith('}run'):
                                for t in run:
                                    if t.tag.endswith('}t') and t.text:
                                        p_text += t.text
                        paragraphs_text.append(p_text)
        return '\n'.join(paragraphs_text)

    def _propagate_gstub_values(
        self,
        row_data: Dict[int, Dict[str, str]],
        gstub_cells: List
    ):
        """gstub/stub 값을 해당 rowspan 범위의 모든 행에 전파"""
        for start_row, end_row, field_name, text in gstub_cells:
            for r in range(start_row, end_row + 1):
                if r not in row_data:
                    row_data[r] = {}
                if field_name not in row_data[r]:
                    row_data[r][field_name] = text


def extract_table_rows(tbl_elem) -> Dict[int, Dict[str, str]]:
    """
    테이블에서 원시 행 데이터 추출 (편의 함수)

    Args:
        tbl_elem: 테이블 XML 요소

    Returns:
        행별 데이터 딕셔너리 {row_idx: {field_name: text}}
    """
    extractor = RowExtractor()
    return extractor.extract_raw(tbl_elem)
