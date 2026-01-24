# -*- coding: utf-8 -*-
"""
HWPX 문서 내용 추출 모듈

본문 문단과 테이블의 순서를 유지하면서 추출합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from pathlib import Path
from io import BytesIO


@dataclass
class ParagraphContent:
    """본문 문단 정보"""
    type: str = "paragraph"  # paragraph, table
    text: str = ""
    para_id: str = ""
    style_id: str = ""
    order: int = 0  # 문서 내 순서


@dataclass
class TableMarker:
    """테이블 위치 마커"""
    type: str = "table"
    table_idx: int = 0
    table_id: str = ""
    row_count: int = 0
    col_count: int = 0
    order: int = 0


@dataclass
class DocumentContent:
    """문서 내용 (문단 + 테이블 순서 유지)"""
    items: List[Union[ParagraphContent, TableMarker]] = field(default_factory=list)


class GetDocumentContent:
    """HWPX 문서 내용 추출"""

    def __init__(self):
        self._styles = {}

    def from_hwpx(self, hwpx_path: Union[str, Path]) -> DocumentContent:
        """HWPX 파일에서 문서 내용 추출 (문단 + 테이블 순서 유지)"""
        hwpx_path = Path(hwpx_path)
        content = DocumentContent()
        self._order = 0
        self._table_idx = 0

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # section 파일 목록
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                xml_content = zf.read(section_file)
                root = ET.fromstring(xml_content)

                # 재귀적으로 문단과 테이블 탐색
                self._parse_element_recursive(root, content)

        return content

    def _parse_element_recursive(self, element, content: DocumentContent):
        """요소를 재귀적으로 탐색하여 문단과 테이블 추출"""
        for child in element:
            tag = child.tag.split('}')[-1]

            if tag == 'p':
                # 문단 텍스트 추출 (테이블 내부 텍스트 제외)
                para = self._parse_paragraph_text_only(child)
                if para.text.strip():
                    para.order = self._order
                    self._order += 1
                    content.items.append(para)

                # 문단 내부에 테이블이 있을 수 있음 (run 안에)
                self._parse_element_recursive(child, content)

            elif tag == 'tbl':
                # 테이블 마커
                marker = TableMarker(
                    table_idx=self._table_idx,
                    table_id=child.get('id', ''),
                    row_count=int(child.get('rowCnt', 0)),
                    col_count=int(child.get('colCnt', 0)),
                    order=self._order
                )
                self._order += 1
                self._table_idx += 1
                content.items.append(marker)

                # 테이블 내부의 중첩 테이블 탐색 (셀 안의 테이블)
                self._parse_tables_in_cells(child, content)

            elif tag in ('run', 'subList', 'ctrl'):
                # 이 요소들 안에 테이블이 있을 수 있음
                self._parse_element_recursive(child, content)

    def _parse_tables_in_cells(self, tbl_elem, content: DocumentContent):
        """테이블 셀 내부의 중첩 테이블 탐색"""
        for tr in tbl_elem:
            if not tr.tag.endswith('}tr'):
                continue
            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue
                # 셀 내부의 subList에서 테이블 찾기
                for child in tc:
                    self._parse_element_recursive(child, content)

    def _parse_paragraph_text_only(self, p_elem) -> ParagraphContent:
        """문단 요소에서 직접 텍스트만 추출 (테이블 내용 제외)"""
        para = ParagraphContent()
        para.para_id = p_elem.get('id', '')
        para.style_id = p_elem.get('paraPrIDRef', '')

        # 문단 내 직접 텍스트 추출 (tbl 요소 안의 텍스트 제외)
        texts = []
        self._extract_text_excluding_tables(p_elem, texts)
        para.text = ''.join(texts)
        return para

    def _extract_text_excluding_tables(self, elem, texts: List[str]):
        """테이블 내부 텍스트를 제외하고 텍스트 추출"""
        for child in elem:
            tag = child.tag.split('}')[-1]
            if tag == 'tbl':
                # 테이블 내용은 건너뜀
                continue
            elif tag == 't':
                if child.text:
                    texts.append(child.text)
            else:
                self._extract_text_excluding_tables(child, texts)

    def _parse_paragraph(self, p_elem) -> ParagraphContent:
        """문단 요소 파싱"""
        para = ParagraphContent()
        para.para_id = p_elem.get('id', '')
        para.style_id = p_elem.get('paraPrIDRef', '')

        # 문단 내 텍스트 추출
        texts = []
        for elem in p_elem.iter():
            if elem.tag.endswith('}t'):
                if elem.text:
                    texts.append(elem.text)

        para.text = ''.join(texts)
        return para

    def _count_nested_tables(self, tbl_elem, table_idx: int) -> int:
        """중첩 테이블 카운트 (재귀)"""
        for tr in tbl_elem:
            if not tr.tag.endswith('}tr'):
                continue
            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue
                for child in tc:
                    if child.tag.endswith('}tbl'):
                        table_idx += 1
                        table_idx = self._count_nested_tables(child, table_idx)
        return table_idx


# 테스트
if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        hwpx_path = sys.argv[1]
    else:
        hwpx_path = "/mnt/c/hwp_xml/test_sample.hwpx"

    parser = GetDocumentContent()
    try:
        content = parser.from_hwpx(hwpx_path)
        print(f"문서 항목: {len(content.items)}개")

        for i, item in enumerate(content.items[:20]):
            if item.type == "paragraph":
                text = item.text[:50] + "..." if len(item.text) > 50 else item.text
                print(f"  {i}: [문단] {text}")
            else:
                print(f"  {i}: [테이블 {item.table_idx}] {item.row_count}x{item.col_count}")

        if len(content.items) > 20:
            print(f"  ... 외 {len(content.items) - 20}개")
    except Exception as e:
        print(f"오류: {e}")
