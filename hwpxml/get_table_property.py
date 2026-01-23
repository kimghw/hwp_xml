"""
HWPX XML 테이블 속성 및 내용 추출 클래스

아래한글 HWPX 파일에서 테이블의 속성과 내용을 추출합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from io import BytesIO
from pathlib import Path


@dataclass
class CellProperty:
    """테이블 셀 속성"""
    # 위치 정보 (cellAddr)
    row_index: int  # rowAddr
    col_index: int  # colAddr

    # 내용
    text: str

    # 병합 정보 (cellSpan)
    col_span: int = 1
    row_span: int = 1

    # 셀 크기 (cellSz)
    width: Optional[int] = None
    height: Optional[int] = None

    # 셀 여백 (cellMargin)
    margin_left: Optional[int] = None
    margin_right: Optional[int] = None
    margin_top: Optional[int] = None
    margin_bottom: Optional[int] = None

    # 셀 속성 (tc 태그 속성)
    name: str = ""
    header: bool = False
    has_margin: bool = True
    protect: bool = False
    editable: bool = False
    border_fill_id_ref: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """딕셔너리로 변환"""
        return {
            'row_index': self.row_index,
            'col_index': self.col_index,
            'text': self.text,
            'col_span': self.col_span,
            'row_span': self.row_span,
            'width': self.width,
            'height': self.height,
            'margin_left': self.margin_left,
            'margin_right': self.margin_right,
            'margin_top': self.margin_top,
            'margin_bottom': self.margin_bottom,
            'name': self.name,
            'header': self.header,
            'has_margin': self.has_margin,
            'protect': self.protect,
            'editable': self.editable,
            'border_fill_id_ref': self.border_fill_id_ref,
        }


@dataclass
class TableProperty:
    """테이블 속성 정보"""
    id: Optional[str] = None
    row_count: int = 0
    col_count: int = 0
    cell_spacing: int = 0
    border_fill_id_ref: Optional[str] = None
    z_order: Optional[str] = None
    numbering_type: Optional[str] = None
    page_break: Optional[str] = None
    repeat_header: Optional[str] = None
    # 추가 속성
    width: Optional[int] = None
    height: Optional[int] = None

    # 셀 데이터
    cells: List[List[CellProperty]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """딕셔너리로 변환"""
        return {
            'id': self.id,
            'row_count': self.row_count,
            'col_count': self.col_count,
            'cell_spacing': self.cell_spacing,
            'border_fill_id_ref': self.border_fill_id_ref,
            'z_order': self.z_order,
            'numbering_type': self.numbering_type,
            'page_break': self.page_break,
            'repeat_header': self.repeat_header,
            'width': self.width,
            'height': self.height,
        }

    def get_data_as_2d_list(self) -> List[List[str]]:
        """셀 데이터를 2D 리스트로 반환"""
        if not self.cells:
            return []
        return [[cell.text for cell in row] for row in self.cells]

    def to_dataframe(self):
        """pandas DataFrame으로 변환 (pandas가 설치된 경우)"""
        try:
            import pandas as pd
            data = self.get_data_as_2d_list()
            if not data:
                return pd.DataFrame()
            return pd.DataFrame(data)
        except ImportError:
            raise ImportError("pandas가 설치되어 있지 않습니다. pip install pandas")


class GetTableProperty:
    """
    HWPX XML에서 테이블 속성 및 내용을 추출하는 클래스

    사용 예:
        # HWPX 파일에서 추출
        parser = GetTableProperty()
        tables = parser.from_hwpx("document.hwpx")

        # section XML 문자열에서 추출
        tables = parser.from_xml_string(xml_string)

        # section XML 파일에서 추출
        tables = parser.from_xml_file("section0.xml")
    """

    # HWPX 네임스페이스 정의
    NAMESPACES = {
        'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
        'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
        'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
        'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    }

    def __init__(self, custom_namespaces: Optional[Dict[str, str]] = None):
        """
        초기화

        Args:
            custom_namespaces: 사용자 정의 네임스페이스 (기본값 덮어쓰기 가능)
        """
        self.namespaces = self.NAMESPACES.copy()
        if custom_namespaces:
            self.namespaces.update(custom_namespaces)

    def _extract_namespaces_from_xml(self, xml_bytes: bytes) -> Dict[str, str]:
        """XML에서 네임스페이스 동적 추출"""
        namespaces = {}
        try:
            for event, elem in ET.iterparse(BytesIO(xml_bytes), events=['start-ns']):
                prefix, uri = elem
                if prefix:
                    namespaces[prefix] = uri
        except ET.ParseError:
            pass
        return namespaces

    def _get_element_text(self, element: ET.Element) -> str:
        """요소 내의 모든 텍스트를 추출 (hp:t 태그 기준)"""
        texts = []

        # 모든 하위 요소 검색
        for child in element.iter():
            # 태그 이름이 't'로 끝나는 경우 (hp:t, hp10:t 등)
            if child.tag.endswith('}t') and child.text:
                texts.append(child.text)

        return ''.join(texts)

    def _parse_table_element(self, tbl_element: ET.Element, ns: Dict[str, str]) -> TableProperty:
        """
        단일 테이블 요소를 파싱하여 TableProperty 반환

        Args:
            tbl_element: hp:tbl XML 요소
            ns: 네임스페이스 딕셔너리

        Returns:
            TableProperty 객체
        """
        # 테이블 속성 추출
        table_prop = TableProperty(
            id=tbl_element.get('id'),
            row_count=int(tbl_element.get('rowCnt', 0)),
            col_count=int(tbl_element.get('colCnt', 0)),
            cell_spacing=int(tbl_element.get('cellSpacing', 0)),
            border_fill_id_ref=tbl_element.get('borderFillIDRef'),
            z_order=tbl_element.get('zOrder'),
            numbering_type=tbl_element.get('numberingType'),
            page_break=tbl_element.get('pageBreak'),
            repeat_header=tbl_element.get('repeatHeader'),
        )

        # 테이블 크기 정보 (있는 경우)
        # hp:sz 요소에서 width, height 추출
        for sz in tbl_element.iter():
            if sz.tag.endswith('}sz'):
                table_prop.width = int(sz.get('width', 0)) if sz.get('width') else None
                table_prop.height = int(sz.get('height', 0)) if sz.get('height') else None
                break

        # 행(tr) 및 셀(tc) 파싱
        rows_data = []
        row_idx = 0

        for child in tbl_element:
            # tr 요소 찾기 (hp:tr)
            if child.tag.endswith('}tr'):
                row_cells = []
                col_idx = 0

                for tc in child:
                    # tc 요소 찾기 (hp:tc)
                    if tc.tag.endswith('}tc'):
                        # 셀 텍스트 추출
                        cell_text = self._get_element_text(tc)

                        # tc 태그 속성 추출
                        cell_prop = CellProperty(
                            row_index=row_idx,
                            col_index=col_idx,
                            text=cell_text,
                            name=tc.get('name', ''),
                            header=tc.get('header', '0') == '1',
                            has_margin=tc.get('hasMargin', '1') == '1',
                            protect=tc.get('protect', '0') == '1',
                            editable=tc.get('editable', '0') == '1',
                            border_fill_id_ref=tc.get('borderFillIDRef'),
                        )

                        # 하위 요소에서 추가 정보 추출
                        for cell_child in tc:
                            # cellAddr: 셀 좌표
                            if cell_child.tag.endswith('}cellAddr'):
                                cell_prop.col_index = int(cell_child.get('colAddr', col_idx))
                                cell_prop.row_index = int(cell_child.get('rowAddr', row_idx))

                            # cellSpan: 병합 정보
                            elif cell_child.tag.endswith('}cellSpan'):
                                cell_prop.col_span = int(cell_child.get('colSpan', 1))
                                cell_prop.row_span = int(cell_child.get('rowSpan', 1))

                            # cellSz: 셀 크기
                            elif cell_child.tag.endswith('}cellSz'):
                                cell_prop.width = int(cell_child.get('width', 0)) if cell_child.get('width') else None
                                cell_prop.height = int(cell_child.get('height', 0)) if cell_child.get('height') else None

                            # cellMargin: 셀 여백
                            elif cell_child.tag.endswith('}cellMargin'):
                                cell_prop.margin_left = int(cell_child.get('left', 0)) if cell_child.get('left') else None
                                cell_prop.margin_right = int(cell_child.get('right', 0)) if cell_child.get('right') else None
                                cell_prop.margin_top = int(cell_child.get('top', 0)) if cell_child.get('top') else None
                                cell_prop.margin_bottom = int(cell_child.get('bottom', 0)) if cell_child.get('bottom') else None

                        row_cells.append(cell_prop)
                        col_idx += 1

                rows_data.append(row_cells)
                row_idx += 1

        table_prop.cells = rows_data

        # 실제 행/열 수 업데이트 (병합 고려)
        if rows_data:
            table_prop.row_count = len(rows_data)
            # 병합을 고려한 실제 열 개수 계산
            max_col = 0
            for row in rows_data:
                for cell in row:
                    end_col = cell.col_index + cell.col_span
                    if end_col > max_col:
                        max_col = end_col
            table_prop.col_count = max_col

        return table_prop

    def _find_tables_in_element(self, root: ET.Element, ns: Dict[str, str]) -> List[TableProperty]:
        """XML 요소에서 모든 테이블 찾기"""
        tables = []

        # 모든 하위 요소에서 tbl 태그 검색
        for elem in root.iter():
            if elem.tag.endswith('}tbl'):
                table_prop = self._parse_table_element(elem, ns)
                tables.append(table_prop)

        return tables

    def from_xml_string(self, xml_string: Union[str, bytes]) -> List[TableProperty]:
        """
        XML 문자열에서 테이블 추출

        Args:
            xml_string: section XML 문자열 또는 바이트

        Returns:
            TableProperty 객체 리스트
        """
        if isinstance(xml_string, str):
            xml_bytes = xml_string.encode('utf-8')
        else:
            xml_bytes = xml_string

        # 네임스페이스 동적 추출
        ns = self._extract_namespaces_from_xml(xml_bytes)
        ns.update(self.namespaces)

        # XML 파싱
        root = ET.fromstring(xml_bytes)

        return self._find_tables_in_element(root, ns)

    def from_xml_file(self, file_path: Union[str, Path]) -> List[TableProperty]:
        """
        XML 파일에서 테이블 추출

        Args:
            file_path: section XML 파일 경로

        Returns:
            TableProperty 객체 리스트
        """
        file_path = Path(file_path)

        with open(file_path, 'rb') as f:
            xml_bytes = f.read()

        return self.from_xml_string(xml_bytes)

    def from_hwpx(self, hwpx_path: Union[str, Path],
                  section_index: Optional[int] = None) -> List[TableProperty]:
        """
        HWPX 파일에서 테이블 추출

        Args:
            hwpx_path: HWPX 파일 경로
            section_index: 특정 섹션 인덱스 (None이면 모든 섹션)

        Returns:
            TableProperty 객체 리스트
        """
        hwpx_path = Path(hwpx_path)
        all_tables = []

        with zipfile.ZipFile(hwpx_path, 'r') as zipf:
            # 파일 목록 가져오기
            file_list = zipf.namelist()

            # section 파일 찾기
            section_files = sorted([
                f for f in file_list
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            if section_index is not None:
                # 특정 섹션만 처리
                section_file = f'Contents/section{section_index}.xml'
                if section_file in file_list:
                    section_files = [section_file]
                else:
                    raise FileNotFoundError(f"섹션 파일을 찾을 수 없습니다: {section_file}")

            # 각 섹션 파일에서 테이블 추출
            for section_file in section_files:
                xml_bytes = zipf.read(section_file)
                tables = self.from_xml_string(xml_bytes)
                all_tables.extend(tables)

        return all_tables

    def get_table_by_index(self, hwpx_path: Union[str, Path],
                           table_index: int = 0) -> Optional[TableProperty]:
        """
        HWPX 파일에서 특정 인덱스의 테이블 가져오기

        Args:
            hwpx_path: HWPX 파일 경로
            table_index: 테이블 인덱스 (0부터 시작)

        Returns:
            TableProperty 객체 또는 None
        """
        tables = self.from_hwpx(hwpx_path)

        if 0 <= table_index < len(tables):
            return tables[table_index]

        return None

    def get_table_by_id(self, hwpx_path: Union[str, Path],
                        table_id: str) -> Optional[TableProperty]:
        """
        HWPX 파일에서 특정 ID의 테이블 가져오기

        Args:
            hwpx_path: HWPX 파일 경로
            table_id: 테이블 ID

        Returns:
            TableProperty 객체 또는 None
        """
        tables = self.from_hwpx(hwpx_path)

        for table in tables:
            if table.id == table_id:
                return table

        return None


# 편의 함수
def extract_tables_from_hwpx(hwpx_path: Union[str, Path]) -> List[TableProperty]:
    """HWPX 파일에서 모든 테이블 추출 (편의 함수)"""
    parser = GetTableProperty()
    return parser.from_hwpx(hwpx_path)


def extract_table_data_as_list(hwpx_path: Union[str, Path],
                                table_index: int = 0) -> List[List[str]]:
    """HWPX 파일에서 특정 테이블의 데이터를 2D 리스트로 추출 (편의 함수)"""
    parser = GetTableProperty()
    table = parser.get_table_by_index(hwpx_path, table_index)

    if table:
        return table.get_data_as_2d_list()

    return []


# 사용 예제
if __name__ == "__main__":
    # 예제 XML 문자열
    sample_xml = '''<?xml version="1.0" encoding="UTF-8"?>
    <hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
            xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
        <hp:p>
            <hp:run>
                <hp:tbl id="1" zOrder="0" numberingType="TABLE" rowCnt="2" colCnt="3"
                        cellSpacing="0" borderFillIDRef="2">
                    <hp:tr>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>항목1</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>항목2</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>항목3</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                    </hp:tr>
                    <hp:tr>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>값1</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>값2</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                        <hp:tc>
                            <hp:subList>
                                <hp:p><hp:run><hp:t>값3</hp:t></hp:run></hp:p>
                            </hp:subList>
                        </hp:tc>
                    </hp:tr>
                </hp:tbl>
            </hp:run>
        </hp:p>
    </hs:sec>
    '''

    # 테이블 추출 테스트
    parser = GetTableProperty()
    tables = parser.from_xml_string(sample_xml)

    print("=" * 50)
    print("HWPX 테이블 속성 추출 예제")
    print("=" * 50)

    for i, table in enumerate(tables):
        print(f"\n테이블 {i + 1}:")
        print(f"  - ID: {table.id}")
        print(f"  - 행 수: {table.row_count}")
        print(f"  - 열 수: {table.col_count}")
        print(f"  - 셀 간격: {table.cell_spacing}")
        print(f"  - 테두리 참조 ID: {table.border_fill_id_ref}")
        print(f"  - Z-Order: {table.z_order}")
        print(f"  - 번호 매기기 유형: {table.numbering_type}")

        print("\n  테이블 데이터:")
        data = table.get_data_as_2d_list()
        for row in data:
            print(f"    {row}")

        print("\n  속성 딕셔너리:")
        print(f"    {table.to_dict()}")
