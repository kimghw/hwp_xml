# -*- coding: utf-8 -*-
"""
HWPX 메타데이터 YAML 추출 모듈

테이블 셀의 list_id, para_id, ctrl(캡션), field_name, 병합 정보를 YAML로 저장
"""

import zipfile
import xml.etree.ElementTree as ET
import json
from pathlib import Path
from typing import Union, List, Dict, Any, Optional
from io import BytesIO
from dataclasses import dataclass, field


@dataclass
class CellMeta:
    """셀 메타 정보"""
    list_id: str = ""
    para_ids: List[str] = field(default_factory=list)
    row: int = 0
    col: int = 0
    row_span: int = 1
    col_span: int = 1
    field_name: str = ""  # tc 태그의 name 속성 (JSON 문자열)
    text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        result = {
            'row': self.row,
            'col': self.col,
        }
        if self.row_span > 1 or self.col_span > 1:
            result['row_span'] = self.row_span
            result['col_span'] = self.col_span
        if self.list_id:
            result['list_id'] = self.list_id
        if self.para_ids:
            result['para_ids'] = self.para_ids
        if self.field_name:
            # JSON 파싱 시도
            try:
                result['field'] = json.loads(self.field_name)
            except:
                result['field_name'] = self.field_name
        if self.text:
            # 텍스트 미리보기 (50자 제한)
            preview = self.text[:50].replace('\n', ' ')
            if len(self.text) > 50:
                preview += '...'
            result['text'] = preview
        return result


@dataclass
class TableMeta:
    """테이블 메타 정보"""
    tbl_idx: int = 0
    table_id: str = ""
    row_count: int = 0
    col_count: int = 0
    caption: str = ""
    cells: List[CellMeta] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        result = {
            'tbl_idx': self.tbl_idx,
            'table_id': self.table_id,
            'size': f'{self.row_count}x{self.col_count}',
        }
        if self.caption:
            result['caption'] = self.caption
        result['cells'] = [c.to_dict() for c in self.cells]
        return result


class ExportMetaYaml:
    """
    HWPX 메타데이터를 YAML로 추출

    사용 예:
        exporter = ExportMetaYaml()
        exporter.export("document.hwpx", "document_meta.yaml")
    """

    def __init__(self):
        pass

    def export(
        self,
        hwpx_path: Union[str, Path],
        output_path: Optional[Union[str, Path]] = None
    ) -> Path:
        """
        HWPX 메타데이터를 YAML로 추출

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 출력 YAML 경로 (없으면 파일명_meta.yaml)

        Returns:
            생성된 YAML 파일 경로
        """
        hwpx_path = Path(hwpx_path)

        if output_path is None:
            output_path = hwpx_path.with_name(hwpx_path.stem + '_meta.yaml')
        else:
            output_path = Path(output_path)

        # 메타데이터 추출
        tables = self._extract_metadata(hwpx_path)

        # YAML 생성 (yaml 모듈 없이 직접 생성)
        yaml_content = self._to_yaml(tables)

        # 파일 저장
        output_path.write_text(yaml_content, encoding='utf-8')

        return output_path

    def _extract_metadata(self, hwpx_path: Path) -> List[TableMeta]:
        """HWPX에서 메타데이터 추출"""
        tables = []
        tbl_idx = 0

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # section 파일 목록
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                xml_content = zf.read(section_file)
                root = ET.parse(BytesIO(xml_content)).getroot()

                # 재귀적으로 테이블 찾기
                tbl_idx = self._find_tables_recursive(root, tables, tbl_idx)

        return tables

    def _find_tables_recursive(
        self, element, tables: List[TableMeta], tbl_idx: int
    ) -> int:
        """재귀적으로 테이블을 찾아 메타 정보 추출"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table_meta = self._parse_table_meta(child, tbl_idx)
                tables.append(table_meta)
                tbl_idx += 1

                # 셀 내부의 중첩 테이블 재귀 탐색
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue
                        tbl_idx = self._find_tables_recursive(tc, tables, tbl_idx)
            else:
                tbl_idx = self._find_tables_recursive(child, tables, tbl_idx)

        return tbl_idx

    def _parse_table_meta(self, tbl_element, tbl_idx: int) -> TableMeta:
        """테이블 메타 정보 파싱"""
        table = TableMeta(
            tbl_idx=tbl_idx,
            table_id=tbl_element.get('id', ''),
            row_count=int(tbl_element.get('rowCnt', 0)),
            col_count=int(tbl_element.get('colCnt', 0)),
        )

        # 캡션 찾기
        for child in tbl_element:
            if child.tag.endswith('}caption'):
                caption_text = self._extract_text(child)
                table.caption = caption_text
                break

        # 셀 파싱
        for tr in tbl_element:
            if not tr.tag.endswith('}tr'):
                continue
            for tc in tr:
                if not tc.tag.endswith('}tc'):
                    continue
                cell_meta = self._parse_cell_meta(tc)
                table.cells.append(cell_meta)

        return table

    def _parse_cell_meta(self, tc_element) -> CellMeta:
        """셀 메타 정보 파싱"""
        cell = CellMeta()
        cell.field_name = tc_element.get('name', '')

        for child in tc_element:
            tag = child.tag.split('}')[-1]

            if tag == 'subList':
                cell.list_id = child.get('id', '')
                # para_id 수집
                for p_elem in child:
                    if p_elem.tag.endswith('}p'):
                        para_id = p_elem.get('id', '')
                        if para_id:
                            cell.para_ids.append(para_id)
                # 텍스트 추출
                cell.text = self._extract_text(child)

            elif tag == 'cellAddr':
                cell.col = int(child.get('colAddr', 0))
                cell.row = int(child.get('rowAddr', 0))

            elif tag == 'cellSpan':
                cell.col_span = int(child.get('colSpan', 1))
                cell.row_span = int(child.get('rowSpan', 1))

        return cell

    def _extract_text(self, element) -> str:
        """요소 내의 모든 텍스트 추출"""
        texts = []
        for t_elem in element.iter():
            if t_elem.tag.endswith('}t') and t_elem.text:
                texts.append(t_elem.text)
        return ''.join(texts)

    def _to_yaml(self, tables: List[TableMeta]) -> str:
        """TableMeta 리스트를 YAML 문자열로 변환 (yaml 모듈 없이)"""
        lines = [
            '# HWPX 메타데이터',
            f'# 테이블 수: {len(tables)}',
            '#',
            '# 셀 형식: [tblIdx, rowAddr, colAddr, rowSpan, colSpan, list_id, para_ids]',
            ''
        ]
        lines.append('tables:')

        for table in tables:
            lines.append(f'  - tbl_idx: {table.tbl_idx}')
            lines.append(f'    table_id: "{table.table_id}"')
            lines.append(f'    size: "{table.row_count}x{table.col_count}"')
            if table.caption:
                caption_escaped = table.caption.replace('"', '\\"').replace('\n', ' ')
                lines.append(f'    caption: "{caption_escaped}"')
            lines.append('    cells:')

            for cell in table.cells:
                # field 정보 파싱
                tbl_idx, row_addr, col_addr, row_span, col_span = 0, 0, 0, 1, 1
                if cell.field_name:
                    try:
                        fd = json.loads(cell.field_name)
                        tbl_idx = fd.get('tblIdx', 0)
                        row_addr = fd.get('rowAddr', 0)
                        col_addr = fd.get('colAddr', 0)
                        row_span = fd.get('rowSpan', 1)
                        col_span = fd.get('colSpan', 1)
                    except:
                        row_addr, col_addr = cell.row, cell.col
                        row_span, col_span = cell.row_span, cell.col_span
                else:
                    tbl_idx = table.tbl_idx
                    row_addr, col_addr = cell.row, cell.col
                    row_span, col_span = cell.row_span, cell.col_span

                # list_id
                list_id = cell.list_id if cell.list_id else ""

                # para_ids를 배열 형식으로
                para_ids_str = ', '.join(cell.para_ids) if cell.para_ids else ""

                # 한 줄로 출력: [tblIdx, row, col, rowSpan, colSpan, list_id, para_ids]
                lines.append(f'      - [{tbl_idx}, {row_addr}, {col_addr}, {row_span}, {col_span}, "{list_id}", [{para_ids_str}]]')

            lines.append('')  # 테이블 간 빈 줄

        return '\n'.join(lines)


# 편의 함수
def export_meta_yaml(
    hwpx_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None
) -> Path:
    """HWPX 메타데이터를 YAML로 추출"""
    exporter = ExportMetaYaml()
    return exporter.export(hwpx_path, output_path)


if __name__ == "__main__":
    import sys

    hwpx_path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/c/hwp_xml/table.hwpx"

    print(f"입력: {hwpx_path}")

    exporter = ExportMetaYaml()
    output_path = exporter.export(hwpx_path)

    print(f"출력: {output_path}")
    print()
    print("=== 내용 미리보기 ===")
    print(output_path.read_text(encoding='utf-8')[:2000])
