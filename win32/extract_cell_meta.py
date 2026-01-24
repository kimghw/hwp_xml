# -*- coding: utf-8 -*-
"""
HWP 테이블 셀 메타데이터 추출

각 셀의 list_id, para_id, field_name(tc.name 속성) 추출
HWP COM API + HWPX XML 결합
"""

import sys
import os
import json
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from io import BytesIO

# 프로젝트 루트 경로 설정
_project_root = Path(__file__).parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

try:
    from config import WIN32HWP_DIR
    if str(WIN32HWP_DIR) not in sys.path:
        sys.path.insert(0, str(WIN32HWP_DIR))
except ImportError:
    win32hwp_dir = os.environ.get('WIN32HWP_DIR', r'C:\win32hwp')
    if win32hwp_dir not in sys.path:
        sys.path.insert(0, win32hwp_dir)

from cursor import get_hwp_instance
import win32com.client as win32


def get_or_create_hwp():
    """실행 중인 한글에 연결하거나, 없으면 새로 실행"""
    hwp = get_hwp_instance()
    if hwp:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except:
            pass
        return hwp
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp


class ExtractCellMeta:
    """HWP 테이블 셀 메타데이터 추출"""

    def __init__(self, hwp=None):
        self.hwp = hwp or get_or_create_hwp()
        if not self.hwp:
            raise RuntimeError("한글에 연결할 수 없습니다.")

    def extract(self, hwp_path: str, output_yaml: str = None) -> str:
        """
        HWP 파일에서 셀 메타데이터 추출하여 YAML로 저장

        Args:
            hwp_path: HWP 파일 경로
            output_yaml: 출력 YAML 경로 (없으면 자동 생성)

        Returns:
            생성된 YAML 파일 경로
        """
        if output_yaml is None:
            output_yaml = os.path.splitext(hwp_path)[0] + "_meta.yaml"

        # 1. HWP 파일 열기
        self.hwp.Open(hwp_path)

        # 2. 임시 HWPX로 저장
        temp_dir = tempfile.gettempdir()
        temp_hwpx = os.path.join(temp_dir, "temp_extract_meta.hwpx")
        self._save_as(temp_hwpx, "HWPX")

        # 3. COM API로 각 셀의 list_id, para_id 추출
        cell_positions = self._extract_cell_positions()

        # 4. HWPX에서 field_name (tc.name 속성) 추출
        field_names = self._extract_field_names_from_hwpx(temp_hwpx)

        # 5. 병합하여 YAML 생성
        yaml_content = self._merge_to_yaml(cell_positions, field_names)

        # 6. YAML 저장
        with open(output_yaml, 'w', encoding='utf-8') as f:
            f.write(yaml_content)

        # 7. 임시 파일 삭제
        try:
            os.remove(temp_hwpx)
        except:
            pass

        return output_yaml

    def _save_as(self, filepath: str, format_type: str):
        """파일 저장"""
        self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        self.hwp.HParameterSet.HFileOpenSave.filename = filepath
        self.hwp.HParameterSet.HFileOpenSave.Format = format_type
        self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)

    def _extract_cell_positions(self) -> list:
        """COM API로 모든 테이블 셀의 list_id + 필드명(JSON)에서 row/col 추출"""
        tables = []
        ctrl = self.hwp.HeadCtrl
        tbl_idx = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                table_data = {
                    'tbl_idx': tbl_idx,
                    'cells': {}  # (row, col) -> list_id
                }

                try:
                    # 테이블로 이동
                    anchor = ctrl.GetAnchorPos(0)
                    self.hwp.SetPosBySet(anchor)
                    self.hwp.HAction.Run("SelectCtrlFront")
                    self.hwp.HAction.Run("ShapeObjTableSelCell")

                    first_list_id = self.hwp.GetPos()[0]

                    # 행별 순회
                    row = 0
                    while row < 100:  # 안전장치
                        # 행의 첫 열로 이동
                        self.hwp.HAction.Run("TableColBegin")
                        row_first_list_id = self.hwp.GetPos()[0]
                        col = 0

                        # 열 순회
                        while col < 100:  # 안전장치
                            pos = self.hwp.GetPos()
                            list_id = pos[0]

                            # GetCurFieldName(0)으로 필드명 가져오기
                            field_name = ""
                            try:
                                field_name = self.hwp.GetCurFieldName(0) or ""
                            except:
                                pass

                            # field_name에서 row/col 파싱
                            if field_name:
                                try:
                                    fd = json.loads(field_name)
                                    r = fd.get('rowAddr', 0)
                                    c = fd.get('colAddr', 0)
                                    if (r, c) not in table_data['cells']:
                                        table_data['cells'][(r, c)] = list_id
                                except:
                                    pass

                            # 오른쪽 셀로 이동
                            self.hwp.HAction.Run("TableRightCell")
                            new_pos = self.hwp.GetPos()

                            # 같은 행의 처음으로 돌아왔으면 열 순회 종료
                            if new_pos[0] == row_first_list_id:
                                break

                            col += 1

                        # 다음 행으로 이동
                        self.hwp.HAction.Run("TableLowerCell")
                        new_pos = self.hwp.GetPos()

                        # 첫 행으로 돌아왔으면 순회 종료
                        if new_pos[0] == first_list_id:
                            break

                        row += 1

                    # 선택 해제
                    self.hwp.HAction.Run("Cancel")
                    self.hwp.HAction.Run("MoveParentList")

                except Exception as e:
                    print(f"테이블 {tbl_idx} 처리 오류: {e}")

                tables.append(table_data)
                tbl_idx += 1

            ctrl = ctrl.Next

        return tables

    def _extract_field_names_from_hwpx(self, hwpx_path: str) -> list:
        """HWPX에서 테이블별 셀의 field_name (tc.name 속성) 추출"""
        tables = []

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                xml_content = zf.read(section_file)
                root = ET.parse(BytesIO(xml_content)).getroot()
                self._find_tables_recursive(root, tables)

        return tables

    def _find_tables_recursive(self, element, tables: list):
        """재귀적으로 테이블 찾아 셀 정보 추출"""
        for child in element:
            if child.tag.endswith('}tbl'):
                table_data = {
                    'table_id': child.get('id', ''),
                    'row_count': int(child.get('rowCnt', 0)),
                    'col_count': int(child.get('colCnt', 0)),
                    'cells': []
                }

                # 캡션 추출
                for sub in child:
                    if sub.tag.endswith('}caption'):
                        texts = []
                        for t in sub.iter():
                            if t.tag.endswith('}t') and t.text:
                                texts.append(t.text)
                        table_data['caption'] = ''.join(texts)
                        break

                # 셀 추출
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue

                        cell_data = {
                            'field_name': tc.get('name', ''),
                            'row': 0,
                            'col': 0,
                            'row_span': 1,
                            'col_span': 1,
                        }

                        for sub in tc:
                            tag = sub.tag.split('}')[-1]
                            if tag == 'cellAddr':
                                cell_data['row'] = int(sub.get('rowAddr', 0))
                                cell_data['col'] = int(sub.get('colAddr', 0))
                            elif tag == 'cellSpan':
                                cell_data['row_span'] = int(sub.get('rowSpan', 1))
                                cell_data['col_span'] = int(sub.get('colSpan', 1))

                        table_data['cells'].append(cell_data)

                tables.append(table_data)

                # 중첩 테이블 재귀 탐색
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue
                        self._find_tables_recursive(tc, tables)
            else:
                self._find_tables_recursive(child, tables)

    def _merge_to_yaml(self, cell_positions: list, field_names: list) -> str:
        """COM API 결과와 HWPX 결과 병합하여 YAML 생성"""
        lines = [
            '# HWP 테이블 셀 메타데이터',
            f'# 테이블 수: {len(field_names)}',
            '#',
            '# 셀 형식: [tblIdx, row, col, rowSpan, colSpan, list_id]',
            ''
        ]
        lines.append('tables:')

        for tbl_idx, tbl_xml in enumerate(field_names):
            lines.append(f'  - tbl_idx: {tbl_idx}')
            lines.append(f'    table_id: "{tbl_xml.get("table_id", "")}"')
            lines.append(f'    size: "{tbl_xml.get("row_count", 0)}x{tbl_xml.get("col_count", 0)}"')

            # COM API에서 가져온 list_id 매핑 (row, col) -> list_id
            com_cells = {}
            if tbl_idx < len(cell_positions):
                com_cells = cell_positions[tbl_idx].get('cells', {})

            # list_range 계산 (min ~ max list_id)
            if com_cells:
                list_ids = list(com_cells.values())
                min_list_id = min(list_ids)
                max_list_id = max(list_ids)
                lines.append(f'    list_range: [{min_list_id}, {max_list_id}]')
                # caption_list_id = 첫 번째 셀 list_id - 1
                caption_list_id = min_list_id - 1
                lines.append(f'    caption_list_id: {caption_list_id}')

            caption = tbl_xml.get('caption', '')
            if caption:
                caption_escaped = caption.replace('"', '\\"').replace('\n', ' ')
                lines.append(f'    caption: "{caption_escaped}"')

            lines.append('    cells:')

            for cell in tbl_xml.get('cells', []):
                row = cell['row']
                col = cell['col']
                row_span = cell['row_span']
                col_span = cell['col_span']

                # COM API에서 list_id 가져오기
                list_id = com_cells.get((row, col), 0)

                # field_name에서 tblIdx 파싱
                field_tbl_idx = tbl_idx
                if cell['field_name']:
                    try:
                        fd = json.loads(cell['field_name'])
                        field_tbl_idx = fd.get('tblIdx', tbl_idx)
                    except:
                        pass

                lines.append(f'      - [{field_tbl_idx}, {row}, {col}, {row_span}, {col_span}, {list_id}]')

            lines.append('')

        return '\n'.join(lines)


def extract_cell_meta(hwp_path: str, output_yaml: str = None) -> str:
    """HWP 파일에서 셀 메타데이터 추출"""
    extractor = ExtractCellMeta()
    return extractor.extract(hwp_path, output_yaml)


if __name__ == "__main__":
    # 테스트
    hwp_path = sys.argv[1] if len(sys.argv) > 1 else r"C:\hwp_xml\test_rev2.hwp"

    print(f"입력: {hwp_path}")
    print("=" * 60)

    extractor = ExtractCellMeta()
    output = extractor.extract(hwp_path)

    print(f"출력: {output}")
    print()
    print("=== 내용 미리보기 ===")
    with open(output, 'r', encoding='utf-8') as f:
        print(f.read()[:2000])
