# -*- coding: utf-8 -*-
"""
HWP 파일의 테이블 셀에 필드 이름 삽입

프로세스:
1. HWP -> HWPX 변환 (임시폴더)
2. HWPX XML에서 tc 태그의 name 속성에 JSON 형식 필드 정보 설정
3. HWPX -> HWP 변환 후 저장

삽입되는 필드 이름 (JSON 형식, XML 속성명 사용):
- {"tblIdx":N,"rowAddr":R,"colAddr":C,"rowSpan":RS,"colSpan":CS,"type":"parent|nested","parentTbl":P,"parentCell":[R,C]}
  - tblIdx: 테이블 인덱스
  - rowAddr: 행 번호 (cellAddr/@rowAddr)
  - colAddr: 열 번호 (cellAddr/@colAddr)
  - rowSpan: 행 병합 (cellSpan/@rowSpan)
  - colSpan: 열 병합 (cellSpan/@colSpan)
  - type: "parent" (최상위) 또는 "nested" (중첩)
  - parentTbl: 부모 테이블 인덱스 (nested인 경우)
  - parentCell: 부모 셀 위치 [row, col] (nested인 경우)
"""

import sys
import os
import json
import tempfile
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

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
from dataclasses import dataclass
from typing import List
import win32com.client as win32


def get_or_create_hwp():
    """실행 중인 한글에 연결하거나, 없으면 새로 실행"""
    hwp = get_hwp_instance()
    if hwp:
        # 기존 인스턴스도 보안 모듈 등록 확인
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except:
            pass
        return hwp
    # 한글 새로 실행
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    return hwp


@dataclass
class TableInfo:
    """테이블 정보 (XML 기반)"""
    index: int = 0
    table_id: str = ""
    row_count: int = 0
    col_count: int = 0
    section_index: int = 0
    cell_count: int = 0
    list_id: int = 0  # win32 API용


class InsertTableField:
    """테이블 셀 필드 이름 삽입 (HWPX XML 직접 수정)"""

    # HWPX 네임스페이스
    NAMESPACES = {
        'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
        'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    }

    def __init__(self, hwp=None):
        self.hwp = hwp or get_or_create_hwp()
        self.tables: List[TableInfo] = []

        if not self.hwp:
            raise RuntimeError("한글에 연결할 수 없습니다.")

        # XML 네임스페이스 등록
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)

    def open_file(self, filepath: str) -> bool:
        """파일 열기"""
        try:
            self.hwp.Open(filepath)
            print(f"파일 열림: {filepath}")
            return True
        except Exception as e:
            print(f"파일 열기 실패: {e}")
            return False

    def save_as(self, filepath: str, format: str = "HWP") -> bool:
        """파일 저장"""
        try:
            self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = filepath
            self.hwp.HParameterSet.HFileOpenSave.Format = format.upper()
            self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
            print(f"파일 저장: {filepath} ({format})")
            return True
        except Exception as e:
            print(f"파일 저장 실패: {e}")
            return False

    def _get_cell_info(self, tc_element: ET.Element) -> dict:
        """셀의 행/열 주소 및 span 정보 추출"""
        info = {
            'row': 0,
            'col': 0,
            'row_span': 1,
            'col_span': 1
        }
        for child in tc_element:
            if child.tag.endswith('}cellAddr'):
                info['row'] = int(child.get('rowAddr', 0))
                info['col'] = int(child.get('colAddr', 0))
            elif child.tag.endswith('}cellSpan'):
                info['row_span'] = int(child.get('rowSpan', 1))
                info['col_span'] = int(child.get('colSpan', 1))
        return info

    def insert_field_to_xml(self, hwpx_path: str) -> int:
        """
        HWPX 파일의 tc 태그 name 속성에 필드 이름 설정

        Returns:
            수정된 테이블 수
        """
        temp_dir = tempfile.mkdtemp()
        self._table_global_index = 0

        try:
            # HWPX 압축 해제
            with zipfile.ZipFile(hwpx_path, 'r') as zipf:
                zipf.extractall(temp_dir)

            # section 파일 목록
            contents_dir = os.path.join(temp_dir, 'Contents')
            section_files = sorted([
                f for f in os.listdir(contents_dir)
                if f.startswith('section') and f.endswith('.xml')
            ])

            for section_idx, section_file in enumerate(section_files):
                section_path = os.path.join(contents_dir, section_file)

                # XML 파일 읽기
                tree = ET.parse(section_path)
                root = tree.getroot()

                # 재귀적으로 테이블 처리
                self._process_tables_recursive(root, section_idx, None, None, None)

                # 수정된 XML 저장
                tree.write(section_path, encoding='utf-8', xml_declaration=True)

            # 수정된 HWPX 다시 압축
            with zipfile.ZipFile(hwpx_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return len(self.tables)

    def _process_tables_recursive(self, element, section_idx: int,
                                   parent_tbl_idx: int, parent_cell_row: int, parent_cell_col: int):
        """재귀적으로 테이블을 찾아 필드 이름 설정"""
        for child in element:
            if child.tag.endswith('}tbl'):
                # 먼저 현재 테이블 인덱스 할당 후 증가
                current_tbl_idx = self._table_global_index
                self._table_global_index += 1

                table_id = child.get('id', '')
                row_cnt = int(child.get('rowCnt', '0'))
                col_cnt = int(child.get('colCnt', '0'))

                is_nested = parent_tbl_idx is not None
                tbl_type = "nested" if is_nested else "parent"

                cell_count = 0
                is_first_cell = True

                # 모든 행(tr) 순회
                for tr in child:
                    if not tr.tag.endswith('}tr'):
                        continue

                    # 모든 셀(tc) 순회
                    for tc in tr:
                        if not tc.tag.endswith('}tc'):
                            continue

                        # 셀 정보 추출
                        cell_info = self._get_cell_info(tc)

                        # JSON 형식 필드 이름
                        field_data = {
                            "tblIdx": current_tbl_idx,
                            "rowAddr": cell_info['row'],
                            "colAddr": cell_info['col'],
                            "rowSpan": cell_info['row_span'],
                            "colSpan": cell_info['col_span'],
                            "type": tbl_type
                        }

                        # nested인 경우 부모 정보 추가
                        if is_nested:
                            field_data["parentTbl"] = parent_tbl_idx
                            field_data["parentCell"] = [parent_cell_row, parent_cell_col]

                        cell_field_name = json.dumps(field_data, separators=(',', ':'))

                        # 첫 번째 셀 로그 출력
                        if is_first_cell:
                            if is_nested:
                                print(f"테이블 {current_tbl_idx}: (id:{table_id}, {row_cnt}x{col_cnt}) [nested in tbl_{parent_tbl_idx} cell[{parent_cell_row},{parent_cell_col}]]")
                            else:
                                print(f"테이블 {current_tbl_idx}: (id:{table_id}, {row_cnt}x{col_cnt}) [parent]")
                            is_first_cell = False

                        # tc 태그의 name 속성 설정
                        tc.set('name', cell_field_name)
                        cell_count += 1

                        # 셀 내부의 중첩 테이블 재귀 탐색
                        self._process_tables_recursive(
                            tc, section_idx,
                            current_tbl_idx, cell_info['row'], cell_info['col']
                        )

                if cell_count > 0:
                    info = TableInfo(
                        index=current_tbl_idx,
                        table_id=table_id,
                        row_count=row_cnt,
                        col_count=col_cnt,
                        section_index=section_idx,
                        cell_count=cell_count
                    )
                    self.tables.append(info)
                    print(f"  - {cell_count}개 셀에 필드 이름 설정")

            else:
                # tbl이 아닌 요소는 자식 탐색
                self._process_tables_recursive(element=child, section_idx=section_idx,
                                               parent_tbl_idx=parent_tbl_idx,
                                               parent_cell_row=parent_cell_row,
                                               parent_cell_col=parent_cell_col)

    def collect_table_list_ids(self) -> List[dict]:
        """win32 API로 테이블 list_id 수집 (첫 셀 list_id 기준)"""
        table_infos = []
        ctrl = self.hwp.HeadCtrl
        idx = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                try:
                    # 테이블로 이동하여 첫 셀 list_id 획득
                    anchor = ctrl.GetAnchorPos(0)
                    self.hwp.SetPosBySet(anchor)
                    self.hwp.HAction.Run("SelectCtrlFront")
                    self.hwp.HAction.Run("ShapeObjTableSelCell")

                    cell_pos = self.hwp.GetPos()
                    first_cell_list_id = cell_pos[0]

                    self.hwp.HAction.Run("Cancel")
                    self.hwp.HAction.Run("MoveParentList")

                    # 캡션 list_id = 첫 셀 list_id - 1
                    caption_list_id = first_cell_list_id - 1

                    table_infos.append({
                        'index': idx,
                        'first_cell_list_id': first_cell_list_id,
                        'caption_list_id': caption_list_id,
                        'ctrl': ctrl
                    })
                    idx += 1
                except:
                    pass
            ctrl = ctrl.Next

        return table_infos

    def insert_caption_text(self, table_infos: List[dict]) -> int:
        """테이블 캡션에 {caption:tbl_N|} 삽입 (캡션 직접 선택)"""
        count = 0

        for info in table_infos:
            try:
                ctrl = info['ctrl']

                # 테이블 앵커로 이동
                anchor = ctrl.GetAnchorPos(0)
                self.hwp.SetPosBySet(anchor)

                # 테이블 선택
                self.hwp.HAction.Run("SelectCtrlFront")

                # 캡션 선택 (표/그림 캡션)
                self.hwp.HAction.Run("TableCaptionCellCreate")

                # 캡션 텍스트 삽입
                caption_text = f"{{caption:tbl_{info['index']}|}}"

                self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                self.hwp.HParameterSet.HInsertText.Text = caption_text
                self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

                # 선택 해제
                self.hwp.HAction.Run("Cancel")

                print(f"  캡션 삽입: tbl_{info['index']}")
                count += 1

            except Exception as e:
                print(f"  캡션 삽입 실패 (tbl_{info['index']}): {e}")

        return count

    def print_tables(self):
        """테이블 정보 출력"""
        print("\n" + "=" * 60)
        print(f"테이블 목록 ({len(self.tables)}개)")
        print("=" * 60)

        for t in self.tables:
            print(f"\n[tbl_{t.index}]")
            print(f"  table_id  : {t.table_id}")
            print(f"  크기      : {t.row_count} x {t.col_count}")
            print(f"  셀 개수   : {t.cell_count}")
            print(f"  섹션      : {t.section_index}")


def process_hwp_file(input_hwp: str, output_hwp: str = None) -> bool:
    """
    HWP 파일 처리 메인 함수
    """
    if output_hwp is None:
        output_hwp = input_hwp

    print("=" * 60)
    print("테이블 셀 필드 이름 삽입")
    print("=" * 60)
    print(f"입력: {input_hwp}")
    print(f"출력: {output_hwp}")
    print()

    try:
        inserter = InsertTableField()
    except RuntimeError as e:
        print(f"오류: {e}")
        return False

    # 1. HWP 파일 열기
    if not inserter.open_file(input_hwp):
        return False

    # 2. 임시 폴더에 HWPX로 저장
    temp_dir = tempfile.gettempdir()
    temp_hwpx = os.path.join(temp_dir, "temp_table_field.hwpx")

    if not inserter.save_as(temp_hwpx, "HWPX"):
        return False

    # 3. HWPX XML 수정 (tc name 속성 설정)
    print("\n" + "-" * 60)
    print("XML 수정 중...")
    count = inserter.insert_field_to_xml(temp_hwpx)
    print(f"\n{count}개 테이블 처리 완료")

    if count == 0:
        print("테이블이 없습니다.")
        inserter.hwp.Clear(1)
        try:
            os.remove(temp_hwpx)
        except:
            pass
        return False

    # 4. 테이블 정보 출력
    inserter.print_tables()

    # 5. 수정된 HWPX 열어서 캡션 삽입 후 HWP로 저장
    print("\n" + "-" * 60)
    if not inserter.open_file(temp_hwpx):
        return False

    # 6. 테이블 list_id 수집 및 캡션 삽입
    print("\n캡션 삽입 중...")
    table_infos = inserter.collect_table_list_ids()
    caption_count = inserter.insert_caption_text(table_infos)
    print(f"{caption_count}개 테이블에 캡션 삽입 완료")

    if not inserter.save_as(output_hwp, "HWP"):
        return False

    # 7. 문서 닫기 (파일 잠금 해제)
    try:
        inserter.hwp.Clear(1)  # 1: 저장 안 함 (이미 저장했으므로)
        print("문서 닫기 완료")
    except:
        pass

    # 8. HWPX → Excel 변환 (모든 테이블)
    try:
        from excel.hwpx_to_excel import HwpxToExcel

        excel_path = os.path.splitext(output_hwp)[0] + ".xlsx"
        converter = HwpxToExcel()
        converter.convert_all(temp_hwpx, excel_path)
        print(f"Excel 변환 완료: {excel_path}")
    except Exception as e:
        print(f"Excel 변환 실패: {e}")

    # 9. 메타데이터 YAML 추출
    try:
        from hwpxml.export_meta_yaml import export_meta_yaml

        yaml_path = os.path.splitext(output_hwp)[0] + "_meta.yaml"
        export_meta_yaml(temp_hwpx, yaml_path)
        print(f"메타데이터 YAML 저장: {yaml_path}")
    except Exception as e:
        print(f"YAML 추출 실패: {e}")

    # 10. 임시 파일 삭제
    try:
        os.remove(temp_hwpx)
        print(f"임시 파일 삭제: {temp_hwpx}")
    except:
        pass

    print("\n" + "=" * 60)
    print("처리 완료!")
    print("=" * 60)

    return True


def get_file_from_dialog() -> str:
    """파일 탐색 대화상자로 HWP 파일 선택"""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # 메인 창 숨김
    root.attributes('-topmost', True)  # 최상위로

    file_path = filedialog.askopenfilename(
        title="HWP 파일 선택",
        filetypes=[("HWP 파일", "*.hwp"), ("모든 파일", "*.*")],
        initialdir=r"C:\hwp_xml"
    )

    root.destroy()
    return file_path


def get_file_from_hwp(hwp) -> str:
    """열려있는 한글 문서 경로 반환"""
    try:
        path = hwp.XHwpDocuments.Active_XHwpDocument.FullName
        if path and os.path.exists(path):
            return path
    except:
        pass
    return ""


def main():
    input_file = None

    # 1. 명령줄 인자 확인
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        if not os.path.exists(input_file):
            print(f"파일을 찾을 수 없습니다: {input_file}")
            input_file = None

    # 2. 열려있는 한글 문서 확인
    if not input_file:
        try:
            hwp = get_hwp_instance()
            if hwp:
                input_file = get_file_from_hwp(hwp)
                if input_file:
                    print(f"열려있는 문서 사용: {input_file}")
        except:
            pass

    # 3. 파일 탐색 대화상자
    if not input_file:
        print("파일을 선택하세요...")
        input_file = get_file_from_dialog()
        if not input_file:
            print("파일이 선택되지 않았습니다.")
            return

    output_file = input_file
    process_hwp_file(input_file, output_file)


if __name__ == "__main__":
    main()
