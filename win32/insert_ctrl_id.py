# -*- coding: utf-8 -*-
"""
한글 테이블 ctrl_id 삽입 및 정보 출력 모듈

기능:
- 테이블 첫 셀에 ctrl_id 삽입: [list_id:X,para_id:X,char_id:X,tbl_N]
- 테이블 정보 별도 출력 (list_id, para_id, char_id, tbl_n)
"""

import sys
import os
from pathlib import Path

# 프로젝트 루트 경로 설정
_project_root = Path(__file__).parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

# config에서 외부 의존성 경로 가져오기
try:
    from config import WIN32HWP_DIR
    if str(WIN32HWP_DIR) not in sys.path:
        sys.path.insert(0, str(WIN32HWP_DIR))
except ImportError:
    # config가 없으면 환경변수 또는 기본 경로 사용
    win32hwp_dir = os.environ.get('WIN32HWP_DIR', r'C:\win32hwp')
    if win32hwp_dir not in sys.path:
        sys.path.insert(0, win32hwp_dir)

from cursor import get_hwp_instance
from dataclasses import dataclass
from typing import List, Optional, Any


@dataclass
class TableInfo:
    """테이블 정보"""
    index: int = 0
    tbl_id: str = ""
    list_id: int = 0
    para_id: int = 0
    char_id: int = 0
    first_cell_list_id: int = 0
    row_count: int = 0
    col_count: int = 0
    ctrl: Any = None


class InsertCtrlId:
    """테이블 ctrl_id 삽입 및 관리"""

    def __init__(self, hwp=None):
        self.hwp = hwp or get_hwp_instance()
        self.tables: List[TableInfo] = []

        if not self.hwp:
            raise RuntimeError("한글에 연결할 수 없습니다.")

    def _get_anchor_pos(self, ctrl):
        """컨트롤 앵커 위치 (list_id, para_id, char_id)"""
        try:
            anchor = ctrl.GetAnchorPos(0)
            return (
                anchor.Item("List"),
                anchor.Item("Para"),
                anchor.Item("Pos")
            )
        except:
            return (0, 0, 0)

    def _get_first_cell_list_id(self, ctrl) -> int:
        """첫 번째 셀의 list_id"""
        try:
            self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            self.hwp.HAction.Run("SelectCtrlFront")
            self.hwp.HAction.Run("ShapeObjTableSelCell")

            first_cell_list_id = self.hwp.GetPos()[0]

            self.hwp.HAction.Run("Cancel")
            self.hwp.HAction.Run("MoveParentList")

            return first_cell_list_id
        except:
            return 0

    def _get_table_size(self, ctrl):
        """테이블 행/열 수"""
        try:
            props = ctrl.Properties
            return props.Item('RowCnt') or 0, props.Item('ColCnt') or 0
        except:
            return 0, 0

    def collect_tables(self) -> List[TableInfo]:
        """모든 테이블 정보 수집"""
        self.tables.clear()
        ctrl = self.hwp.HeadCtrl
        index = 0

        while ctrl:
            try:
                if ctrl.CtrlID == "tbl":
                    list_id, para_id, char_id = self._get_anchor_pos(ctrl)
                    first_cell = self._get_first_cell_list_id(ctrl)
                    row, col = self._get_table_size(ctrl)

                    info = TableInfo(
                        index=index,
                        tbl_id=f"tbl_{index}",
                        list_id=list_id,
                        para_id=para_id,
                        char_id=char_id,
                        first_cell_list_id=first_cell,
                        row_count=row,
                        col_count=col,
                        ctrl=ctrl
                    )
                    self.tables.append(info)
                    index += 1
            except:
                pass
            ctrl = ctrl.Next

        return self.tables

    def insert_ctrl_id(self, info: TableInfo) -> bool:
        """테이블 첫 셀에 ctrl_id 삽입"""
        ctrl_id_text = f"[list_id:{info.list_id},para_id:{info.para_id},char_id:{info.char_id},{info.tbl_id}]"

        try:
            # 첫 셀 첫 위치로 이동
            self.hwp.SetPos(info.first_cell_list_id, 0, 0)

            # 텍스트 삽입
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = ctrl_id_text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

            return True
        except Exception as e:
            print(f"삽입 실패: {e}")
            return False

    def insert_all(self) -> int:
        """모든 테이블에 ctrl_id 삽입"""
        if not self.tables:
            self.collect_tables()

        count = 0
        for info in self.tables:
            if self.insert_ctrl_id(info):
                count += 1
                print(f"테이블 {info.index}: ctrl_id 삽입 완료")

        return count

    # =========================================================================
    # 출력 기능
    # =========================================================================

    def print_tables(self):
        """모든 테이블 정보 출력"""
        if not self.tables:
            self.collect_tables()

        print("\n" + "=" * 60)
        print(f"테이블 목록 ({len(self.tables)}개)")
        print("=" * 60)

        for t in self.tables:
            print(f"\n[{t.tbl_id}]")
            print(f"  list_id : {t.list_id}")
            print(f"  para_id : {t.para_id}")
            print(f"  char_id : {t.char_id}")
            print(f"  크기    : {t.row_count} x {t.col_count}")

    def get_list_ids(self) -> List[int]:
        """모든 테이블의 list_id 목록"""
        return [t.list_id for t in self.tables]

    def get_para_ids(self) -> List[int]:
        """모든 테이블의 para_id 목록"""
        return [t.para_id for t in self.tables]

    def get_char_ids(self) -> List[int]:
        """모든 테이블의 char_id 목록"""
        return [t.char_id for t in self.tables]

    def get_tbl_ids(self) -> List[str]:
        """모든 테이블의 tbl_id 목록"""
        return [t.tbl_id for t in self.tables]

    def print_ids_separately(self):
        """list_id, para_id, char_id, tbl_id 별도 출력"""
        if not self.tables:
            self.collect_tables()

        print("\n[list_id 목록]")
        print(self.get_list_ids())

        print("\n[para_id 목록]")
        print(self.get_para_ids())

        print("\n[char_id 목록]")
        print(self.get_char_ids())

        print("\n[tbl_id 목록]")
        print(self.get_tbl_ids())


def main():
    print("테이블 ctrl_id 삽입 시작...")

    try:
        ctrl = InsertCtrlId()
    except RuntimeError as e:
        print(e)
        return

    # 테이블 수집
    tables = ctrl.collect_tables()
    print(f"발견된 테이블: {len(tables)}개")

    if not tables:
        print("테이블이 없습니다.")
        return

    # 정보 출력
    ctrl.print_tables()
    ctrl.print_ids_separately()

    # ctrl_id 삽입
    print("\n" + "-" * 60)
    count = ctrl.insert_all()
    print(f"\n완료! {count}개 테이블에 ctrl_id 삽입")


if __name__ == "__main__":
    main()
