# -*- coding: utf-8 -*-
"""
한글 파일의 테이블 정보 조회 스크립트
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
    win32hwp_dir = os.environ.get('WIN32HWP_DIR', r'C:\win32hwp')
    if win32hwp_dir not in sys.path:
        sys.path.insert(0, win32hwp_dir)

from cursor import get_hwp_instance


def get_all_table_info(hwp):
    """모든 테이블 정보 조회"""
    tables = []
    ctrl = hwp.HeadCtrl
    index = 0

    while ctrl:
        try:
            if ctrl.CtrlID == "tbl":
                # 앵커 위치
                anchor = ctrl.GetAnchorPos(0)
                list_id = anchor.Item("List")
                para_id = anchor.Item("Para")
                char_id = anchor.Item("Pos")

                # 테이블로 이동하여 첫 셀 list_id 획득
                hwp.SetPosBySet(anchor)
                hwp.HAction.Run("SelectCtrlFront")
                hwp.HAction.Run("ShapeObjTableSelCell")

                pos = hwp.GetPos()
                first_cell_list_id = pos[0]

                # 행/열 수
                row_cnt = 0
                col_cnt = 0
                try:
                    props = ctrl.Properties
                    row_cnt = props.Item("RowCnt") or 0
                    col_cnt = props.Item("ColCnt") or 0
                except:
                    pass

                tables.append({
                    'index': index,
                    'list_id': list_id,
                    'para_id': para_id,
                    'char_id': char_id,
                    'first_cell_list_id': first_cell_list_id,
                    'row_cnt': row_cnt,
                    'col_cnt': col_cnt,
                    'ctrl': ctrl,
                })

                # 선택 해제
                hwp.HAction.Run("Cancel")
                hwp.HAction.Run("MoveParentList")

                index += 1
        except Exception as e:
            print(f"오류: {e}")

        ctrl = ctrl.Next

    return tables


def main():
    print("=" * 60)
    print("한글 파일 테이블 정보 조회")
    print("=" * 60)

    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        return

    tables = get_all_table_info(hwp)
    print(f"\n총 테이블 수: {len(tables)}개\n")

    for t in tables:
        print(f"[테이블 {t['index']}]")
        print(f"  list_id          : {t['list_id']}")
        print(f"  para_id          : {t['para_id']}")
        print(f"  char_id          : {t['char_id']}")
        print(f"  first_cell_list_id: {t['first_cell_list_id']}")
        print(f"  행/열            : {t['row_cnt']} x {t['col_cnt']}")
        print()


if __name__ == "__main__":
    main()
