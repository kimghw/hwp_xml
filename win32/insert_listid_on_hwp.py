# -*- coding: utf-8 -*-
"""열린 HWP 파일의 각 셀에 list_id 삽입"""
import sys
import json

sys.path.insert(0, r'C:\win32hwp')
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글 인스턴스 없음")
    sys.exit(1)

ctrl = hwp.HeadCtrl
tbl_count = 0

while ctrl:
    if ctrl.CtrlID == "tbl":
        tbl_count += 1
        print(f"\n테이블 {tbl_count} 처리 중...")

        try:
            anchor = ctrl.GetAnchorPos(0)
            hwp.SetPosBySet(anchor)
            hwp.HAction.Run("SelectCtrlFront")
            hwp.HAction.Run("ShapeObjTableSelCell")

            first_list_id = hwp.GetPos()[0]
            processed = set()

            row = 0
            while row < 100:
                hwp.HAction.Run("TableColBegin")
                row_first_list_id = hwp.GetPos()[0]
                col = 0

                while col < 100:
                    pos = hwp.GetPos()
                    list_id = pos[0]

                    if list_id not in processed:
                        processed.add(list_id)
                        # 셀에 list_id 텍스트 삽입
                        hwp.HAction.Run("MoveLineEnd")
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = f"\n[list_id:{list_id}]"
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                    hwp.HAction.Run("TableRightCell")
                    new_pos = hwp.GetPos()

                    if new_pos[0] == row_first_list_id:
                        break
                    col += 1

                hwp.HAction.Run("TableLowerCell")
                new_pos = hwp.GetPos()

                if new_pos[0] == first_list_id:
                    break
                row += 1

            hwp.HAction.Run("Cancel")
            hwp.HAction.Run("MoveParentList")
            print(f"  {len(processed)}개 셀에 list_id 삽입 완료")

        except Exception as e:
            print(f"  오류: {e}")

    ctrl = ctrl.Next

print(f"\n총 {tbl_count}개 테이블 처리 완료")
