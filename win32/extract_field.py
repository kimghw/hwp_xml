# -*- coding: utf-8 -*-
"""
HWP 파일에서 기존 필드 추출 및 삭제

기능:
1. HWP 파일의 모든 필드를 추출하여 list_id와 필드명을 YAML로 저장
2. 추출 후 기존 필드 삭제

출력 파일: {파일명}_field.yaml
"""

import sys
import os
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

try:
    from hwp_utils import get_hwp_instance, create_hwp_instance, get_active_filepath
except ImportError:
    from win32.hwp_utils import get_hwp_instance, create_hwp_instance, get_active_filepath
from typing import List, Dict, Optional
from dataclasses import dataclass, field


@dataclass
class FieldInfo:
    """필드 정보"""
    name: str  # 필드 이름
    list_id: int  # list_id
    para_id: int = 0  # para_id
    char_pos: int = 0  # 문자 위치
    text: str = ""  # 필드 텍스트


# 필드 옵션 상수
FIELD_CELL = 1          # 셀 필드
FIELD_CLICKHERE = 2     # 누름틀 필드
FIELD_NUMBER = 1        # {{#}} 형식 일련번호


class ExtractField:
    """HWP 필드 추출 및 삭제"""

    def __init__(self, hwp=None):
        self.hwp = hwp
        if not self.hwp:
            self.hwp = get_hwp_instance()
        if not self.hwp:
            self.hwp = create_hwp_instance(visible=True)
        if not self.hwp:
            raise RuntimeError("한글에 연결할 수 없습니다.")

        self.fields: List[FieldInfo] = []

    def extract_fields(self, option: int = 0) -> List[FieldInfo]:
        """
        문서의 모든 필드 추출 (GetFieldList API 사용)

        Args:
            option: 필드 옵션 (0=모두, 1=셀, 2=누름틀)

        Returns:
            필드 정보 리스트
        """
        self.fields = []

        # GetFieldList로 모든 필드 이름 가져오기
        field_list_str = self.hwp.GetFieldList(FIELD_NUMBER, option)
        if not field_list_str:
            print("  필드 없음")
            return self.fields

        field_names = field_list_str.split('\x02')

        for name in field_names:
            if not name:
                continue

            try:
                # 필드로 이동하여 위치 정보 획득
                if self.hwp.MoveToField(name, True, True, False):
                    pos = self.hwp.GetPos()
                    list_id = pos[0] if pos else 0
                    para_id = pos[1] if len(pos) > 1 else 0
                    char_pos = pos[2] if len(pos) > 2 else 0

                    # 필드 텍스트 획득
                    text = self.hwp.GetFieldText(name) or ""
                    text = text.strip('\x02')

                    field_info = FieldInfo(
                        name=name,
                        list_id=list_id,
                        para_id=para_id,
                        char_pos=char_pos,
                        text=text
                    )
                    self.fields.append(field_info)

            except Exception as e:
                print(f"  필드 추출 오류 ({name}): {e}")

        print(f"  총 {len(self.fields)}개 필드 추출")
        return self.fields

    def extract_cell_field_names(self) -> List[FieldInfo]:
        """
        테이블 셀에 설정된 필드 이름 추출 (GetCurFieldName API 사용)

        각 테이블의 각 셀을 순회하며 셀에 부여된 필드 이름을 추출합니다.

        Returns:
            필드 정보 리스트
        """
        self.fields = []

        # HeadCtrl로 모든 테이블 순회
        ctrl = self.hwp.HeadCtrl
        tbl_idx = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                try:
                    # 테이블로 이동
                    anchor = ctrl.GetAnchorPos(0)
                    self.hwp.SetPosBySet(anchor)
                    self.hwp.HAction.Run("SelectCtrlFront")
                    self.hwp.HAction.Run("ShapeObjTableSelCell")

                    # 첫 셀에서 시작하여 모든 셀 순회
                    visited = set()
                    while True:
                        pos = self.hwp.GetPos()
                        list_id = pos[0]

                        if list_id in visited:
                            break
                        visited.add(list_id)

                        # 현재 셀의 필드 이름 가져오기 (option=1: 셀 필드)
                        field_name = self.hwp.GetCurFieldName(FIELD_CELL)

                        if field_name:
                            field_info = FieldInfo(
                                name=field_name,
                                list_id=list_id,
                                para_id=pos[1] if len(pos) > 1 else 0,
                                char_pos=pos[2] if len(pos) > 2 else 0,
                                text=""
                            )
                            self.fields.append(field_info)

                        # 다음 셀로 이동 (오른쪽 → 다음 행)
                        self.hwp.HAction.Run("TableRightCell")

                    self.hwp.HAction.Run("Cancel")
                    self.hwp.HAction.Run("MoveParentList")
                    tbl_idx += 1

                except Exception as e:
                    print(f"  테이블 {tbl_idx} 처리 오류: {e}")
                    try:
                        self.hwp.HAction.Run("Cancel")
                    except:
                        pass

            ctrl = ctrl.Next

        print(f"  총 {len(self.fields)}개 셀 필드 이름 추출 ({tbl_idx}개 테이블)")
        return self.fields

    def delete_all_fields(self) -> int:
        """
        문서의 모든 필드 삭제 (필드 속성만 제거, 텍스트 유지)

        Returns:
            삭제된 필드 수
        """
        deleted = 0

        # 추출된 필드 목록 사용
        if not self.fields:
            self.extract_fields()

        for field in self.fields:
            try:
                # 필드로 이동
                if self.hwp.MoveToField(field.name, True, True, False):
                    # 필드 삭제 (텍스트는 유지, 필드 속성만 제거)
                    # UnsetFieldName: 현재 위치 필드 해제
                    self.hwp.HAction.Run("DeleteField")
                    deleted += 1
            except Exception as e:
                print(f"  필드 삭제 오류 ({field.name}): {e}")

        print(f"  총 {deleted}개 필드 삭제")
        return deleted

    def to_yaml(self) -> str:
        """
        추출된 필드를 YAML 형식으로 변환

        Returns:
            YAML 문자열
        """
        lines = [
            "# HWP 필드 정보",
            f"# 필드 수: {len(self.fields)}",
            "#",
            "# 형식: [name, list_id, text]",
            "",
            "fields:"
        ]

        for f in self.fields:
            # 텍스트에 특수문자 있으면 따옴표로 감싸기
            text_escaped = f.text.replace('"', '\\"') if f.text else ""
            lines.append(f"  - [\"{f.name}\", {f.list_id}, \"{text_escaped}\"]")

        return "\n".join(lines)

    def save_yaml(self, output_path: str) -> bool:
        """
        YAML 파일로 저장

        Args:
            output_path: 출력 파일 경로

        Returns:
            성공 여부
        """
        try:
            yaml_content = self.to_yaml()
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(yaml_content)
            print(f"  YAML 저장: {output_path}")
            return True
        except Exception as e:
            print(f"  YAML 저장 실패: {e}")
            return False


def extract_and_delete_fields(hwp=None, filepath: str = None, delete: bool = True) -> str:
    """
    HWP 파일에서 필드 추출 및 삭제

    Args:
        hwp: 한글 인스턴스 (없으면 자동 연결)
        filepath: HWP 파일 경로 (없으면 열린 문서 사용)
        delete: 필드 삭제 여부

    Returns:
        생성된 YAML 파일 경로
    """
    print("=" * 60)
    print("HWP 필드 추출" + (" 및 삭제" if delete else ""))
    print("=" * 60)

    extractor = ExtractField(hwp)

    # 파일 경로 확인
    if not filepath:
        filepath = get_active_filepath(extractor.hwp)
    if not filepath:
        raise ValueError("파일 경로를 찾을 수 없습니다.")

    print(f"파일: {filepath}")
    print()

    # 1. 필드 추출
    print("필드 추출 중...")
    fields = extractor.extract_fields()

    if not fields:
        print("필드가 없습니다.")
        return None

    # 2. YAML 저장
    base_path = os.path.splitext(filepath)[0]
    yaml_path = base_path + "_field.yaml"
    extractor.save_yaml(yaml_path)

    # 3. 필드 삭제 (옵션)
    if delete:
        print()
        print("필드 삭제 중...")
        extractor.delete_all_fields()

    print()
    print("=" * 60)
    print("완료!")
    print(f"  추출된 필드: {len(fields)}개")
    print(f"  YAML 파일: {yaml_path}")
    print("=" * 60)

    return yaml_path


def main():
    """메인 함수"""
    filepath = None
    delete = True

    # 명령줄 인자 처리
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        if not os.path.isabs(filepath):
            filepath = os.path.abspath(filepath)

    if len(sys.argv) > 2:
        if sys.argv[2].lower() in ('false', 'no', '0'):
            delete = False

    extract_and_delete_fields(filepath=filepath, delete=delete)


if __name__ == "__main__":
    main()
