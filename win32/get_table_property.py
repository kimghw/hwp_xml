"""
한글 COM API를 사용한 테이블 속성 추출 클래스

아래아한글 프로그램의 win32com API를 사용하여 테이블의 속성과 내용을 추출합니다.
Windows 환경에서 한글 프로그램이 설치되어 있어야 합니다.

주요 속성:
- link_id: 연결 ID
- para_id: 문단 ID
- ctrl: 컨트롤 객체
- char_id: 문자 ID
- 기타 테이블 관련 속성들
"""

import sys
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from enum import IntEnum


# 한글 컨트롤 타입 상수
class CtrlType(IntEnum):
    """한글 컨트롤 타입"""
    TABLE = 0x6c627474  # 'tbl' - 표
    GSOBJECT = 0x2467736f  # '$gso' - 그리기 개체
    EQUATION = 0x6575716e  # 'eqn' - 수식
    SECTION = 0x73656364  # 'secd' - 구역
    COLUMN = 0x636f6c64  # 'cold' - 단
    HEADER = 0x68656164  # 'head' - 머리말
    FOOTER = 0x666f6f74  # 'foot' - 꼬리말
    FOOTNOTE = 0x666e2020  # 'fn  ' - 각주
    ENDNOTE = 0x656e2020  # 'en  ' - 미주
    FIELD = 0x25666c64  # '%fld' - 필드


@dataclass
class CellInfo:
    """테이블 셀 정보"""
    row: int
    col: int
    text: str
    # 셀 속성
    row_span: int = 1
    col_span: int = 1
    width: int = 0
    height: int = 0
    # ID 정보
    para_id: Optional[int] = None
    char_id: Optional[int] = None


@dataclass
class TableProperty:
    """테이블 속성 정보"""
    # 기본 식별 정보
    table_index: int = 0
    ctrl: Any = None  # 컨트롤 객체 참조

    # ID 정보
    link_id: Optional[int] = None
    para_id: Optional[int] = None
    char_id: Optional[int] = None
    ctrl_id: Optional[str] = None

    # 테이블 구조
    row_count: int = 0
    col_count: int = 0

    # 테이블 속성
    treat_as_char: bool = False  # 글자처럼 취급
    protect: bool = False  # 보호
    page_break: str = ""  # 페이지 나눔
    repeat_header: bool = False  # 제목 줄 반복

    # 크기 정보
    width: int = 0
    height: int = 0

    # 위치 정보
    x_pos: int = 0
    y_pos: int = 0

    # 여백
    left_margin: int = 0
    right_margin: int = 0
    top_margin: int = 0
    bottom_margin: int = 0

    # 셀 간격
    cell_spacing: int = 0

    # 테두리/배경
    border_fill_id: Optional[int] = None

    # 셀 데이터
    cells: List[List[CellInfo]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """속성을 딕셔너리로 반환"""
        return {
            'table_index': self.table_index,
            'link_id': self.link_id,
            'para_id': self.para_id,
            'char_id': self.char_id,
            'ctrl_id': self.ctrl_id,
            'row_count': self.row_count,
            'col_count': self.col_count,
            'treat_as_char': self.treat_as_char,
            'protect': self.protect,
            'page_break': self.page_break,
            'repeat_header': self.repeat_header,
            'width': self.width,
            'height': self.height,
            'x_pos': self.x_pos,
            'y_pos': self.y_pos,
            'left_margin': self.left_margin,
            'right_margin': self.right_margin,
            'top_margin': self.top_margin,
            'bottom_margin': self.bottom_margin,
            'cell_spacing': self.cell_spacing,
            'border_fill_id': self.border_fill_id,
        }

    def get_data_as_2d_list(self) -> List[List[str]]:
        """셀 데이터를 2D 리스트로 반환"""
        if not self.cells:
            return []
        return [[cell.text for cell in row] for row in self.cells]

    def to_dataframe(self):
        """pandas DataFrame으로 변환"""
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
    한글 COM API를 사용하여 테이블 속성을 추출하는 클래스

    사용 예:
        # 기존 한글 인스턴스에 연결
        getter = GetTableProperty()

        # 파일 열기
        getter.open("document.hwp")

        # 모든 테이블 속성 가져오기
        tables = getter.get_all_tables()

        # 특정 테이블 속성 가져오기
        table = getter.get_table_by_index(0)

        # 현재 커서 위치의 테이블 속성 가져오기
        table = getter.get_current_table()
    """

    def __init__(self, hwp=None, new_instance: bool = False, visible: bool = True):
        """
        초기화

        Args:
            hwp: 기존 한글 객체 (None이면 새로 생성/연결)
            new_instance: True면 새 인스턴스 생성, False면 기존 인스턴스에 연결
            visible: 한글 창 표시 여부
        """
        self.hwp = hwp
        self._visible = visible
        self._new_instance = new_instance

        if self.hwp is None:
            self._init_hwp()

    def _init_hwp(self):
        """한글 COM 객체 초기화"""
        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError(
                "pywin32가 설치되어 있지 않습니다.\n"
                "pip install pywin32"
            )

        try:
            if self._new_instance:
                # 새 인스턴스 생성
                self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            else:
                # 기존 인스턴스에 연결 시도
                try:
                    self.hwp = win32.GetActiveObject("HWPFrame.HwpObject")
                except:
                    # 연결 실패 시 새로 생성
                    self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

            # 보안 모듈 등록
            self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

            # 창 표시 설정
            self.hwp.XHwpWindows.Item(0).Visible = self._visible

        except Exception as e:
            raise RuntimeError(f"한글 프로그램 연결 실패: {e}")

    def open(self, file_path: str, read_only: bool = False) -> bool:
        """
        한글 문서 열기

        Args:
            file_path: 파일 경로
            read_only: 읽기 전용 여부

        Returns:
            성공 여부
        """
        return self.hwp.Open(file_path, "HWP", "")

    def _get_ctrl_properties(self, ctrl) -> Dict[str, Any]:
        """컨트롤의 속성 추출"""
        props = {}

        try:
            # 기본 속성
            props['ctrl_id'] = ctrl.CtrlID if hasattr(ctrl, 'CtrlID') else None

            # Properties 객체에서 속성 추출
            if hasattr(ctrl, 'Properties'):
                ctrl_props = ctrl.Properties

                # 표 속성 항목들
                prop_names = [
                    'TreatAsChar',  # 글자처럼 취급
                    'Protect',  # 보호
                    'PageBreak',  # 페이지 나눔
                    'RepeatHeader',  # 제목 줄 반복
                    'RowCnt',  # 행 수
                    'ColCnt',  # 열 수
                    'CellSpacing',  # 셀 간격
                    'BorderFillIDRef',  # 테두리/채우기 참조
                ]

                for name in prop_names:
                    try:
                        props[name] = ctrl_props.Item(name)
                    except:
                        pass

        except Exception as e:
            pass

        return props

    def _get_table_cells(self, ctrl) -> List[List[CellInfo]]:
        """테이블의 모든 셀 정보 추출"""
        cells = []

        try:
            # 현재 위치 저장
            self.hwp.SetPos(ctrl.GetAnchorPos(0), 0, 0)

            # 표 선택
            self.hwp.FindCtrl()

            # 테이블 정보 가져오기
            table_set = self.hwp.HParameterSet.HTableCreation
            self.hwp.HAction.GetDefault("TablePropertyDialog", table_set.HSet)

            row_count = table_set.Rows
            col_count = table_set.Cols

            # 표 안으로 이동
            self.hwp.HAction.Run("TableCellBlock")
            self.hwp.HAction.Run("Cancel")

            # 각 셀 순회
            for row in range(row_count):
                row_cells = []
                for col in range(col_count):
                    # 해당 셀로 이동
                    self._move_to_cell(row, col)

                    # 셀 텍스트 추출
                    cell_text = self._get_cell_text()

                    cell_info = CellInfo(
                        row=row,
                        col=col,
                        text=cell_text,
                    )

                    row_cells.append(cell_info)

                cells.append(row_cells)

        except Exception as e:
            pass

        return cells

    def _move_to_cell(self, row: int, col: int):
        """지정된 셀로 이동"""
        try:
            # TableCellAddr 액션 사용
            param_set = self.hwp.HParameterSet.HTableCellAddr
            self.hwp.HAction.GetDefault("TableCellAddr", param_set.HSet)
            param_set.Row = row
            param_set.Col = col
            self.hwp.HAction.Execute("TableCellAddr", param_set.HSet)
        except:
            pass

    def _get_cell_text(self) -> str:
        """현재 셀의 텍스트 추출"""
        try:
            # 셀 전체 선택
            self.hwp.HAction.Run("SelectAll")

            # 선택된 텍스트 가져오기
            text = self.hwp.GetTextFile("TEXT", "")

            # 선택 해제
            self.hwp.HAction.Run("Cancel")

            return text.strip() if text else ""
        except:
            return ""

    def _extract_table_property(self, ctrl, table_index: int) -> TableProperty:
        """단일 테이블 컨트롤에서 속성 추출"""
        table_prop = TableProperty(table_index=table_index)
        table_prop.ctrl = ctrl

        try:
            # 기본 ID 정보
            if hasattr(ctrl, 'GetAnchorPos'):
                pos = ctrl.GetAnchorPos(0)
                # pos는 (List, Para, Pos) 튜플 형태
                if pos:
                    table_prop.para_id = pos[1] if len(pos) > 1 else None
                    table_prop.char_id = pos[2] if len(pos) > 2 else None

            # UserDesc에서 추가 정보 추출
            if hasattr(ctrl, 'UserDesc'):
                table_prop.ctrl_id = ctrl.UserDesc

            # CtrlID 추출
            if hasattr(ctrl, 'CtrlID'):
                table_prop.ctrl_id = ctrl.CtrlID

            # Properties에서 상세 속성 추출
            if hasattr(ctrl, 'Properties'):
                props = ctrl.Properties

                try:
                    table_prop.row_count = props.Item('RowCnt')
                except:
                    pass

                try:
                    table_prop.col_count = props.Item('ColCnt')
                except:
                    pass

                try:
                    table_prop.treat_as_char = bool(props.Item('TreatAsChar'))
                except:
                    pass

                try:
                    table_prop.protect = bool(props.Item('Protect'))
                except:
                    pass

                try:
                    table_prop.repeat_header = bool(props.Item('RepeatHeader'))
                except:
                    pass

                try:
                    table_prop.cell_spacing = props.Item('CellSpacing')
                except:
                    pass

                try:
                    table_prop.border_fill_id = props.Item('BorderFillIDRef')
                except:
                    pass

            # 크기 정보 (ShapeObject)
            if hasattr(ctrl, 'ShapeObject'):
                shape = ctrl.ShapeObject
                if shape:
                    try:
                        table_prop.width = shape.Width
                        table_prop.height = shape.Height
                    except:
                        pass

                    try:
                        table_prop.x_pos = shape.XPos
                        table_prop.y_pos = shape.YPos
                    except:
                        pass

                    try:
                        table_prop.left_margin = shape.LeftMargin
                        table_prop.right_margin = shape.RightMargin
                        table_prop.top_margin = shape.TopMargin
                        table_prop.bottom_margin = shape.BottomMargin
                    except:
                        pass

        except Exception as e:
            pass

        return table_prop

    def get_all_tables(self, include_cells: bool = False) -> List[TableProperty]:
        """
        문서의 모든 테이블 속성 가져오기

        Args:
            include_cells: True면 셀 내용도 포함

        Returns:
            TableProperty 리스트
        """
        tables = []
        table_index = 0

        try:
            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 컨트롤 순회
            ctrl = self.hwp.HeadCtrl

            while ctrl:
                # 표 컨트롤인지 확인
                if ctrl.CtrlID == "tbl":
                    table_prop = self._extract_table_property(ctrl, table_index)

                    if include_cells:
                        table_prop.cells = self._get_table_cells(ctrl)

                    tables.append(table_prop)
                    table_index += 1

                # 다음 컨트롤로
                ctrl = ctrl.Next

        except Exception as e:
            pass

        return tables

    def get_table_by_index(self, index: int, include_cells: bool = True) -> Optional[TableProperty]:
        """
        인덱스로 특정 테이블 속성 가져오기

        Args:
            index: 테이블 인덱스 (0부터 시작)
            include_cells: True면 셀 내용도 포함

        Returns:
            TableProperty 또는 None
        """
        tables = self.get_all_tables(include_cells=include_cells)

        if 0 <= index < len(tables):
            return tables[index]

        return None

    def get_current_table(self) -> Optional[TableProperty]:
        """
        현재 커서 위치의 테이블 속성 가져오기

        Returns:
            TableProperty 또는 None (커서가 표 안에 없으면)
        """
        try:
            # 현재 위치의 컨트롤 확인
            ctrl = self.hwp.ParentCtrl

            if ctrl and ctrl.CtrlID == "tbl":
                return self._extract_table_property(ctrl, -1)

            # 표 찾기 시도
            if self.hwp.HAction.Run("TableCellBlock"):
                ctrl = self.hwp.ParentCtrl
                if ctrl and ctrl.CtrlID == "tbl":
                    table_prop = self._extract_table_property(ctrl, -1)
                    self.hwp.HAction.Run("Cancel")
                    return table_prop

        except:
            pass

        return None

    def move_to_table(self, index: int) -> bool:
        """
        지정된 인덱스의 테이블로 이동

        Args:
            index: 테이블 인덱스 (0부터 시작)

        Returns:
            성공 여부
        """
        try:
            # 문서 처음으로
            self.hwp.HAction.Run("MoveDocBegin")

            # n번째 표까지 이동
            for i in range(index + 1):
                # FindCtrl로 다음 표 찾기
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                self.hwp.HParameterSet.HGotoE.SetItem("DialogResult", 31)  # 표
                self.hwp.HParameterSet.HGotoE.SetItem("SetItemCount", i + 1)

                if not self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet):
                    return False

            return True

        except:
            return False

    def get_table_as_dataframe(self, index: int = 0):
        """
        테이블을 pandas DataFrame으로 가져오기

        Args:
            index: 테이블 인덱스

        Returns:
            pandas DataFrame
        """
        table = self.get_table_by_index(index, include_cells=True)

        if table:
            return table.to_dataframe()

        return None

    def get_position_info(self) -> Dict[str, Any]:
        """
        현재 커서 위치 정보 가져오기

        Returns:
            위치 정보 딕셔너리 (list_id, para_id, char_id 포함)
        """
        try:
            pos = self.hwp.GetPos()
            return {
                'list_id': pos[0] if len(pos) > 0 else None,
                'para_id': pos[1] if len(pos) > 1 else None,
                'char_id': pos[2] if len(pos) > 2 else None,
            }
        except:
            return {'list_id': None, 'para_id': None, 'char_id': None}

    def get_ctrl_at_position(self) -> Optional[Dict[str, Any]]:
        """
        현재 위치의 컨트롤 정보 가져오기

        Returns:
            컨트롤 정보 딕셔너리
        """
        try:
            ctrl = self.hwp.ParentCtrl
            if ctrl:
                return {
                    'ctrl_id': ctrl.CtrlID,
                    'ctrl': ctrl,
                    'user_desc': ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else None,
                }
        except:
            pass

        return None


# 편의 함수
def get_tables_from_file(file_path: str, include_cells: bool = False) -> List[TableProperty]:
    """
    한글 파일에서 모든 테이블 속성 추출 (편의 함수)

    Args:
        file_path: 한글 파일 경로
        include_cells: 셀 내용 포함 여부

    Returns:
        TableProperty 리스트
    """
    getter = GetTableProperty(visible=False)
    getter.open(file_path)
    return getter.get_all_tables(include_cells=include_cells)


def get_table_data_as_list(file_path: str, table_index: int = 0) -> List[List[str]]:
    """
    한글 파일에서 특정 테이블의 데이터를 2D 리스트로 추출 (편의 함수)

    Args:
        file_path: 한글 파일 경로
        table_index: 테이블 인덱스

    Returns:
        2D 문자열 리스트
    """
    getter = GetTableProperty(visible=False)
    getter.open(file_path)
    table = getter.get_table_by_index(table_index, include_cells=True)

    if table:
        return table.get_data_as_2d_list()

    return []


# 사용 예제
if __name__ == "__main__":
    print("=" * 60)
    print("한글 COM API 테이블 속성 추출 클래스")
    print("=" * 60)
    print()
    print("사용 예제:")
    print()
    print("```python")
    print("from get_table_property import GetTableProperty")
    print()
    print("# 1. 한글 프로그램에 연결")
    print("getter = GetTableProperty()")
    print()
    print("# 2. 파일 열기")
    print('getter.open("document.hwp")')
    print()
    print("# 3. 모든 테이블 속성 가져오기")
    print("tables = getter.get_all_tables()")
    print("for table in tables:")
    print("    print(f'테이블 {table.table_index}:')")
    print("    print(f'  - para_id: {table.para_id}')")
    print("    print(f'  - char_id: {table.char_id}')")
    print("    print(f'  - ctrl_id: {table.ctrl_id}')")
    print("    print(f'  - 행 수: {table.row_count}')")
    print("    print(f'  - 열 수: {table.col_count}')")
    print()
    print("# 4. 특정 테이블 속성 (셀 포함)")
    print("table = getter.get_table_by_index(0, include_cells=True)")
    print("print(table.get_data_as_2d_list())")
    print()
    print("# 5. DataFrame으로 변환")
    print("df = table.to_dataframe()")
    print()
    print("# 6. 현재 커서 위치의 테이블")
    print("current_table = getter.get_current_table()")
    print()
    print("# 7. 위치 정보 조회")
    print("pos_info = getter.get_position_info()")
    print("print(f'para_id: {pos_info[\"para_id\"]}, char_id: {pos_info[\"char_id\"]}')")
    print("```")
    print()
    print("주요 속성:")
    print("  - link_id: 연결 ID")
    print("  - para_id: 문단 ID")
    print("  - char_id: 문자 ID")
    print("  - ctrl_id: 컨트롤 ID")
    print("  - ctrl: 컨트롤 객체 (COM 객체)")
    print("  - row_count / col_count: 행/열 수")
    print("  - width / height: 크기")
    print("  - treat_as_char: 글자처럼 취급")
    print("  - repeat_header: 제목 줄 반복")
