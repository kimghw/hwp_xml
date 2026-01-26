# -*- coding: utf-8 -*-
"""
한글 COM API를 사용한 문단/글자 스타일 추출

모든 list_id의 para_id별 문단 스타일, 글자 스타일 정보를 추출합니다.
"""

import sys
import json
import yaml
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any


@dataclass
class CharStyle:
    """글자 스타일 정보"""
    font_name: Optional[str] = None  # 글꼴 이름
    font_size: Optional[float] = None  # 글자 크기 (pt)
    bold: bool = False  # 굵게
    italic: bool = False  # 기울임
    underline: bool = False  # 밑줄
    strikeout: bool = False  # 취소선
    text_color: Optional[str] = None  # 글자색 (RGB hex)
    highlight_color: Optional[str] = None  # 형광펜 색상

    def to_dict(self) -> Dict[str, Any]:
        return {
            'font_name': self.font_name,
            'font_size': self.font_size,
            'bold': self.bold,
            'italic': self.italic,
            'underline': self.underline,
            'strikeout': self.strikeout,
            'text_color': self.text_color,
            'highlight_color': self.highlight_color,
        }


@dataclass
class ParaStyle:
    """문단 스타일 정보"""
    # 위치 정보
    list_id: int = 0
    para_id: int = 0

    # 문단 내용
    text: str = ""
    line_count: int = 1  # 줄 수

    # char_id 정보
    start_char_id: int = 0  # 시작 char_id
    next_line_char_id: Optional[int] = None  # 다음 줄 시작 char_id (1줄이면 None)
    end_char_id: int = 0  # 마지막 char_id

    # 문단 스타일 이름
    style_name: Optional[str] = None

    # 정렬
    align: Optional[str] = None  # 왼쪽, 가운데, 오른쪽, 양쪽, 배분

    # 들여쓰기/내어쓰기
    indent: float = 0  # 들여쓰기 (HWPUNIT)
    margin_left: float = 0  # 왼쪽 여백
    margin_right: float = 0  # 오른쪽 여백

    # 줄 간격
    line_spacing: float = 0  # 줄 간격
    line_spacing_type: Optional[str] = None  # 퍼센트, 고정값 등

    # 문단 간격
    space_before: float = 0  # 문단 앞 간격
    space_after: float = 0  # 문단 뒤 간격

    # 글자 스타일 (대표)
    char_style: Optional[CharStyle] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            'list_id': self.list_id,
            'para_id': self.para_id,
            'text': self.text,
            'line_count': self.line_count,
            'start_char_id': self.start_char_id,
            'next_line_char_id': self.next_line_char_id,
            'end_char_id': self.end_char_id,
            'style_name': self.style_name,
            'align': self.align,
            'indent': self.indent,
            'margin_left': self.margin_left,
            'margin_right': self.margin_right,
            'line_spacing': self.line_spacing,
            'line_spacing_type': self.line_spacing_type,
            'space_before': self.space_before,
            'space_after': self.space_after,
            'char_style': self.char_style.to_dict() if self.char_style else None,
        }


class GetParaStyle:
    """
    한글 COM API를 사용하여 문단/글자 스타일을 추출하는 클래스
    """

    # 정렬 타입 매핑
    ALIGN_TYPES = {
        0: "양쪽",
        1: "왼쪽",
        2: "오른쪽",
        3: "가운데",
        4: "배분",
        5: "나눔",
    }

    # 줄간격 타입 매핑
    LINE_SPACING_TYPES = {
        0: "퍼센트",
        1: "고정값",
        2: "여백만",
        3: "최소",
    }

    def __init__(self, hwp=None):
        """
        초기화

        Args:
            hwp: 기존 한글 객체 (None이면 연결/생성 시도)
        """
        self.hwp = hwp
        if self.hwp is None:
            self._init_hwp()

    def _init_hwp(self):
        """한글 COM 객체 초기화"""
        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError("pywin32가 설치되어 있지 않습니다.")

        # 먼저 열린 한글에 연결 시도
        try:
            self.hwp = win32.GetActiveObject("HWPFrame.HwpObject")
            print("열린 한글 문서에 연결됨")
        except:
            # 연결 실패 시 파일 선택 대화상자
            self.hwp = None

    def _open_file_dialog(self) -> Optional[str]:
        """파일 선택 대화상자 열기"""
        try:
            import win32com.client as win32
            from win32com.shell import shell, shellcon
            import win32gui

            # 파일 열기 대화상자
            filter_str = "한글 파일 (*.hwp;*.hwpx)\0*.hwp;*.hwpx\0모든 파일 (*.*)\0*.*\0"

            try:
                filename, customfilter, flags = win32gui.GetOpenFileNameW(
                    Filter=filter_str,
                    Title="한글 파일 선택"
                )
                return filename
            except:
                return None
        except:
            return None

    def open(self, file_path: str) -> bool:
        """한글 문서 열기"""
        if self.hwp is None:
            import win32com.client as win32
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            self.hwp.XHwpWindows.Item(0).Visible = True

        return self.hwp.Open(file_path, "HWP", "")

    def _get_position(self) -> tuple:
        """현재 커서 위치 반환 (list_id, para_id, char_id)"""
        try:
            pos = self.hwp.GetPos()
            return pos[0], pos[1], pos[2]
        except:
            return 0, 0, 0

    def _rgb_to_hex(self, color_value: int) -> str:
        """색상값을 hex 문자열로 변환"""
        if color_value is None or color_value < 0:
            return None
        r = color_value & 0xFF
        g = (color_value >> 8) & 0xFF
        b = (color_value >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"

    def _get_para_shape(self) -> Dict[str, Any]:
        """현재 문단의 모양 정보 추출"""
        result = {}

        try:
            # 문단 모양 대화상자 값 가져오기
            act = self.hwp.CreateAction("ParagraphShape")
            pset = act.CreateSet()
            act.GetDefault(pset)

            # 정렬
            align_val = pset.Item("Align")
            result['align'] = self.ALIGN_TYPES.get(align_val, str(align_val))

            # 들여쓰기/여백
            result['indent'] = pset.Item("Indent")  # 첫줄 들여쓰기
            result['margin_left'] = pset.Item("LeftMargin")
            result['margin_right'] = pset.Item("RightMargin")

            # 줄 간격
            result['line_spacing'] = pset.Item("LineSpacing")
            ls_type = pset.Item("LineSpacingType")
            result['line_spacing_type'] = self.LINE_SPACING_TYPES.get(ls_type, str(ls_type))

            # 문단 간격
            result['space_before'] = pset.Item("SpaceBeforePara")
            result['space_after'] = pset.Item("SpaceAfterPara")

        except Exception as e:
            pass

        return result

    def _get_char_shape(self) -> CharStyle:
        """현재 위치의 글자 모양 정보 추출"""
        char_style = CharStyle()

        try:
            # 글자 모양 대화상자 값 가져오기
            act = self.hwp.CreateAction("CharShape")
            pset = act.CreateSet()
            act.GetDefault(pset)

            # 글꼴
            try:
                face_name = pset.Item("FaceNameHangul")
                char_style.font_name = face_name
            except:
                pass

            # 글자 크기 (HWPUNIT -> pt 변환: /100)
            try:
                height = pset.Item("Height")
                char_style.font_size = height / 100.0 if height else None
            except:
                pass

            # 굵게/기울임
            try:
                char_style.bold = bool(pset.Item("Bold"))
            except:
                pass

            try:
                char_style.italic = bool(pset.Item("Italic"))
            except:
                pass

            # 밑줄
            try:
                underline_type = pset.Item("UnderlineType")
                char_style.underline = underline_type > 0
            except:
                pass

            # 취소선
            try:
                strikeout_type = pset.Item("StrikeOutType")
                char_style.strikeout = strikeout_type > 0
            except:
                pass

            # 글자색
            try:
                text_color = pset.Item("TextColor")
                char_style.text_color = self._rgb_to_hex(text_color)
            except:
                pass

            # 형광펜
            try:
                highlight = pset.Item("HighlightColor")
                if highlight and highlight != -1:
                    char_style.highlight_color = self._rgb_to_hex(highlight)
            except:
                pass

        except Exception as e:
            pass

        return char_style

    def _get_para_text(self) -> str:
        """현재 문단의 텍스트 추출 (해당 문단만)"""
        try:
            # 현재 위치 저장
            orig_list, orig_para, _ = self._get_position()

            # 문단 시작으로 이동
            self.hwp.HAction.Run("MoveParaBegin")

            # 문단 끝까지 블록 선택 (Shift+End 방식)
            self.hwp.HAction.Run("MoveSelParaEnd")

            # 선택된 텍스트 가져오기
            text = self.hwp.GetTextFile("TEXT", "saveblock")

            # 선택 해제
            self.hwp.HAction.Run("Cancel")

            # 원래 위치로 복귀
            self.hwp.SetPos(orig_list, orig_para, 0)

            # 줄바꿈을 공백으로, 앞뒤 공백 제거
            if text:
                text = text.replace("\r\n", " ").replace("\n", " ").strip()
            return text or ""
        except:
            return ""

    def _get_para_line_info(self) -> dict:
        """
        현재 문단의 줄 수 및 char_id 정보 계산

        Returns:
            {
                'line_count': 줄 수,
                'start_char_id': 시작 char_id,
                'next_line_char_id': 다음 줄 시작 char_id (1줄이면 None),
                'end_char_id': 마지막 char_id
            }

        Note:
            - MoveLineBegin: 현재 줄의 시작으로 이동
            - MoveLineEnd: 현재 줄의 끝으로 이동
            - MoveDown: 아래 줄로 이동 (같은 문단 내에서)
        """
        result = {
            'line_count': 1,
            'start_char_id': 0,
            'next_line_char_id': None,
            'end_char_id': 0
        }

        try:
            # 현재 위치 저장
            orig_list, orig_para, _ = self._get_position()

            # 문단 끝으로 이동해서 마지막 char_id 확인
            self.hwp.HAction.Run("MoveParaEnd")
            _, _, end_char = self._get_position()
            result['end_char_id'] = end_char

            # 문단 시작으로 이동
            self.hwp.HAction.Run("MoveParaBegin")
            _, _, start_char = self._get_position()
            result['start_char_id'] = start_char

            # 줄 시작으로 이동해서 첫 줄 시작 char_id 확인
            self.hwp.HAction.Run("MoveLineBegin")
            _, _, line_start_char = self._get_position()
            # 문단 시작이 줄 시작보다 작으면 문단 시작 사용
            if start_char < line_start_char:
                result['start_char_id'] = start_char
            else:
                result['start_char_id'] = line_start_char

            line_count = 1
            next_line_char_id = None

            max_lines = 100  # 무한루프 방지
            prev_char = line_start_char

            for i in range(max_lines):
                # 아래 줄로 이동
                self.hwp.HAction.Run("MoveDown")
                curr_list, curr_para, curr_char = self._get_position()

                # para_id가 바뀌면 다른 문단으로 넘어간 것
                if curr_para != orig_para or curr_list != orig_list:
                    break

                # char_id가 변하지 않으면 마지막 줄
                if curr_char == prev_char:
                    break

                # char_id가 end_char를 넘어가면 종료
                if curr_char > end_char:
                    break

                line_count += 1

                # 줄 시작으로 이동해서 정확한 줄 시작 char_id 확인
                self.hwp.HAction.Run("MoveLineBegin")
                _, _, line_begin_char = self._get_position()

                # 첫 번째 다음 줄의 시작 char_id 저장
                if i == 0:
                    next_line_char_id = line_begin_char

                prev_char = line_begin_char

            result['line_count'] = line_count
            result['next_line_char_id'] = next_line_char_id

            # 다시 문단 시작으로 복귀
            self.hwp.SetPos(orig_list, orig_para, 0)

            return result

        except Exception as e:
            return result

    def _get_style_name(self) -> Optional[str]:
        """현재 문단의 스타일 이름 추출"""
        try:
            # 스타일 정보 가져오기
            act = self.hwp.CreateAction("Style")
            pset = act.CreateSet()
            act.GetDefault(pset)

            style_name = pset.Item("Name")
            return style_name if style_name else None
        except:
            return None

    def _extract_para_at_current_pos(self, visited: set) -> Optional[ParaStyle]:
        """현재 위치에서 문단 스타일 추출"""
        list_id, para_id, _ = self._get_position()
        pos_key = (list_id, para_id)

        if pos_key in visited:
            return None

        visited.add(pos_key)

        # 문단 스타일 정보 추출
        para_style = ParaStyle(list_id=list_id, para_id=para_id)

        # 문단 텍스트
        para_style.text = self._get_para_text()

        # 줄 수 및 char_id 정보 계산
        line_info = self._get_para_line_info()
        para_style.line_count = line_info['line_count']
        para_style.start_char_id = line_info['start_char_id']
        para_style.next_line_char_id = line_info['next_line_char_id']
        para_style.end_char_id = line_info['end_char_id']

        # 스타일 이름
        para_style.style_name = self._get_style_name()

        # 문단 모양
        para_shape = self._get_para_shape()
        para_style.align = para_shape.get('align')
        para_style.indent = para_shape.get('indent', 0)
        para_style.margin_left = para_shape.get('margin_left', 0)
        para_style.margin_right = para_shape.get('margin_right', 0)
        para_style.line_spacing = para_shape.get('line_spacing', 0)
        para_style.line_spacing_type = para_shape.get('line_spacing_type')
        para_style.space_before = para_shape.get('space_before', 0)
        para_style.space_after = para_shape.get('space_after', 0)

        # 글자 모양 (문단 시작 위치 기준)
        self.hwp.HAction.Run("MoveParaBegin")
        para_style.char_style = self._get_char_shape()

        return para_style

    def _traverse_all_paras(self, visited: set) -> List[ParaStyle]:
        """현재 위치부터 모든 문단 순회"""
        para_styles = []
        max_iterations = 10000
        iteration = 0

        while iteration < max_iterations:
            iteration += 1

            list_id, para_id, _ = self._get_position()
            pos_key = (list_id, para_id)

            # 이미 방문한 경우 다음으로 이동
            if pos_key in visited:
                prev_pos = pos_key

                self.hwp.HAction.Run("MoveNextPara")
                new_list_id, new_para_id, _ = self._get_position()

                if (new_list_id, new_para_id) == prev_pos:
                    self.hwp.HAction.Run("MoveRight")
                    new_list_id, new_para_id, _ = self._get_position()
                    if (new_list_id, new_para_id) == prev_pos:
                        break
                continue

            # 문단 추출
            para_style = self._extract_para_at_current_pos(visited)
            if para_style:
                para_styles.append(para_style)

            # 다음 문단으로 이동
            prev_pos = pos_key
            self.hwp.HAction.Run("MoveNextPara")

            new_list_id, new_para_id, _ = self._get_position()
            if (new_list_id, new_para_id) == prev_pos:
                self.hwp.HAction.Run("MoveRight")
                new_list_id, new_para_id, _ = self._get_position()
                if (new_list_id, new_para_id) == prev_pos:
                    break

        return para_styles

    def get_all_para_styles(self) -> List[ParaStyle]:
        """
        문서의 모든 문단 스타일 정보 추출 (테이블 셀 포함)

        list_id를 순차적으로 조회하여 모든 문단 추출

        Returns:
            ParaStyle 리스트
        """
        if self.hwp is None:
            raise RuntimeError("한글 문서가 열려있지 않습니다.")

        para_styles = []
        visited = set()  # (list_id, para_id) 중복 방지

        # list_id를 0부터 순차적으로 조회
        max_list_id = 1000  # 충분히 큰 값
        consecutive_failures = 0

        for list_id in range(max_list_id):
            # 해당 list_id로 이동 시도
            try:
                self.hwp.SetPos(list_id, 0, 0)
                curr_list, curr_para, _ = self._get_position()

                # 이동 성공 확인
                if curr_list != list_id:
                    consecutive_failures += 1
                    if consecutive_failures > 50:
                        break
                    continue

                consecutive_failures = 0

                # 해당 list_id의 모든 문단 추출
                for para_idx in range(100):
                    self.hwp.SetPos(list_id, para_idx, 0)
                    actual_list, actual_para, _ = self._get_position()

                    # list_id가 바뀌면 해당 list_id 끝
                    if actual_list != list_id:
                        break

                    # para_id가 다르면 해당 para_id 없음
                    if actual_para != para_idx:
                        continue

                    ps = self._extract_para_at_current_pos(visited)
                    if ps:
                        para_styles.append(ps)

            except Exception as e:
                consecutive_failures += 1
                if consecutive_failures > 50:
                    break

        print(f"  총 {len(visited)}개 문단 추출")

        # list_id 순으로 정렬
        para_styles.sort(key=lambda x: (x.list_id, x.para_id))

        return para_styles

    def get_para_styles_by_list_id(self) -> Dict[int, List[ParaStyle]]:
        """
        list_id별로 그룹화된 문단 스타일 정보 반환

        Returns:
            {list_id: [ParaStyle, ...], ...}
        """
        all_styles = self.get_all_para_styles()

        grouped = {}
        for style in all_styles:
            if style.list_id not in grouped:
                grouped[style.list_id] = []
            grouped[style.list_id].append(style)

        return grouped

    def to_json(self, para_styles: List[ParaStyle]) -> str:
        """ParaStyle 리스트를 JSON 문자열로 변환"""
        data = [ps.to_dict() for ps in para_styles]
        return json.dumps(data, ensure_ascii=False, indent=2)

    def to_yaml(self, para_styles: List[ParaStyle]) -> str:
        """ParaStyle 리스트를 YAML 문자열로 변환 (list_id별 그룹화, 주석 포함)"""
        # list_id별로 그룹화
        grouped = {}
        for ps in para_styles:
            if ps.list_id not in grouped:
                grouped[ps.list_id] = []
            # list_id는 키로 사용하므로 개별 항목에서 제외
            item = ps.to_dict()
            del item['list_id']
            grouped[ps.list_id].append(item)

        # 헤더 주석
        header = """# 한글 문서 문단 스타일 정보
#
# 구조:
#   list_id:                # 리스트 ID (본문=0, 테이블 셀=각각 고유 ID)
#     - para_id:            # 문단 ID (해당 list_id 내 문단 순번)
#       line_count:         # 줄 수
#       start_char_id:      # 시작 char_id
#       next_line_char_id:  # 다음 줄 시작 char_id (1줄이면 null)
#       end_char_id:        # 마지막 char_id
#       style_name:         # 문단 스타일 이름
#       align:              # 정렬 (왼쪽/가운데/오른쪽/양쪽/배분)
#       indent:             # 첫줄 들여쓰기 (HWPUNIT)
#       margin_left:        # 왼쪽 여백
#       margin_right:       # 오른쪽 여백
#       line_spacing:       # 줄 간격
#       line_spacing_type:  # 줄 간격 타입 (퍼센트/고정값/여백만/최소)
#       space_before:       # 문단 앞 간격
#       space_after:        # 문단 뒤 간격
#       char_style:         # 글자 스타일
#         font_name:        # 글꼴
#         font_size:        # 크기 (pt)
#         bold:             # 굵게
#         italic:           # 기울임
#         underline:        # 밑줄
#         strikeout:        # 취소선
#         text_color:       # 글자색 (RGB hex)
#         highlight_color:  # 형광펜 색상

"""
        yaml_content = yaml.dump(grouped, allow_unicode=True, default_flow_style=False, sort_keys=False)
        return header + yaml_content


try:
    from hwp_utils import get_hwp_instance, open_file_dialog, get_active_filepath, create_hwp_instance
except ImportError:
    from win32.hwp_utils import get_hwp_instance, open_file_dialog, get_active_filepath, create_hwp_instance


def open_file_dialog_win32() -> Optional[str]:
    """Windows 파일 선택 대화상자 (하위 호환용)"""
    return open_file_dialog()


def main():
    """메인 함수"""
    import os

    print("=" * 60)
    print("한글 문단/글자 스타일 추출")
    print("=" * 60)
    print()

    # 열린 한글 인스턴스 확인
    hwp = get_hwp_instance()
    filepath = None

    if hwp:
        print("열린 한글 문서를 사용합니다.")
        filepath = get_active_filepath(hwp)
        getter = GetParaStyle(hwp)
    else:
        print("열린 한글 문서가 없습니다. 파일을 선택하세요.")

        filepath = open_file_dialog_win32()
        if not filepath:
            print("파일 선택 취소")
            return

        print(f"선택된 파일: {filepath}")

        getter = GetParaStyle()
        if not getter.open(filepath):
            print("파일 열기 실패")
            return

    print()
    print("문단 스타일 추출 중...")
    print()

    # 문단 스타일 추출
    para_styles = getter.get_all_para_styles()

    print(f"총 {len(para_styles)}개 문단 추출됨")
    print()
    print("-" * 60)

    # 결과 출력
    for ps in para_styles:
        print(f"[list_id={ps.list_id}, para_id={ps.para_id}]")
        print(f"  스타일: {ps.style_name or '(없음)'}")
        print(f"  줄 수: {ps.line_count}")
        print(f"  정렬: {ps.align}")
        print(f"  들여쓰기: {ps.indent}, 왼쪽여백: {ps.margin_left}, 오른쪽여백: {ps.margin_right}")
        print(f"  줄간격: {ps.line_spacing} ({ps.line_spacing_type})")
        print(f"  문단간격: 앞 {ps.space_before}, 뒤 {ps.space_after}")

        if ps.char_style:
            cs = ps.char_style
            print(f"  글자: {cs.font_name}, {cs.font_size}pt", end="")
            if cs.bold: print(", 굵게", end="")
            if cs.italic: print(", 기울임", end="")
            if cs.underline: print(", 밑줄", end="")
            if cs.strikeout: print(", 취소선", end="")
            if cs.text_color: print(f", 색상={cs.text_color}", end="")
            print()

        text_preview = ps.text[:50] + "..." if len(ps.text) > 50 else ps.text
        print(f"  내용: {text_preview}")
        print()

    # YAML 파일로 저장 (파일명_para.yaml)
    if filepath:
        output_path = os.path.splitext(filepath)[0] + "_para.yaml"
    else:
        output_path = r"C:\hwp_xml\win32\para_styles.yaml"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(getter.to_yaml(para_styles))
    print(f"YAML 저장: {output_path}")


if __name__ == "__main__":
    main()
