# -*- coding: utf-8 -*-
"""
테이블/이미지 객체 서식 설정 모듈

HWPX 파일에서 테이블과 이미지 객체의 위치, 정렬, 글자처럼 취급 등을 설정합니다.

## 가운데 정렬 원리
한글(HWP)에서 테이블/이미지를 가운데 정렬하려면 다음 속성들을 설정해야 합니다:

1. 객체 속성 (hp:pos):
   - treatAsChar: 1 ✅ (글자처럼 취급)
   - horzAlign: CENTER ✅ (수평 정렬)

2. 문단 속성 (hp:p):
   - paraPrIDRef: 20 ✅ (가운데 정렬 스타일 참조, header.xml에 정의됨)
   - paraShape: 없음 ✅ (인라인 paraShape 대신 paraPrIDRef 사용)

주의: 인라인 <paraShape align="CENTER" />는 한글에서 인식되지 않습니다.
      반드시 paraPrIDRef를 통해 header.xml의 스타일을 참조해야 합니다.

사용 예:
    from merge.formatters import ObjectFormatter

    formatter = ObjectFormatter()

    # 테이블을 글자처럼 취급 + 가운데 정렬
    formatter.set_table_format("document.hwpx", treat_as_char=True, h_align="CENTER")

    # 모든 객체(테이블+이미지) 서식 일괄 적용
    formatter.set_all_format("document.hwpx", treat_as_char=True, h_align="CENTER")
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Literal
from io import BytesIO
import shutil


# XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

# 유효한 정렬 값
HAlign = Literal["LEFT", "CENTER", "RIGHT"]
VAlign = Literal["TOP", "CENTER", "BOTTOM"]


class ObjectFormatter:
    """
    테이블/이미지 객체 서식 설정기

    HWPX 파일에서 테이블(<tbl>)과 이미지(<pic>) 객체의 위치와 정렬을 설정합니다.
    """

    def __init__(self):
        """초기화"""
        pass

    def set_treat_as_char(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        output_path: Optional[str] = None,
        element_type: Optional[str] = None,
        h_align: Optional[HAlign] = None,
        v_align: Optional[VAlign] = None
    ) -> str:
        """
        HWPX 파일의 테이블/이미지를 '글자처럼 취급' 설정

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            treat_as_char: True면 글자처럼 취급, False면 어울림(앵커)
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            element_type: 대상 요소 ("table", "image", None이면 전체)
            h_align: 수평 정렬 ("LEFT", "CENTER", "RIGHT")
            v_align: 수직 정렬 ("TOP", "CENTER", "BOTTOM")

        Returns:
            저장된 파일 경로
        """
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        if not hwpx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwpx_path}")

        # 네임스페이스 등록
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

        temp_path = hwpx_path.with_suffix('.hwpx.tmp')

        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf_in:
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for item in zf_in.namelist():
                        content = zf_in.read(item)

                        if item.startswith('Contents/section') and item.endswith('.xml'):
                            content = self._change_object_format(
                                content, treat_as_char, element_type, h_align, v_align
                            )

                        zf_out.writestr(item, content)

            if output_path != hwpx_path:
                shutil.copy2(temp_path, output_path)
            else:
                shutil.move(temp_path, output_path)

            return str(output_path)

        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _change_object_format(
        self,
        xml_content: bytes,
        treat_as_char: bool,
        element_type: Optional[str] = None,
        h_align: Optional[HAlign] = None,
        v_align: Optional[VAlign] = None
    ) -> bytes:
        """섹션 XML의 객체 서식 변경"""
        root = ET.parse(BytesIO(xml_content)).getroot()
        modified = False
        treat_as_char_value = "1" if treat_as_char else "0"

        # 대상 태그 결정
        target_tags = []
        if element_type == "table":
            target_tags = ['}tbl']
        elif element_type == "image":
            target_tags = ['}pic']
        else:
            target_tags = ['}tbl', '}pic']

        # 부모-자식 맵 생성 (문단 정렬을 위해)
        parent_map = {c: p for p in root.iter() for c in p}

        for elem in root.iter():
            # 테이블 또는 이미지 요소 확인
            is_target = any(elem.tag.endswith(tag) for tag in target_tags)
            if not is_target:
                continue

            # hp:pos 요소 찾기
            for child in elem:
                if child.tag.endswith('}pos'):
                    # treatAsChar 설정
                    current_treat_as_char = child.get('treatAsChar', '0')
                    if current_treat_as_char != treat_as_char_value:
                        child.set('treatAsChar', treat_as_char_value)
                        modified = True

                    # 수평 정렬 설정 (horzAlign과 hAlign 모두 설정)
                    if h_align:
                        # horzAlign 속성 변경
                        current_horz_align = child.get('horzAlign', '')
                        if current_horz_align != h_align:
                            child.set('horzAlign', h_align)
                            modified = True

                        # hAlign 속성도 변경 (호환성)
                        current_h_align = child.get('hAlign', '')
                        if current_h_align != h_align:
                            child.set('hAlign', h_align)
                            modified = True

                    # 수직 정렬 설정 (vertAlign과 vAlign 모두 설정)
                    if v_align:
                        # vertAlign 속성 변경
                        current_vert_align = child.get('vertAlign', '')
                        if current_vert_align != v_align:
                            child.set('vertAlign', v_align)
                            modified = True

                        # vAlign 속성도 변경 (호환성)
                        current_v_align = child.get('vAlign', '')
                        if current_v_align != v_align:
                            child.set('vAlign', v_align)
                            modified = True
                    break

            # 글자처럼 취급일 때 문단 정렬 설정
            if treat_as_char and h_align:
                para_modified = self._set_paragraph_align(elem, parent_map, h_align)
                if para_modified:
                    modified = True

        if modified:
            return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
        return xml_content

    def _set_paragraph_align(self, obj_elem, parent_map, align: HAlign) -> bool:
        """
        객체(테이블/이미지)를 포함한 문단의 정렬 설정

        Args:
            obj_elem: 테이블 또는 이미지 요소
            parent_map: 부모 맵
            align: 정렬 값

        Returns:
            수정 여부
        """
        modified = False

        # 객체의 부모 문단(p) 찾기
        current = obj_elem
        while current is not None:
            current = parent_map.get(current)
            if current is not None and current.tag.endswith('}p'):
                # 문단의 paraShape 찾기 또는 생성
                para_shape = None
                for child in current:
                    if child.tag.endswith('}paraShape'):
                        para_shape = child
                        break

                # 가운데 정렬용 paraPrIDRef 설정
                # CENTER 정렬은 paraPrIDRef=20 사용 (header.xml에 정의됨)
                if align == "CENTER":
                    para_pr_id = "20"
                elif align == "LEFT":
                    para_pr_id = "0"
                elif align == "RIGHT":
                    para_pr_id = "0"  # RIGHT용 ID가 필요하면 추가
                else:
                    para_pr_id = "0"

                current_para_pr = current.get('paraPrIDRef', '')
                if current_para_pr != para_pr_id:
                    current.set('paraPrIDRef', para_pr_id)
                    modified = True

                # paraShape가 있으면 제거 (paraPrIDRef 사용)
                if para_shape is not None:
                    current.remove(para_shape)
                    modified = True

                break

        return modified

    def set_table_format(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        h_align: HAlign = "CENTER",
        v_align: Optional[VAlign] = None,
        output_path: Optional[str] = None
    ) -> str:
        """
        테이블 서식 설정 (글자처럼 취급 + 정렬)

        Args:
            hwpx_path: HWPX 파일 경로
            treat_as_char: 글자처럼 취급 여부
            h_align: 수평 정렬 (기본: CENTER)
            v_align: 수직 정렬
            output_path: 저장 경로

        Returns:
            저장된 파일 경로
        """
        return self.set_treat_as_char(
            hwpx_path, treat_as_char, output_path, "table", h_align, v_align
        )

    def set_image_format(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        h_align: HAlign = "CENTER",
        v_align: Optional[VAlign] = None,
        output_path: Optional[str] = None
    ) -> str:
        """
        이미지 서식 설정 (글자처럼 취급 + 정렬)

        Args:
            hwpx_path: HWPX 파일 경로
            treat_as_char: 글자처럼 취급 여부
            h_align: 수평 정렬 (기본: CENTER)
            v_align: 수직 정렬
            output_path: 저장 경로

        Returns:
            저장된 파일 경로
        """
        return self.set_treat_as_char(
            hwpx_path, treat_as_char, output_path, "image", h_align, v_align
        )

    def set_all_format(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        h_align: HAlign = "CENTER",
        v_align: Optional[VAlign] = None,
        output_path: Optional[str] = None
    ) -> str:
        """
        모든 테이블/이미지 서식 일괄 설정

        Args:
            hwpx_path: HWPX 파일 경로
            treat_as_char: 글자처럼 취급 여부 (기본: True)
            h_align: 수평 정렬 (기본: CENTER)
            v_align: 수직 정렬
            output_path: 저장 경로

        Returns:
            저장된 파일 경로
        """
        return self.set_treat_as_char(
            hwpx_path, treat_as_char, output_path, None, h_align, v_align
        )

    # ========== 편의 메서드 ==========

    def set_table_as_char_center(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """테이블을 글자처럼 취급 + 가운데 정렬"""
        return self.set_table_format(hwpx_path, True, "CENTER", None, output_path)

    def set_image_as_char_center(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """이미지를 글자처럼 취급 + 가운데 정렬"""
        return self.set_image_format(hwpx_path, True, "CENTER", None, output_path)

    def set_all_as_char_center(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """모든 테이블/이미지를 글자처럼 취급 + 가운데 정렬"""
        return self.set_all_format(hwpx_path, True, "CENTER", None, output_path)

    def set_table_left_align(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        output_path: Optional[str] = None
    ) -> str:
        """테이블을 왼쪽 정렬"""
        return self.set_table_format(hwpx_path, treat_as_char, "LEFT", None, output_path)

    def set_table_right_align(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        output_path: Optional[str] = None
    ) -> str:
        """테이블을 오른쪽 정렬"""
        return self.set_table_format(hwpx_path, treat_as_char, "RIGHT", None, output_path)

    def set_alignment_only(
        self,
        hwpx_path: str,
        h_align: HAlign = "CENTER",
        v_align: Optional[VAlign] = None,
        element_type: Optional[str] = None,
        output_path: Optional[str] = None
    ) -> str:
        """
        정렬만 설정 (treatAsChar는 변경하지 않음)

        Args:
            hwpx_path: HWPX 파일 경로
            h_align: 수평 정렬
            v_align: 수직 정렬
            element_type: "table", "image", None(전체)
            output_path: 저장 경로

        Returns:
            저장된 파일 경로
        """
        hwpx_path = Path(hwpx_path)
        output_path = Path(output_path) if output_path else hwpx_path

        if not hwpx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwpx_path}")

        # 네임스페이스 등록
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)

        temp_path = hwpx_path.with_suffix('.hwpx.tmp')

        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf_in:
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for item in zf_in.namelist():
                        content = zf_in.read(item)

                        if item.startswith('Contents/section') and item.endswith('.xml'):
                            content = self._change_alignment_only(
                                content, h_align, v_align, element_type
                            )

                        zf_out.writestr(item, content)

            if output_path != hwpx_path:
                shutil.copy2(temp_path, output_path)
            else:
                shutil.move(temp_path, output_path)

            return str(output_path)

        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _change_alignment_only(
        self,
        xml_content: bytes,
        h_align: HAlign,
        v_align: Optional[VAlign] = None,
        element_type: Optional[str] = None
    ) -> bytes:
        """정렬만 변경 (treatAsChar는 유지)"""
        root = ET.parse(BytesIO(xml_content)).getroot()
        modified = False

        # 대상 태그 결정
        target_tags = []
        if element_type == "table":
            target_tags = ['}tbl']
        elif element_type == "image":
            target_tags = ['}pic']
        else:
            target_tags = ['}tbl', '}pic']

        # 부모-자식 맵 생성 (문단 정렬을 위해)
        parent_map = {c: p for p in root.iter() for c in p}

        for elem in root.iter():
            is_target = any(elem.tag.endswith(tag) for tag in target_tags)
            if not is_target:
                continue

            # treatAsChar 상태 확인
            treat_as_char = False
            for child in elem:
                if child.tag.endswith('}pos'):
                    treat_as_char = child.get('treatAsChar', '0') == '1'

                    # 수평 정렬 설정 (horzAlign과 hAlign 모두 설정)
                    if h_align:
                        # horzAlign 속성 변경
                        current_horz_align = child.get('horzAlign', '')
                        if current_horz_align != h_align:
                            child.set('horzAlign', h_align)
                            modified = True

                        # hAlign 속성도 변경 (호환성)
                        current_h_align = child.get('hAlign', '')
                        if current_h_align != h_align:
                            child.set('hAlign', h_align)
                            modified = True

                    # 수직 정렬 설정 (vertAlign과 vAlign 모두 설정)
                    if v_align:
                        # vertAlign 속성 변경
                        current_vert_align = child.get('vertAlign', '')
                        if current_vert_align != v_align:
                            child.set('vertAlign', v_align)
                            modified = True

                        # vAlign 속성도 변경 (호환성)
                        current_v_align = child.get('vAlign', '')
                        if current_v_align != v_align:
                            child.set('vAlign', v_align)
                            modified = True
                    break

            # 글자처럼 취급일 때 문단 정렬도 설정
            if treat_as_char and h_align:
                para_modified = self._set_paragraph_align(elem, parent_map, h_align)
                if para_modified:
                    modified = True

        if modified:
            return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
        return xml_content
