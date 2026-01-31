# -*- coding: utf-8 -*-
"""
캡션 양식 변환 모듈 (정규식 기반)

HWPX 파일에서 캡션을 조회하고 정규식을 사용하여 양식을 변환합니다.
SDK 기반 변환이 필요하면 agent.CaptionFormatter를 사용하세요.

사용 예:
    from merge.formatters import CaptionFormatter

    formatter = CaptionFormatter()

    # 캡션 조회
    captions = formatter.get_all_captions("document.hwpx")

    # 캡션 양식 변환 (정규식)
    result = formatter.to_bracket_format("그림 1. 테스트 이미지")
"""

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Callable
from pathlib import Path
from io import BytesIO


# XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

# 캡션 유형별 정규식 패턴
CAPTION_PATTERNS = {
    "figure": [
        r'^(그림)\s*(\d+)\s*[.:]\s*(.+)$',
        r'^(Figure)\s*(\d+)\s*[.:]\s*(.+)$',
        r'^(Fig\.?)\s*(\d+)\s*[.:]\s*(.+)$',
    ],
    "table": [
        r'^(표)\s*(\d+)\s*[.:]\s*(.+)$',
        r'^(Table)\s*(\d+)\s*[.:]\s*(.+)$',
        r'^(Tbl\.?)\s*(\d+)\s*[.:]\s*(.+)$',
    ],
    "equation": [
        r'^(식)\s*[\[(]?(\d+)[\])]?\s*(.*)$',
        r'^(Equation)\s*[\[(]?(\d+)[\])]?\s*(.*)$',
        r'^(Eq\.?)\s*[\[(]?(\d+)[\])]?\s*(.*)$',
    ],
}


@dataclass
class CaptionInfo:
    """캡션 정보"""
    text: str = ""                    # 캡션 텍스트
    caption_type: str = ""            # 캡션 유형 (figure, table, equation 등)
    number: Optional[int] = None      # 캡션 번호
    parent_type: str = ""             # 부모 요소 유형 (tbl, pic 등)
    section_index: int = 0            # 섹션 인덱스
    position: int = 0                 # 문서 내 위치
    new_number: Optional[int] = None  # 재정렬된 새 번호


@dataclass
class FormatResult:
    """양식 변환 결과"""
    success: bool = True
    original_text: str = ""
    formatted_text: str = ""
    changes: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class CaptionFormatter:
    """
    캡션 양식 변환기 (정규식 기반)

    HWPX 파일에서 캡션을 조회하고 정규식을 사용하여 양식을 변환합니다.
    """

    def __init__(self):
        """초기화"""
        pass

    def get_all_captions(self, hwpx_path: str) -> List[CaptionInfo]:
        """
        HWPX 파일에서 모든 캡션 조회

        Args:
            hwpx_path: HWPX 파일 경로

        Returns:
            CaptionInfo 리스트
        """
        captions = []
        hwpx_path = Path(hwpx_path)

        if not hwpx_path.exists():
            return captions

        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
                # section XML 파일들 처리
                section_files = sorted([
                    name for name in zf.namelist()
                    if name.startswith('Contents/section') and name.endswith('.xml')
                ])

                position = 0
                for section_idx, section_name in enumerate(section_files):
                    section_content = zf.read(section_name)
                    section_captions = self._parse_section_captions(
                        section_content, section_idx, position
                    )
                    captions.extend(section_captions)
                    position += len(section_captions)

        except Exception as e:
            print(f"캡션 조회 실패: {e}")

        return captions

    def _parse_section_captions(
        self,
        xml_content: bytes,
        section_idx: int,
        start_position: int
    ) -> List[CaptionInfo]:
        """섹션 XML에서 캡션 파싱"""
        captions = []
        root = ET.parse(BytesIO(xml_content)).getroot()
        position = start_position

        # 모든 캡션 요소 찾기
        for elem in root.iter():
            # caption 태그 찾기
            if elem.tag.endswith('}caption'):
                caption_text, auto_number = self._extract_caption_text(elem)
                if caption_text:
                    # 부모 요소 유형 확인
                    parent_type = self._get_parent_type(elem, root)
                    caption_type = self._detect_caption_type(caption_text, parent_type)
                    # 텍스트에서 번호 추출 시도, 없으면 autoNum 사용
                    number = self._extract_caption_number(caption_text) or auto_number

                    captions.append(CaptionInfo(
                        text=caption_text,
                        caption_type=caption_type,
                        number=number,
                        parent_type=parent_type,
                        section_index=section_idx,
                        position=position
                    ))
                    position += 1

        return captions

    def _extract_caption_text(self, caption_elem) -> Tuple[str, Optional[int]]:
        """
        캡션 요소에서 텍스트와 자동 번호 추출

        Returns:
            (캡션 텍스트, 자동 번호) 튜플
        """
        texts = []
        auto_number = None

        # subList > p > run 구조에서 텍스트 및 autoNum 추출
        for sublist in caption_elem:
            if sublist.tag.endswith('}subList'):
                for p in sublist:
                    if p.tag.endswith('}p'):
                        for run in p:
                            if run.tag.endswith('}run'):
                                for child in run:
                                    # 텍스트 요소
                                    if child.tag.endswith('}t') and child.text:
                                        texts.append(child.text)
                                    # ctrl > autoNum 구조 처리
                                    elif child.tag.endswith('}ctrl'):
                                        for ctrl_child in child:
                                            if ctrl_child.tag.endswith('}autoNum'):
                                                num_str = ctrl_child.get('num')
                                                if num_str:
                                                    try:
                                                        auto_number = int(num_str)
                                                        texts.append(num_str)  # 번호를 텍스트에 포함
                                                    except ValueError:
                                                        pass

        # 직접 텍스트 확인 (fallback)
        if not texts:
            for elem in caption_elem.iter():
                if elem.tag.endswith('}t') and elem.text:
                    texts.append(elem.text)
                elif elem.tag.endswith('}autoNum'):
                    num_str = elem.get('num')
                    if num_str:
                        try:
                            auto_number = int(num_str)
                            texts.append(num_str)
                        except ValueError:
                            pass

        return ''.join(texts).strip(), auto_number

    def _get_parent_type(self, caption_elem, root) -> str:
        """캡션의 부모 요소 유형 확인"""
        # 부모를 찾기 위해 전체 트리 순회
        def find_parent(elem, target, parent=None):
            if elem is target:
                return parent
            for child in elem:
                result = find_parent(child, target, elem)
                if result is not None:
                    return result
            return None

        parent = find_parent(root, caption_elem)
        if parent is not None:
            tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
            return tag

        return ""

    def _detect_caption_type(self, text: str, parent_type: str) -> str:
        """캡션 유형 감지"""
        text_lower = text.lower()

        # 텍스트 기반 감지
        if any(kw in text_lower for kw in ['그림', 'figure', 'fig.', 'fig ']):
            return "figure"
        if any(kw in text_lower for kw in ['표', 'table', 'tbl.', 'tbl ']):
            return "table"
        if any(kw in text_lower for kw in ['식', 'equation', 'eq.', 'eq ']):
            return "equation"
        if any(kw in text_lower for kw in ['사진', 'photo', 'image']):
            return "figure"

        # 부모 요소 기반 감지
        if parent_type == 'tbl':
            return "table"
        if parent_type == 'pic':
            return "figure"

        return "default"

    def _extract_caption_number(self, text: str) -> Optional[int]:
        """캡션에서 번호 추출"""
        # 다양한 번호 패턴 매칭
        patterns = [
            r'(?:그림|표|식|Figure|Table|Equation|Fig\.|Tbl\.)\s*(\d+)',
            r'(\d+)\s*[.:]',
            r'\[(\d+)\]',
            r'\((\d+)\)',
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    return int(match.group(1))
                except ValueError:
                    pass

        return None

    def extract_title(self, text: str) -> str:
        """
        캡션에서 제목(설명)만 추출 (번호 제거)

        Args:
            text: 캡션 텍스트

        Returns:
            제목만 추출된 텍스트
        """
        # "그림 1. 제목" → "제목"
        # "표 2. 제목" → "제목"
        # "Figure 3. Title" → "Title"
        # "Figure3-Title" → "Title"
        patterns = [
            # 구분자가 있는 경우: "그림 1. 제목", "Figure 3: Title"
            r'^(?:그림|표|식|Figure|Table|Equation|Fig\.?|Tbl\.?|Eq\.?)\s*\d+\s*[.:\-–—]\s*(.+)$',
            # 공백으로 구분: "그림 1 제목"
            r'^(?:그림|표|식|Figure|Table|Equation|Fig\.?|Tbl\.?|Eq\.?)\s*\d+\s+(.+)$',
            # 번호가 바로 붙은 경우: "Figure3Title" → "Title" (숫자 뒤 비숫자)
            r'^(?:그림|표|식|Figure|Table|Equation|Fig\.?|Tbl\.?|Eq\.?)\s*\d+([^\d\s].*)$',
            # 번호 없이 유형만: "그림 제목"
            r'^(?:그림|표|식|Figure|Table|Equation|Fig\.?|Tbl\.?|Eq\.?)\s+(.+)$',
        ]

        for pattern in patterns:
            match = re.match(pattern, text.strip(), re.IGNORECASE)
            if match:
                title = match.group(1).strip()
                # 앞의 구분자 제거
                title = re.sub(r'^[\-–—:.\s]+', '', title)
                return title

        return text.strip()

    def get_type_prefix(self, text: str, caption_type: str = "") -> str:
        """캡션에서 유형 접두어 추출"""
        text_lower = text.lower()

        if '그림' in text or caption_type == 'figure':
            return '그림'
        if '표' in text or caption_type == 'table':
            return '표'
        if '식' in text or caption_type == 'equation':
            return '식'
        if 'figure' in text_lower:
            return 'Figure'
        if 'table' in text_lower:
            return 'Table'
        if 'equation' in text_lower:
            return 'Equation'

        return ''

    def to_bracket_format(
        self,
        text: str,
        caption_type: str = "default"
    ) -> FormatResult:
        """
        캡션을 대괄호 형식으로 변환 (번호 및 유형 접두어 제거)
        예: "표 1. 연구 결과" → "[연구 결과]"

        대괄호 안에는 제목만 들어감 (유형 접두어 제외)

        Args:
            text: 캡션 텍스트
            caption_type: 캡션 유형

        Returns:
            FormatResult
        """
        title = self.extract_title(text)

        bracket_text = f"[{title}]" if title else text

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=bracket_text,
            changes=["대괄호 형식으로 변환"]
        )

    def replace_number(self, text: str, new_number: int, caption_type: str) -> str:
        """캡션 텍스트에서 번호만 교체"""
        # 유형별 패턴
        patterns = {
            "figure": [
                (r'(그림\s*)\d+', rf'\g<1>{new_number}'),
                (r'(Figure\s*)\d+', rf'\g<1>{new_number}'),
                (r'(Fig\.?\s*)\d+', rf'\g<1>{new_number}'),
            ],
            "table": [
                (r'(표\s*)\d+', rf'\g<1>{new_number}'),
                (r'(Table\s*)\d+', rf'\g<1>{new_number}'),
                (r'(Tbl\.?\s*)\d+', rf'\g<1>{new_number}'),
            ],
            "equation": [
                (r'(식\s*[\[(]?)\d+', rf'\g<1>{new_number}'),
                (r'(Equation\s*[\[(]?)\d+', rf'\g<1>{new_number}'),
                (r'(Eq\.?\s*[\[(]?)\d+', rf'\g<1>{new_number}'),
            ],
        }

        # 유형에 맞는 패턴 적용
        type_patterns = patterns.get(caption_type, [])
        for pattern, replacement in type_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return re.sub(pattern, replacement, text, count=1, flags=re.IGNORECASE)

        # 기본 패턴 (숫자만 교체)
        default_patterns = [
            (r'^(\D*)\d+', rf'\g<1>{new_number}'),  # 텍스트 시작 후 첫 숫자
        ]
        for pattern, replacement in default_patterns:
            if re.search(pattern, text):
                return re.sub(pattern, replacement, text, count=1)

        return text

    def renumber_captions(
        self,
        captions: List[CaptionInfo],
        by_type: bool = True,
        start_number: int = 1
    ) -> List[CaptionInfo]:
        """
        캡션 번호 재정렬

        Args:
            captions: CaptionInfo 리스트
            by_type: True면 유형별로 따로 번호 매김, False면 전체 순서로 번호 매김
            start_number: 시작 번호 (기본값: 1)

        Returns:
            번호가 재정렬된 CaptionInfo 리스트
        """
        if not captions:
            return captions

        # 위치 순으로 정렬
        sorted_captions = sorted(captions, key=lambda c: (c.section_index, c.position))

        if by_type:
            # 유형별로 따로 번호 매김
            type_counters: Dict[str, int] = {}
            for caption in sorted_captions:
                cap_type = caption.caption_type or "default"
                if cap_type not in type_counters:
                    type_counters[cap_type] = start_number
                caption.new_number = type_counters[cap_type]
                type_counters[cap_type] += 1
        else:
            # 전체 순서로 번호 매김
            for i, caption in enumerate(sorted_captions, start_number):
                caption.new_number = i

        return sorted_captions

    def format_all_to_bracket(
        self,
        captions: List[CaptionInfo]
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 대괄호 형식으로 일괄 변환

        Args:
            captions: CaptionInfo 리스트

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []
        for caption in captions:
            result = self.to_bracket_format(caption.text, caption.caption_type)
            results.append((caption, result))
        return results

    def apply_to_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        transform_func: Optional[Callable[[str, str], str]] = None,
        to_bracket: bool = False,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 변환하여 저장

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            transform_func: 커스텀 변환 함수 (text, caption_type) -> new_text
            to_bracket: True면 대괄호 형식으로 변환 "[표 제목]"
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬 (1부터 시작)

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

        # 임시 파일로 작업
        temp_path = hwpx_path.with_suffix('.hwpx.tmp')

        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf_in:
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for item in zf_in.namelist():
                        content = zf_in.read(item)

                        # section XML 파일 처리
                        if item.startswith('Contents/section') and item.endswith('.xml'):
                            content = self._transform_section_captions(
                                content, transform_func, to_bracket,
                                keep_auto_num, renumber
                            )

                        zf_out.writestr(item, content)

            # 결과 파일로 이동
            if output_path != hwpx_path:
                shutil.copy2(temp_path, output_path)
            else:
                shutil.move(temp_path, output_path)

            return str(output_path)

        finally:
            if temp_path.exists():
                temp_path.unlink()

    def _transform_section_captions(
        self,
        xml_content: bytes,
        transform_func: Optional[Callable[[str, str], str]],
        to_bracket: bool,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> bytes:
        """섹션 XML의 캡션 변환"""
        root = ET.parse(BytesIO(xml_content)).getroot()
        modified = False

        # 번호 재정렬을 위한 카운터 (유형별)
        type_counters: Dict[str, int] = {}

        for caption_elem in root.iter():
            if not caption_elem.tag.endswith('}caption'):
                continue

            # 캡션 텍스트 추출
            caption_text, auto_number = self._extract_caption_text(caption_elem)
            if not caption_text:
                continue

            # 부모 유형으로 캡션 타입 결정
            parent_type = self._get_parent_type(caption_elem, root)
            caption_type = self._detect_caption_type(caption_text, parent_type)

            # 번호 재정렬
            new_number = None
            if renumber:
                if caption_type not in type_counters:
                    type_counters[caption_type] = 1
                new_number = type_counters[caption_type]
                type_counters[caption_type] += 1

            # 변환 수행
            if transform_func:
                new_text = transform_func(caption_text, caption_type)
            elif to_bracket:
                result = self.to_bracket_format(caption_text, caption_type)
                new_text = result.formatted_text
            else:
                # 번호만 재정렬하는 경우
                if renumber and new_number:
                    self._update_auto_num(caption_elem, new_number, caption_type)
                    modified = True
                continue

            if new_text and new_text != caption_text:
                # 캡션 텍스트 업데이트
                self._update_caption_text(
                    caption_elem, new_text,
                    keep_auto_num=keep_auto_num,
                    new_number=new_number if renumber else None,
                    caption_type=caption_type
                )
                modified = True

        if modified:
            return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
        return xml_content

    def _update_caption_text(
        self,
        caption_elem,
        new_text: str,
        keep_auto_num: bool = False,
        new_number: Optional[int] = None,
        caption_type: str = "default"
    ):
        """캡션 요소의 텍스트 업데이트"""
        # subList > p > run 구조에서 텍스트 요소 찾아서 업데이트
        for sublist in caption_elem:
            if sublist.tag.endswith('}subList'):
                for p in sublist:
                    if p.tag.endswith('}p'):
                        for run in p:
                            if run.tag.endswith('}run'):
                                if keep_auto_num:
                                    # autoNum 유지하면서 텍스트만 변경
                                    self._update_text_keep_autonum(run, new_text, new_number, caption_type)
                                else:
                                    # autoNum 제거하고 텍스트만 설정
                                    self._update_text_remove_autonum(run, new_text)
                                return

    def _update_text_remove_autonum(self, run_elem, new_text: str):
        """autoNum을 제거하고 텍스트만 설정"""
        # ctrl(autoNum) 요소 제거
        to_remove = []
        for child in run_elem:
            if child.tag.endswith('}ctrl'):
                to_remove.append(child)
        for elem in to_remove:
            run_elem.remove(elem)

        # 첫 번째 t 요소에 새 텍스트 설정
        first_t = None
        other_ts = []
        for child in run_elem:
            if child.tag.endswith('}t'):
                if first_t is None:
                    first_t = child
                else:
                    other_ts.append(child)

        if first_t is not None:
            first_t.text = new_text
            # 나머지 t 요소 제거
            for t in other_ts:
                run_elem.remove(t)

    def _update_text_keep_autonum(
        self,
        run_elem,
        new_text: str,
        new_number: Optional[int] = None,
        caption_type: str = "default"
    ):
        """autoNum을 유지하면서 텍스트 변경"""
        # 유형 접두어와 제목 분리
        prefix, title = self._parse_bracket_text(new_text, caption_type)

        # autoNum 번호 업데이트
        if new_number is not None:
            for child in run_elem:
                if child.tag.endswith('}ctrl'):
                    for auto_num in child:
                        if auto_num.tag.endswith('}autoNum'):
                            auto_num.set('num', str(new_number))

        # 텍스트 요소들 수집
        t_elements = []
        ctrl_elem = None
        for child in list(run_elem):
            if child.tag.endswith('}t'):
                t_elements.append(child)
            elif child.tag.endswith('}ctrl'):
                ctrl_elem = child

        # 첫 번째 t에 접두어, 나머지 제거 후 제목용 t 추가
        if t_elements:
            # 첫 번째 t에 접두어 설정
            t_elements[0].text = prefix + " " if prefix else ""

            # 나머지 t 요소 제거
            for t in t_elements[1:]:
                run_elem.remove(t)

            # ctrl 뒤에 제목용 t 추가
            if ctrl_elem is not None:
                # ctrl 다음 위치에 새 t 삽입
                ns = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
                new_t = ET.Element(f'{ns}t')
                new_t.text = " " + title + " "

                # ctrl 다음 위치 찾기
                idx = list(run_elem).index(ctrl_elem)
                run_elem.insert(idx + 1, new_t)

    def _parse_bracket_text(self, text: str, caption_type: str) -> Tuple[str, str]:
        """대괄호 형식 텍스트에서 접두어와 제목 분리"""
        # "[표 제목]" → ("표", "제목")
        # "[그림 테스트]" → ("그림", "테스트")
        match = re.match(r'^\[(그림|표|식|Figure|Table|Equation)\s+(.+)\]$', text, re.IGNORECASE)
        if match:
            return match.group(1), match.group(2)

        # 대괄호만 제거
        if text.startswith('[') and text.endswith(']'):
            inner = text[1:-1]
            # 유형 접두어 추출 시도
            prefix = self.get_type_prefix(inner, caption_type)
            if prefix and inner.startswith(prefix):
                title = inner[len(prefix):].strip()
                return prefix, title
            return "", inner

        return "", text

    def _update_auto_num(self, caption_elem, new_number: int, caption_type: str):
        """캡션의 autoNum 번호만 업데이트"""
        for elem in caption_elem.iter():
            if elem.tag.endswith('}autoNum'):
                elem.set('num', str(new_number))
                # numType도 업데이트
                num_type_map = {
                    'figure': 'FIGURE',
                    'table': 'TABLE',
                    'equation': 'EQUATION',
                }
                if caption_type in num_type_map:
                    elem.set('numType', num_type_map[caption_type])
                return

    def apply_bracket_format(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 대괄호 형식으로 변환하여 저장
        예: "표 1. 연구 결과" → "[표 연구 결과]"

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬 (1부터 시작)

        Returns:
            저장된 파일 경로
        """
        return self.apply_to_hwpx(
            hwpx_path, output_path,
            to_bracket=True,
            keep_auto_num=keep_auto_num, renumber=renumber
        )

    def renumber_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        HWPX 파일의 캡션 번호만 재정렬 (텍스트 변경 없이)

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)

        Returns:
            저장된 파일 경로
        """
        return self.apply_to_hwpx(hwpx_path, output_path, renumber=True)

    # ========== 캡션 위치 변경 기능 ==========

    def set_caption_position(
        self,
        hwpx_path: str,
        position: str = "TOP",
        output_path: Optional[str] = None,
        caption_type: Optional[str] = None
    ) -> str:
        """
        HWPX 파일의 캡션 위치 변경

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            position: 캡션 위치 ("TOP", "BOTTOM", "LEFT", "RIGHT")
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            caption_type: 특정 유형만 변경 ("table", "figure", "equation", None이면 전체)

        Returns:
            저장된 파일 경로
        """
        valid_positions = {"TOP", "BOTTOM", "LEFT", "RIGHT"}
        position = position.upper()
        if position not in valid_positions:
            raise ValueError(f"유효하지 않은 위치: {position}. 가능한 값: {valid_positions}")

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
                            content = self._change_caption_position(
                                content, position, caption_type
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

    def _change_caption_position(
        self,
        xml_content: bytes,
        position: str,
        caption_type: Optional[str] = None
    ) -> bytes:
        """섹션 XML의 캡션 위치 변경"""
        root = ET.parse(BytesIO(xml_content)).getroot()
        modified = False

        for caption_elem in root.iter():
            if not caption_elem.tag.endswith('}caption'):
                continue

            # 캡션 유형 필터링
            if caption_type:
                caption_text, _ = self._extract_caption_text(caption_elem)
                parent_type = self._get_parent_type(caption_elem, root)
                detected_type = self._detect_caption_type(caption_text, parent_type)
                if detected_type != caption_type:
                    continue

            # side 속성 변경
            current_side = caption_elem.get('side', '')
            if current_side != position:
                caption_elem.set('side', position)
                modified = True

        if modified:
            return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
        return xml_content

    def set_caption_to_top(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        caption_type: Optional[str] = None
    ) -> str:
        """
        캡션을 표/그림 위로 이동

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            caption_type: 특정 유형만 변경 (None이면 전체)

        Returns:
            저장된 파일 경로
        """
        return self.set_caption_position(hwpx_path, "TOP", output_path, caption_type)

    def set_caption_to_bottom(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        caption_type: Optional[str] = None
    ) -> str:
        """
        캡션을 표/그림 아래로 이동

        Args:
            hwpx_path: HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            caption_type: 특정 유형만 변경 (None이면 전체)

        Returns:
            저장된 파일 경로
        """
        return self.set_caption_position(hwpx_path, "BOTTOM", output_path, caption_type)

    # ========== 글자처럼 취급 설정 ==========

    def set_treat_as_char(
        self,
        hwpx_path: str,
        treat_as_char: bool = True,
        output_path: Optional[str] = None,
        element_type: Optional[str] = None
    ) -> str:
        """
        HWPX 파일의 테이블/이미지를 '글자처럼 취급' 설정

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            treat_as_char: True면 글자처럼 취급, False면 어울림(앵커)
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            element_type: 대상 요소 ("table", "image", None이면 전체)

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
                            content = self._change_treat_as_char(
                                content, treat_as_char, element_type
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

    def _change_treat_as_char(
        self,
        xml_content: bytes,
        treat_as_char: bool,
        element_type: Optional[str] = None
    ) -> bytes:
        """섹션 XML의 treatAsChar 속성 변경"""
        root = ET.parse(BytesIO(xml_content)).getroot()
        modified = False
        value = "1" if treat_as_char else "0"

        # 대상 태그 결정
        target_tags = []
        if element_type == "table":
            target_tags = ['}tbl']
        elif element_type == "image":
            target_tags = ['}pic']
        else:
            target_tags = ['}tbl', '}pic']

        for elem in root.iter():
            # 테이블 또는 이미지 요소 확인
            is_target = any(elem.tag.endswith(tag) for tag in target_tags)
            if not is_target:
                continue

            # hp:pos 요소 찾기
            for child in elem:
                if child.tag.endswith('}pos'):
                    current_value = child.get('treatAsChar', '0')
                    if current_value != value:
                        child.set('treatAsChar', value)
                        modified = True
                    break

        if modified:
            return ET.tostring(root, encoding='UTF-8', xml_declaration=True)
        return xml_content

    def set_table_as_char(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """테이블을 글자처럼 취급으로 설정"""
        return self.set_treat_as_char(hwpx_path, True, output_path, "table")

    def set_table_as_anchor(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """테이블을 어울림(앵커)으로 설정"""
        return self.set_treat_as_char(hwpx_path, False, output_path, "table")

    def set_image_as_char(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """이미지를 글자처럼 취급으로 설정"""
        return self.set_treat_as_char(hwpx_path, True, output_path, "image")

    def set_image_as_anchor(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """이미지를 어울림(앵커)으로 설정"""
        return self.set_treat_as_char(hwpx_path, False, output_path, "image")

    def set_all_as_char(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """모든 테이블/이미지를 글자처럼 취급으로 설정"""
        return self.set_treat_as_char(hwpx_path, True, output_path, None)

    def set_all_as_anchor(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """모든 테이블/이미지를 어울림(앵커)으로 설정"""
        return self.set_treat_as_char(hwpx_path, False, output_path, None)

    # ========== 스타일 자동 적용 기능 ==========

    def auto_format(self, text: str) -> FormatResult:
        """
        캡션 텍스트를 자동으로 표준 양식으로 변환

        텍스트를 분석하여 유형을 감지하고 표준 양식으로 변환합니다.
        예: "그림1 테스트" → "그림 1. 테스트"

        Args:
            text: 변환할 캡션 텍스트

        Returns:
            FormatResult: 변환 결과
        """
        if not text or not text.strip():
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=text,
                changes=[]
            )

        # 유형 감지
        caption_type = self._detect_caption_type(text, "")

        # 번호 추출
        number = self._extract_caption_number(text)

        # 제목 추출
        title = self.extract_title(text)

        # 표준 양식으로 변환
        formatted = self._build_standard_caption(caption_type, number, title)

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=[f"표준 양식으로 변환 (유형: {caption_type})"]
        )

    def _build_standard_caption(
        self,
        caption_type: str,
        number: Optional[int],
        title: str
    ) -> str:
        """표준 캡션 형식 생성"""
        type_prefixes = {
            "figure": "그림",
            "table": "표",
            "equation": "식",
        }

        prefix = type_prefixes.get(caption_type, "")

        if caption_type == "equation":
            # 수식: "식 (1)" 형식
            if number:
                return f"{prefix} ({number})" + (f" {title}" if title else "")
            return f"{prefix}" + (f" {title}" if title else "")
        else:
            # 그림/표: "그림 1. 제목" 형식
            if number and title:
                return f"{prefix} {number}. {title}"
            elif number:
                return f"{prefix} {number}."
            elif title:
                return f"{prefix} {title}"
            return prefix

    def normalize_format(self, text: str) -> FormatResult:
        """
        캡션 형식을 정규화 (공백, 구두점 통일)

        Args:
            text: 변환할 캡션 텍스트

        Returns:
            FormatResult: 변환 결과
        """
        if not text or not text.strip():
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=text,
                changes=[]
            )

        original = text
        normalized = text.strip()

        # 유형 키워드 뒤 공백 정규화
        # "그림1" → "그림 1", "표  2" → "표 2"
        normalized = re.sub(r'(그림|표|식|Figure|Table|Equation)\s*(\d+)', r'\1 \2', normalized)

        # 번호 뒤 구두점 정규화
        # "그림 1:" → "그림 1.", "그림 1 -" → "그림 1."
        normalized = re.sub(r'(\d+)\s*[:\-–—]\s*', r'\1. ', normalized)

        # 다중 공백 제거
        normalized = re.sub(r'\s+', ' ', normalized)

        changes = []
        if normalized != original:
            changes.append("형식 정규화 완료")

        return FormatResult(
            success=True,
            original_text=original,
            formatted_text=normalized,
            changes=changes
        )

    def to_standard_format(
        self,
        text: str,
        caption_type: str = "",
        separator: str = ". "
    ) -> FormatResult:
        """
        캡션을 표준 형식으로 변환
        예: "그림 1. 제목" 형식

        Args:
            text: 변환할 캡션 텍스트
            caption_type: 캡션 유형 (자동 감지하려면 빈 문자열)
            separator: 번호와 제목 사이 구분자 (기본: ". ")

        Returns:
            FormatResult: 변환 결과
        """
        if not text or not text.strip():
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=text,
                changes=[]
            )

        # 유형 감지
        if not caption_type:
            caption_type = self._detect_caption_type(text, "")

        # 번호 추출
        number = self._extract_caption_number(text)

        # 제목 추출
        title = self.extract_title(text)

        # 유형 접두어
        prefix = self.get_type_prefix(text, caption_type)

        # 표준 형식 생성
        if number and title:
            formatted = f"{prefix} {number}{separator}{title}"
        elif number:
            formatted = f"{prefix} {number}"
        elif title:
            formatted = f"{prefix}{separator}{title}" if prefix else title
        else:
            formatted = prefix

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=["표준 형식으로 변환"]
        )

    def to_parenthesis_format(self, text: str) -> FormatResult:
        """
        캡션을 괄호 형식으로 변환
        예: "그림 1. 제목" → "그림 (1) 제목"

        Args:
            text: 변환할 캡션 텍스트

        Returns:
            FormatResult: 변환 결과
        """
        caption_type = self._detect_caption_type(text, "")
        number = self._extract_caption_number(text)
        title = self.extract_title(text)
        prefix = self.get_type_prefix(text, caption_type)

        if number:
            if title:
                formatted = f"{prefix} ({number}) {title}"
            else:
                formatted = f"{prefix} ({number})"
        else:
            formatted = f"{prefix} {title}" if title else prefix

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=["괄호 형식으로 변환"]
        )

    def remove_number(self, text: str) -> FormatResult:
        """
        캡션에서 번호 제거
        예: "그림 1. 제목" → "그림 제목"

        Args:
            text: 변환할 캡션 텍스트

        Returns:
            FormatResult: 변환 결과
        """
        caption_type = self._detect_caption_type(text, "")
        title = self.extract_title(text)
        prefix = self.get_type_prefix(text, caption_type)

        if prefix and title:
            formatted = f"{prefix} {title}"
        elif prefix:
            formatted = prefix
        else:
            formatted = title

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=["번호 제거"]
        )

    def format_all_auto(
        self,
        captions: List[CaptionInfo]
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 자동 포맷 일괄 적용

        Args:
            captions: CaptionInfo 리스트

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []
        for caption in captions:
            result = self.auto_format(caption.text)
            results.append((caption, result))
        return results

    def apply_auto_format_to_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        renumber: bool = True
    ) -> str:
        """
        HWPX 파일의 캡션을 자동으로 표준 양식으로 변환

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            renumber: True면 자동 번호 재정렬

        Returns:
            저장된 파일 경로
        """
        def auto_transform(text: str, caption_type: str) -> str:
            result = self.auto_format(text)
            return result.formatted_text

        return self.apply_to_hwpx(
            hwpx_path, output_path,
            transform_func=auto_transform,
            renumber=renumber
        )


def get_captions(hwpx_path: str) -> List[CaptionInfo]:
    """
    HWPX 파일에서 캡션 조회 (편의 함수)

    Args:
        hwpx_path: HWPX 파일 경로

    Returns:
        CaptionInfo 리스트
    """
    formatter = CaptionFormatter()
    return formatter.get_all_captions(hwpx_path)


def print_captions(captions: List[CaptionInfo]):
    """캡션 목록 출력"""
    if not captions:
        print("캡션이 없습니다.")
        return

    print(f"총 {len(captions)}개의 캡션:")
    print("-" * 60)
    for i, caption in enumerate(captions, 1):
        print(f"{i}. [{caption.caption_type}] {caption.text}")
        if caption.number:
            number_info = f"번호: {caption.number}"
            if caption.new_number and caption.new_number != caption.number:
                number_info += f" → {caption.new_number}"
            print(f"   {number_info}")
        print(f"   부모: {caption.parent_type}, 섹션: {caption.section_index}")
        print()


def renumber_captions(
    hwpx_path: str,
    by_type: bool = True,
    start_number: int = 1
) -> List[CaptionInfo]:
    """
    HWPX 파일의 캡션 번호 재정렬 (편의 함수)

    Args:
        hwpx_path: HWPX 파일 경로
        by_type: True면 유형별로 따로 번호 매김 (그림 1, 2... 표 1, 2...)
        start_number: 시작 번호 (기본값: 1)

    Returns:
        번호가 재정렬된 CaptionInfo 리스트
    """
    formatter = CaptionFormatter()
    captions = formatter.get_all_captions(hwpx_path)
    return formatter.renumber_captions(captions, by_type, start_number)
