# -*- coding: utf-8 -*-
"""
HWPX 스타일 포맷터 모듈

YAML 설정에서 선택한 스타일(paraPrIDRef, charPrIDRef)을
HWPX 문단/글자에 적용합니다.

사용법:
    # YAML 설정 기반
    formatter = StyleFormatter.from_config(config)
    formatter.apply_style(para_elem, level=0)

    # 직접 스타일 지정
    formatter = StyleFormatter()
    formatter.apply_style(para_elem, style_name="outline_1")
"""

import re
import xml.etree.ElementTree as ET
import zipfile
import tempfile
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path

from .base_formatter import BaseFormatter, FormatResult


# 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


@dataclass
class StyleDefinition:
    """스타일 정의"""
    name: str                           # 스타일 이름
    para_pr_id: Optional[str] = None    # paraPrIDRef (문단 스타일)
    char_pr_id: Optional[str] = None    # charPrIDRef (글자 스타일)
    strip_numbers: bool = True          # 기존 번호 제거 여부
    description: str = ""               # 설명


# 기본 스타일 정의 (템플릿 기준)
DEFAULT_STYLES = {
    # 일반 문단
    "body": StyleDefinition(
        name="body",
        para_pr_id="0",
        char_pr_id="0",
        strip_numbers=False,
        description="바탕글"
    ),

    # 개요 스타일 (level 0~6)
    "outline_1": StyleDefinition(
        name="outline_1",
        para_pr_id="2",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 1 (제1급)"
    ),
    "outline_2": StyleDefinition(
        name="outline_2",
        para_pr_id="3",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 2 (제2급)"
    ),
    "outline_3": StyleDefinition(
        name="outline_3",
        para_pr_id="4",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 3 (제3급)"
    ),
    "outline_4": StyleDefinition(
        name="outline_4",
        para_pr_id="5",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 4 (제4급)"
    ),
    "outline_5": StyleDefinition(
        name="outline_5",
        para_pr_id="6",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 5 (제5급)"
    ),
    "outline_6": StyleDefinition(
        name="outline_6",
        para_pr_id="7",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 6 (제6급)"
    ),
    "outline_7": StyleDefinition(
        name="outline_7",
        para_pr_id="8",
        char_pr_id="0",
        strip_numbers=True,
        description="개요 7 (제7급)"
    ),
}

# 개요 번호 패턴 (제거용)
OUTLINE_NUMBER_PATTERNS = [
    r'^\d+\.\s*',                    # 1. 2. 3.
    r'^\d+\.\d+\.\s*',               # 1.1. 1.2.
    r'^\d+\.\d+\.\d+\.\s*',          # 1.1.1.
    r'^[가-힣]\.\s*',                 # 가. 나. 다.
    r'^\([0-9]+\)\s*',               # (1) (2) (3)
    r'^[0-9]+\)\s*',                 # 1) 2) 3)
    r'^[가-힣]\)\s*',                 # 가) 나) 다)
    r'^[IVX]+\.\s*',                 # I. II. III. (로마)
    r'^[ivx]+\.\s*',                 # i. ii. iii.
    r'^[A-Z]\.\s*',                  # A. B. C.
    r'^[a-z]\.\s*',                  # a. b. c.
]


class StyleFormatter(BaseFormatter):
    """
    HWPX 스타일 포맷터

    YAML 설정에서 선택한 스타일을 HWPX 요소에 적용합니다.
    개요 스타일, 일반 문단 스타일, 글자 스타일 등을 지원합니다.
    """

    def __init__(
        self,
        styles: Optional[Dict[str, StyleDefinition]] = None,
        level_mapping: Optional[Dict[int, str]] = None,
        default_style: str = "body"
    ):
        """
        Args:
            styles: 스타일 정의 딕셔너리 {name: StyleDefinition}
            level_mapping: SDK 레벨 → 스타일 이름 매핑 {0: "outline_1", ...}
            default_style: 기본 스타일 이름
        """
        self.styles = styles or DEFAULT_STYLES.copy()
        self.default_style = default_style

        # 레벨 → 스타일 매핑 (기본값: outline_N)
        self.level_mapping = level_mapping or {
            0: "outline_1",
            1: "outline_2",
            2: "outline_3",
            3: "outline_4",
            4: "outline_5",
            5: "outline_6",
            6: "outline_7",
        }

    @classmethod
    def from_config(cls, config: Dict[str, Any]) -> 'StyleFormatter':
        """
        YAML 설정에서 StyleFormatter 생성

        Args:
            config: YAML 설정 딕셔너리

        Returns:
            StyleFormatter 인스턴스
        """
        styles = DEFAULT_STYLES.copy()
        level_mapping = None
        default_style = "body"

        style_config = config.get('style', {})

        # 커스텀 스타일 정의 로드
        if 'definitions' in style_config:
            for name, def_dict in style_config['definitions'].items():
                styles[name] = StyleDefinition(
                    name=name,
                    para_pr_id=str(def_dict.get('paraPrIDRef', '')),
                    char_pr_id=str(def_dict.get('charPrIDRef', '')),
                    strip_numbers=def_dict.get('strip_numbers', True),
                    description=def_dict.get('description', '')
                )

        # 레벨 매핑 로드
        if 'level_mapping' in style_config:
            level_mapping = {
                int(k): v for k, v in style_config['level_mapping'].items()
            }

        # 기본 스타일
        if 'default' in style_config:
            default_style = style_config['default']

        return cls(
            styles=styles,
            level_mapping=level_mapping,
            default_style=default_style
        )

    def get_style(self, name: str) -> Optional[StyleDefinition]:
        """스타일 정의 반환"""
        return self.styles.get(name)

    def get_style_for_level(self, level: int) -> StyleDefinition:
        """레벨에 해당하는 스타일 반환"""
        style_name = self.level_mapping.get(level, self.default_style)
        return self.styles.get(style_name, self.styles[self.default_style])

    def format_with_levels(
        self,
        texts: List[str],
        levels: List[int]
    ) -> FormatResult:
        """
        레벨이 지정된 텍스트 처리

        개요 스타일은 paraPrIDRef로 적용되므로,
        여기서는 기존 번호만 제거하고 적용할 스타일 정보를 반환합니다.

        Args:
            texts: 각 줄의 텍스트 리스트
            levels: 각 줄의 레벨 리스트

        Returns:
            FormatResult: 처리된 텍스트 및 스타일 정보
        """
        if not texts:
            return FormatResult(success=True, formatted_text="", changes=[])

        original_text = '\n'.join(texts)
        formatted_lines = []
        changes = []
        style_info = []  # 각 줄의 스타일 정보

        for i, (text, level) in enumerate(zip(texts, levels)):
            text = text.strip()
            if not text:
                formatted_lines.append("")
                style_info.append(None)
                continue

            style = self.get_style_for_level(level)

            # 기존 번호 제거 (개요 스타일인 경우)
            cleaned_text = text
            if style.strip_numbers:
                cleaned_text = self._strip_outline_numbers(text)
                if cleaned_text != text:
                    changes.append(f"줄 {i+1}: 번호 제거")

            formatted_lines.append(cleaned_text)
            style_info.append({
                'style_name': style.name,
                'para_pr_id': style.para_pr_id,
                'char_pr_id': style.char_pr_id,
            })
            changes.append(
                f"줄 {i+1}: 스타일 '{style.name}' "
                f"(paraPrIDRef={style.para_pr_id}, charPrIDRef={style.char_pr_id})"
            )

        result = FormatResult(
            success=True,
            original_text=original_text,
            formatted_text='\n'.join(formatted_lines),
            changes=changes
        )
        result.style_info = style_info  # 추가 정보
        return result

    def format_text(
        self,
        text: str,
        levels: Optional[List[int]] = None,
        auto_detect: bool = True
    ) -> FormatResult:
        """
        텍스트 처리

        Args:
            text: 변환할 텍스트
            levels: 각 줄의 레벨 지정
            auto_detect: 자동 레벨 감지 여부

        Returns:
            FormatResult: 처리 결과
        """
        if not text or not text.strip():
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=text,
                changes=[]
            )

        lines = text.strip().split('\n')

        if levels is None:
            if auto_detect:
                levels = self._detect_levels(lines)
            else:
                levels = [0] * len(lines)

        # 레벨 개수 맞추기
        while len(levels) < len(lines):
            levels.append(levels[-1] if levels else 0)

        return self.format_with_levels(lines, levels)

    def _detect_levels(self, lines: List[str]) -> List[int]:
        """줄별 레벨 자동 감지"""
        levels = []

        for line in lines:
            line = line.strip()
            if not line:
                levels.append(0)
                continue

            level = self._detect_level_from_number(line)
            levels.append(level)

        return levels

    def _detect_level_from_number(self, line: str) -> int:
        """번호 패턴에서 레벨 감지"""
        # 1.2.3. 형식 (점 개수로 레벨)
        match = re.match(r'^(\d+(?:\.\d+)*)\.\s*', line)
        if match:
            dots = match.group(1).count('.')
            return min(dots, 6)

        # 가. 나. 다. → 레벨 1
        if re.match(r'^[가-힣]\.\s*', line):
            return 1

        # (1) (2) → 레벨 2
        if re.match(r'^\([0-9]+\)\s*', line):
            return 2

        # 가) 나) → 레벨 3
        if re.match(r'^[가-힣]\)\s*', line):
            return 3

        # 1) 2) → 레벨 2
        if re.match(r'^[0-9]+\)\s*', line):
            return 2

        return 0

    def _strip_outline_numbers(self, text: str) -> str:
        """텍스트에서 개요 번호 제거"""
        result = text.strip()
        for pattern in OUTLINE_NUMBER_PATTERNS:
            new_result = re.sub(pattern, '', result)
            if new_result != result:
                return new_result.strip()
        return result

    def has_format(self, text: str) -> bool:
        """텍스트에 개요 번호가 있는지 확인"""
        text = text.strip()
        if not text:
            return False

        for pattern in OUTLINE_NUMBER_PATTERNS:
            if re.match(pattern, text):
                return True
        return False

    def get_style_name(self) -> str:
        """스타일 이름 반환"""
        return "style"

    # ========================================
    # HWPX 요소에 스타일 적용
    # ========================================

    def apply_style_to_paragraph(
        self,
        para_elem: ET.Element,
        style_name: Optional[str] = None,
        level: Optional[int] = None
    ) -> bool:
        """
        문단 요소에 스타일 적용

        Args:
            para_elem: <hp:p> 요소
            style_name: 스타일 이름 (지정 시 우선)
            level: 레벨 (style_name이 없을 때 사용)

        Returns:
            성공 여부
        """
        if style_name:
            style = self.get_style(style_name)
        elif level is not None:
            style = self.get_style_for_level(level)
        else:
            style = self.styles.get(self.default_style)

        if not style:
            return False

        # paraPrIDRef 적용
        if style.para_pr_id:
            para_elem.set('paraPrIDRef', style.para_pr_id)

        return True

    def apply_style_to_run(
        self,
        run_elem: ET.Element,
        style_name: Optional[str] = None,
        level: Optional[int] = None
    ) -> bool:
        """
        run 요소에 글자 스타일 적용

        Args:
            run_elem: <hp:run> 요소
            style_name: 스타일 이름
            level: 레벨

        Returns:
            성공 여부
        """
        if style_name:
            style = self.get_style(style_name)
        elif level is not None:
            style = self.get_style_for_level(level)
        else:
            style = self.styles.get(self.default_style)

        if not style:
            return False

        # charPrIDRef 적용
        if style.char_pr_id:
            run_elem.set('charPrIDRef', style.char_pr_id)

        return True

    def apply_style_to_paragraph_with_runs(
        self,
        para_elem: ET.Element,
        style_name: Optional[str] = None,
        level: Optional[int] = None
    ) -> bool:
        """
        문단과 하위 run 요소 모두에 스타일 적용

        Args:
            para_elem: <hp:p> 요소
            style_name: 스타일 이름
            level: 레벨

        Returns:
            성공 여부
        """
        # 문단 스타일 적용
        self.apply_style_to_paragraph(para_elem, style_name, level)

        # 하위 run 요소에 글자 스타일 적용
        for elem in para_elem.iter():
            if elem.tag.endswith('}run'):
                self.apply_style_to_run(elem, style_name, level)

        return True

    # ========================================
    # 유틸리티
    # ========================================

    def list_styles(self) -> List[str]:
        """사용 가능한 스타일 이름 목록"""
        return list(self.styles.keys())

    def add_style(self, style: StyleDefinition):
        """스타일 추가"""
        self.styles[style.name] = style

    def set_level_mapping(self, level: int, style_name: str):
        """레벨 매핑 설정"""
        if style_name in self.styles:
            self.level_mapping[level] = style_name

    def get_para_pr_id(self, style_name: str) -> Optional[str]:
        """스타일의 paraPrIDRef 반환"""
        style = self.styles.get(style_name)
        return style.para_pr_id if style else None

    def get_char_pr_id(self, style_name: str) -> Optional[str]:
        """스타일의 charPrIDRef 반환"""
        style = self.styles.get(style_name)
        return style.char_pr_id if style else None


def load_style_formatter(config_path: str) -> StyleFormatter:
    """
    YAML 설정 파일에서 StyleFormatter 로드

    Args:
        config_path: YAML 파일 경로

    Returns:
        StyleFormatter 인스턴스
    """
    import yaml

    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    return StyleFormatter.from_config(config)
