# -*- coding: utf-8 -*-
"""
글머리 기호 양식 변환 모듈 (정규식 기반)

정규식을 사용하여 텍스트를 글머리 기호 양식으로 변환합니다.
SDK 기반 변환이 필요하면 agent.BulletFormatter를 사용하세요.

예: "항목1\\n항목2\\n항목3" → " □ 항목1\\n   ○항목2\\n    - 항목3"
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


# 글머리 기호 스타일 정의
BULLET_STYLES = {
    # 기본 스타일 (네모 → 원 → 대시)
    "default": {
        0: ("□ ", " "),      # 1단계: 빈 네모 + 공백, 들여쓰기 1칸
        1: ("○", "   "),     # 2단계: 빈 원, 들여쓰기 3칸
        2: ("- ", "    "),   # 3단계: 대시 + 공백, 들여쓰기 4칸
    },
    # 채움 스타일 (채운 네모 → 채운 원 → 대시)
    "filled": {
        0: ("■ ", " "),
        1: ("● ", "   "),
        2: ("- ", "    "),
    },
    # 숫자 스타일 (1. → 1) → -)
    "numbered": {
        0: ("1. ", " "),
        1: ("1) ", "   "),
        2: ("- ", "    "),
    },
    # 화살표 스타일
    "arrow": {
        0: ("▶ ", ""),
        1: ("▷ ", "  "),
        2: ("- ", "    "),
    },
}


@dataclass
class FormatResult:
    """형식 변환 결과"""
    success: bool = True
    original_text: str = ""
    formatted_text: str = ""
    changes: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


@dataclass
class BulletItem:
    """글머리 항목"""
    level: int = 0
    text: str = ""
    bullet: str = ""
    indent: str = ""


class BulletFormatter:
    """
    글머리 기호 양식 변환기 (정규식 기반)

    정규식을 사용하여 내용을 글머리 양식으로 변환합니다.
    """

    def __init__(self, style: str = "default"):
        """
        Args:
            style: 글머리 스타일 ("default", "filled", "numbered")
        """
        self.style = style
        self.bullet_config = BULLET_STYLES.get(style, BULLET_STYLES["default"])

    def format_text(
        self,
        text: str,
        levels: Optional[List[int]] = None,
        auto_detect: bool = True
    ) -> FormatResult:
        """
        텍스트를 글머리 기호 양식으로 변환

        Args:
            text: 변환할 텍스트 (줄바꿈으로 구분된 항목들)
            levels: 각 줄의 레벨 지정 [0, 1, 2, ...], None이면 자동 감지
            auto_detect: 자동 레벨 감지 여부

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
        changes = []
        lines = text.strip().split('\n')

        # 레벨 결정
        if levels is None:
            if auto_detect:
                levels = self._detect_levels(lines)
            else:
                levels = [0] * len(lines)

        # 레벨 개수 맞추기
        while len(levels) < len(lines):
            levels.append(levels[-1] if levels else 0)

        # 변환
        formatted_lines = []
        for i, (line, level) in enumerate(zip(lines, levels)):
            line = line.strip()
            if not line:
                formatted_lines.append("")
                continue

            # 기존 글머리 기호 제거
            cleaned_line, had_bullet = self._remove_existing_bullet(line)

            # 새 글머리 기호 적용
            bullet, indent = self._get_bullet_for_level(level)
            formatted_line = f"{indent}{bullet}{cleaned_line}"
            formatted_lines.append(formatted_line)

            if had_bullet:
                changes.append(f"줄 {i+1}: 글머리 기호 변경 (레벨 {level})")
            else:
                changes.append(f"줄 {i+1}: 글머리 기호 추가 (레벨 {level})")

        formatted_text = '\n'.join(formatted_lines)

        return FormatResult(
            success=True,
            original_text=original,
            formatted_text=formatted_text,
            changes=changes
        )

    def _detect_levels(self, lines: List[str]) -> List[int]:
        """줄별 레벨 자동 감지"""
        levels = []

        for line in lines:
            line = line.strip()
            if not line:
                levels.append(0)
                continue

            # 기존 글머리 기호 기반 레벨 감지
            level = self._detect_bullet_level(line)
            if level >= 0:
                levels.append(level)
            else:
                # 들여쓰기 기반
                original = line
                leading_spaces = len(line) - len(line.lstrip())
                if leading_spaces >= 4:
                    levels.append(2)
                elif leading_spaces >= 2:
                    levels.append(1)
                else:
                    levels.append(0)

        return levels

    def _detect_bullet_level(self, line: str) -> int:
        """기존 글머리 기호에서 레벨 감지"""
        line = line.strip()

        # 레벨 0 글머리
        if line.startswith(('□', '■', '◆', '◇')):
            return 0

        # 레벨 1 글머리
        if line.startswith(('○', '●', '◎', '•')):
            return 1

        # 레벨 2 글머리
        if line.startswith(('-', '–', '—', '·', '∙')):
            return 2

        # 숫자 글머리
        if re.match(r'^\d+[.)]', line):
            return 0

        return -1

    def _remove_existing_bullet(self, line: str) -> Tuple[str, bool]:
        """
        기존 글머리 기호 제거 (SDK에서 처리됨)

        Note: 글머리 제거는 SDK(agent.BulletFormatter.analyze_and_strip)에서 처리합니다.
        이 메서드는 SDK가 없을 때 fallback으로만 사용됩니다.
        """
        # SDK에서 처리하므로 여기서는 제거하지 않음
        return line, False

    def _get_bullet_for_level(self, level: int) -> Tuple[str, str]:
        """레벨에 해당하는 글머리 기호와 들여쓰기 반환"""
        level = min(level, 2)
        level = max(level, 0)
        bullet, indent = self.bullet_config.get(level, ("-", "    "))
        return bullet, indent

    def has_bullet(self, text: str) -> bool:
        """텍스트에 글머리 기호가 있는지 확인"""
        text = text.strip()
        if not text:
            return False

        bullet_chars = '□■◆◇○●◎•-–—·∙'
        first_char = text[0] if text else ''
        if first_char in bullet_chars:
            return True

        if re.match(r'^\d+[.)]\s', text):
            return True

        return False

    def parse_items(self, text: str) -> List[BulletItem]:
        """글머리 기호 목록 파싱"""
        items = []
        lines = text.split('\n')

        for line in lines:
            if not line.strip():
                continue
            cleaned, _ = self._remove_existing_bullet(line.strip())
            level = self._detect_bullet_level(line.strip())
            if level < 0:
                level = 0

            items.append(BulletItem(
                level=level,
                text=cleaned,
                bullet="",
                indent=""
            ))

        return items

    def auto_format(self, text: str) -> FormatResult:
        """
        텍스트 구조를 분석하여 자동으로 글머리 기호 스타일 적용

        들여쓰기, 번호, 기존 기호 등을 분석하여 적절한 레벨을 자동 지정합니다.

        Args:
            text: 변환할 텍스트

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

        lines = text.split('\n')
        levels = []

        for line in lines:
            if not line.strip():
                levels.append(0)
                continue

            level = self._auto_detect_level(line)
            levels.append(level)

        return self.format_text(text, levels=levels, auto_detect=False)

    def _auto_detect_level(self, line: str) -> int:
        """
        단일 줄의 레벨을 자동 감지

        우선순위:
        1. 기존 글머리 기호
        2. 들여쓰기
        3. 번호 패턴
        4. 문맥 분석
        """
        original_line = line
        stripped = line.strip()

        if not stripped:
            return 0

        # 1. 기존 글머리 기호 확인
        bullet_level = self._detect_bullet_level(stripped)
        if bullet_level >= 0:
            return bullet_level

        # 2. 들여쓰기 확인
        leading_spaces = len(line) - len(line.lstrip())
        leading_tabs = line.count('\t', 0, len(line) - len(line.lstrip()))

        # 탭은 4칸으로 계산
        effective_indent = leading_spaces + (leading_tabs * 4)

        if effective_indent >= 8:
            return 2
        elif effective_indent >= 4:
            return 1
        elif effective_indent >= 2:
            # 2칸 들여쓰기는 문맥에 따라 레벨 1 또는 0
            return 1

        # 3. 번호/알파벳 패턴 확인
        # "1. ", "1) ", "a. ", "가. " 등
        if re.match(r'^\d+[.)]\s', stripped):
            return 0  # 숫자 번호는 보통 최상위
        if re.match(r'^[a-zA-Z][.)]\s', stripped):
            return 1  # 알파벳은 보통 하위
        if re.match(r'^[가-힣][.)]\s', stripped):
            return 1  # 한글 번호는 보통 하위

        # 4. 하이픈/대시로 시작
        if stripped.startswith(('-', '–', '—')):
            return 2

        return 0

    def normalize_style(self, text: str) -> FormatResult:
        """
        기존 글머리 기호를 현재 스타일로 통일

        다양한 글머리 기호가 섞여 있는 텍스트를 현재 설정된 스타일로 통일합니다.

        Args:
            text: 변환할 텍스트

        Returns:
            FormatResult: 변환 결과
        """
        return self.format_text(text, auto_detect=True)

    def set_style(self, style: str):
        """
        글머리 스타일 변경

        Args:
            style: 글머리 스타일 ("default", "filled", "numbered")
        """
        if style in BULLET_STYLES:
            self.style = style
            self.bullet_config = BULLET_STYLES[style]

    def convert_style(self, text: str, target_style: str) -> FormatResult:
        """
        텍스트의 글머리 기호를 다른 스타일로 변환

        Args:
            text: 변환할 텍스트
            target_style: 대상 스타일 ("default", "filled", "numbered")

        Returns:
            FormatResult: 변환 결과
        """
        original_style = self.style
        self.set_style(target_style)
        result = self.format_text(text, auto_detect=True)
        self.set_style(original_style)
        return result

    def apply_hierarchy(self, items: List[str], base_level: int = 0) -> FormatResult:
        """
        항목 리스트에 계층적 글머리 기호 적용

        Args:
            items: 항목 리스트
            base_level: 시작 레벨 (0, 1, 2)

        Returns:
            FormatResult: 변환 결과
        """
        text = '\n'.join(items)
        levels = [min(base_level + i, 2) for i in range(len(items))]
        return self.format_text(text, levels=levels, auto_detect=False)

    def apply_flat(self, items: List[str], level: int = 0) -> FormatResult:
        """
        항목 리스트에 동일 레벨 글머리 기호 적용

        Args:
            items: 항목 리스트
            level: 적용할 레벨 (0, 1, 2)

        Returns:
            FormatResult: 변환 결과
        """
        text = '\n'.join(items)
        levels = [level] * len(items)
        return self.format_text(text, levels=levels, auto_detect=False)
