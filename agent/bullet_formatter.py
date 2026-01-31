# -*- coding: utf-8 -*-
"""
글머리 기호 양식 변환 모듈 (Claude SDK 기반)

Claude Code SDK를 사용하여 텍스트를 글머리 기호 양식으로 변환합니다.
정규식만 사용하려면 merge.formatters.BulletFormatter를 사용하세요.

예: "항목1\\n항목2\\n항목3" → " □ 항목1\\n   ○항목2\\n    - 항목3"
"""

from dataclasses import dataclass, field
from typing import List, Optional

from .sdk import ClaudeSDK

# 정규식 기반 포맷터 import (fallback용)
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from merge.formatters.bullet_formatter import (
    BulletFormatter as RegexBulletFormatter,
    FormatResult,
    BulletItem,
    BULLET_STYLES,
)


# 기본 프롬프트 - 글머리 양식 변환
DEFAULT_INSTRUCTION = """다음 내용을 개요 3단계 글머리 기호 양식으로 변환해주세요.

글머리 기호 규칙:
- 1단계: " □ " (공백 + 빈 네모 + 공백)
- 2단계: "   ○" (공백 3개 + 빈 원)
- 3단계: "    - " (공백 4개 + 대시 + 공백)

중요 규칙:
1. 내용의 계층 구조를 파악하여 적절한 레벨을 지정해주세요.
2. 변환된 텍스트만 반환하세요.
3. 설명, 코멘트, 부가 설명을 절대 추가하지 마세요.
4. 코드 블록(```)이나 백틱(`)을 사용하지 마세요.
5. 원본 텍스트의 내용만 글머리 기호를 붙여 반환하세요."""

# 레벨 분석 프롬프트
LEVEL_ANALYSIS_INSTRUCTION = """다음 텍스트의 각 줄에 대해 계층 레벨(0, 1, 2)을 분석해주세요.

레벨 기준:
- 0: 최상위 항목, 주요 내용, 독립적인 항목
- 1: 하위 항목, 세부 설명, 부연 설명
- 2: 가장 하위 항목, 세부 사항, 예시

분석 기준:
1. 문맥과 의미를 파악하여 계층 구조를 결정
2. 들여쓰기, 번호, 기존 글머리 기호도 참고
3. 상위 항목을 설명하거나 보충하는 내용은 하위 레벨
4. 새로운 주제나 독립적인 내용은 상위 레벨

출력 형식:
각 줄마다 레벨 번호만 출력 (0, 1, 2 중 하나)
예시:
0
1
1
2
0

텍스트:
{text}

레벨 (숫자만):"""


class BulletFormatter:
    """
    글머리 기호 양식 변환기 (Claude SDK 기반)

    Claude Code SDK를 사용하여 내용을 글머리 양식으로 변환합니다.
    SDK 호출 실패 시 정규식 기반 변환으로 fallback합니다.
    """

    def __init__(self, style: str = "default", instruction: Optional[str] = None):
        """
        Args:
            style: 글머리 스타일 ("default", "filled", "numbered")
            instruction: 커스텀 프롬프트 (None이면 기본 프롬프트 사용)
        """
        self.style = style
        self.instruction = instruction or DEFAULT_INSTRUCTION
        self.sdk = ClaudeSDK(timeout=30)
        # 정규식 기반 포맷터 (fallback용)
        self._regex_formatter = RegexBulletFormatter(style=style)

    def format_text(self, text: str) -> FormatResult:
        """
        텍스트를 글머리 기호 양식으로 변환 (SDK 사용)

        Args:
            text: 변환할 텍스트 (줄바꿈으로 구분된 항목들)

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

        prompt = f"""{self.instruction}

변환할 내용:
{text}
"""

        result = self.sdk.call(prompt)

        if result.success and result.output:
            cleaned = self._clean_response(result.output)
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=cleaned,
                changes=["Claude SDK를 통한 양식 변환 완료"]
            )
        else:
            # SDK 실패 시 정규식 기반 변환 사용
            return self.format_basic(text)

    def format_basic(
        self,
        text: str,
        levels: Optional[List[int]] = None,
        auto_detect: bool = True
    ) -> FormatResult:
        """
        기본 변환 (SDK 없이 정규식 사용)

        Args:
            text: 변환할 텍스트
            levels: 각 줄의 레벨 지정 [0, 1, 2, ...], None이면 자동 감지
            auto_detect: 자동 레벨 감지 여부

        Returns:
            FormatResult: 변환 결과
        """
        return self._regex_formatter.format_text(text, levels, auto_detect)

    def _clean_response(self, text: str) -> str:
        """SDK 응답에서 마크다운 코드 블록 및 불필요한 내용 제거"""
        text = text.strip()

        # ```로 시작하고 끝나는 경우 (코드 블록)
        if text.startswith('```') and text.endswith('```'):
            lines = text.split('\n')
            if len(lines) >= 2:
                lines = lines[1:-1]
            text = '\n'.join(lines)

        # 인라인 코드 블록 제거
        while '```' in text:
            text = text.replace('```', '')

        # 단일 백틱 제거
        while '`' in text:
            text = text.replace('`', '')

        # 글머리 기호로 시작하는 줄만 추출 (추가 설명 제거)
        lines = text.split('\n')
        bullet_chars = '□■◆◇○●◎•-–—·∙ '
        result_lines = []

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            first_char = stripped[0] if stripped else ''
            if first_char in bullet_chars or line.startswith(' '):
                result_lines.append(line)
            else:
                # 글머리 기호가 아닌 줄이 나오면 중단
                break

        if result_lines:
            text = '\n'.join(result_lines)

        return text.strip()

    def analyze_levels(self, text: str) -> List[int]:
        """
        SDK를 사용하여 텍스트의 각 줄에 대한 계층 레벨 분석

        Args:
            text: 분석할 텍스트

        Returns:
            각 줄의 레벨 리스트 [0, 1, 2, ...]
        """
        if not text or not text.strip():
            return []

        lines = [l for l in text.split('\n') if l.strip()]
        if not lines:
            return []

        prompt = LEVEL_ANALYSIS_INSTRUCTION.format(text=text)
        result = self.sdk.call(prompt)

        if result.success and result.output:
            levels = self._parse_level_response(result.output, len(lines))
            return levels
        else:
            # SDK 실패 시 정규식 기반 분석
            return self._regex_formatter._detect_levels(lines)

    def _parse_level_response(self, response: str, expected_count: int) -> List[int]:
        """SDK 응답에서 레벨 숫자 파싱"""
        import re

        levels = []
        # 숫자만 추출
        numbers = re.findall(r'\b([0-2])\b', response)

        for num in numbers:
            try:
                level = int(num)
                if 0 <= level <= 2:
                    levels.append(level)
            except ValueError:
                pass

        # 개수 맞추기
        while len(levels) < expected_count:
            levels.append(0)

        return levels[:expected_count]

    def format_with_analyzed_levels(self, text: str) -> FormatResult:
        """
        SDK로 레벨을 분석한 후 정규식으로 글머리 기호 적용

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

        # SDK로 레벨 분석
        levels = self.analyze_levels(text)

        # 정규식으로 글머리 기호 적용
        result = self._regex_formatter.format_text(text, levels=levels, auto_detect=False)

        if result.success:
            result.changes = ["SDK로 레벨 분석 후 글머리 기호 적용"]

        return result

    def auto_format(self, text: str, use_sdk_for_levels: bool = True) -> FormatResult:
        """
        자동으로 레벨을 분석하고 글머리 기호 적용

        Args:
            text: 변환할 텍스트
            use_sdk_for_levels: SDK로 레벨 분석 여부 (False면 정규식만 사용)

        Returns:
            FormatResult: 변환 결과
        """
        if use_sdk_for_levels:
            return self.format_with_analyzed_levels(text)
        else:
            return self._regex_formatter.auto_format(text)

    # 정규식 포맷터 메서드 위임
    def has_bullet(self, text: str) -> bool:
        """텍스트에 글머리 기호가 있는지 확인"""
        return self._regex_formatter.has_bullet(text)

    def parse_items(self, text: str) -> List[BulletItem]:
        """글머리 기호 목록 파싱"""
        return self._regex_formatter.parse_items(text)

    def normalize_style(self, text: str) -> FormatResult:
        """기존 글머리 기호를 현재 스타일로 통일"""
        return self._regex_formatter.normalize_style(text)

    def convert_style(self, text: str, target_style: str) -> FormatResult:
        """텍스트의 글머리 기호를 다른 스타일로 변환"""
        return self._regex_formatter.convert_style(text, target_style)

    def set_style(self, style: str):
        """글머리 스타일 변경"""
        self.style = style
        self._regex_formatter.set_style(style)

    # 내부 메서드 위임 (하위 호환성)
    def _detect_bullet_level(self, line: str) -> int:
        """기존 글머리 기호에서 레벨 감지"""
        return self._regex_formatter._detect_bullet_level(line)

    def _remove_existing_bullet(self, line: str):
        """기존 글머리 기호 제거"""
        return self._regex_formatter._remove_existing_bullet(line)
