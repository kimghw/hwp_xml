# -*- coding: utf-8 -*-
"""
글머리 기호 양식 변환 모듈 (Claude SDK 기반)

Claude Code SDK를 사용하여 텍스트를 글머리 기호 양식으로 변환합니다.
정규식만 사용하려면 merge.formatters.BulletFormatter를 사용하세요.

예: "항목1\\n항목2\\n항목3" → " □ 항목1\\n   ○항목2\\n    - 항목3"
"""

from dataclasses import dataclass, field
from typing import List, Optional
import yaml

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


# YAML에서 프롬프트 로드
def _load_prompts() -> dict:
    """bullet_formatter.yaml에서 프롬프트 로드"""
    yaml_path = Path(__file__).parent / "bullet_formatter.yaml"
    if yaml_path.exists():
        with open(yaml_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    return {}

_PROMPTS = _load_prompts()

# 프롬프트 상수 (YAML에서 로드, fallback 포함)
DEFAULT_INSTRUCTION = _PROMPTS.get('default_instruction', '')
LEVEL_ANALYSIS_INSTRUCTION = _PROMPTS.get('level_analysis_instruction', '')
LEVEL_ANALYSIS_WITH_CONTEXT_INSTRUCTION = _PROMPTS.get('level_analysis_with_context_instruction', '')
BODY_ANALYZE_STRIP_INSTRUCTION = _PROMPTS.get('body_analyze_strip_instruction', '')
TABLE_ANALYZE_STRIP_INSTRUCTION = _PROMPTS.get('table_analyze_strip_instruction', '')


class BulletFormatter:
    """
    글머리 기호 양식 변환기 (Claude SDK 기반)

    Claude Code SDK를 사용하여 내용을 글머리 양식으로 변환합니다.
    SDK 호출 실패 시 정규식 기반 변환으로 fallback합니다.
    """

    def __init__(
        self,
        style: str = "default",
        instruction: Optional[str] = None,
        custom_styles: Optional[dict] = None,
        context: str = "body"
    ):
        """
        Args:
            style: 글머리 스타일 ("default", "filled", "numbered")
            instruction: 커스텀 프롬프트 (None이면 기본 프롬프트 사용)
            custom_styles: YAML에서 로드한 커스텀 스타일 {style_name: {level: (symbol, indent)}}
            context: 사용 문맥 ("body": 본문, "table": 테이블 셀)
        """
        self.style = style
        self.instruction = instruction or DEFAULT_INSTRUCTION
        self.context = context
        self.sdk = ClaudeSDK(timeout=30)
        self._custom_styles = custom_styles
        # 정규식 기반 포맷터 (fallback용, 커스텀 스타일 전달)
        self._regex_formatter = RegexBulletFormatter(style=style, custom_styles=custom_styles)

        # 문맥에 따른 analyze_and_strip 프롬프트 선택
        if context == "table":
            self._analyze_strip_instruction = TABLE_ANALYZE_STRIP_INSTRUCTION
        else:
            self._analyze_strip_instruction = BODY_ANALYZE_STRIP_INSTRUCTION

    def format_text(self, text: str, auto_detect: bool = True) -> FormatResult:
        """
        텍스트를 글머리 기호 양식으로 변환 (SDK 레벨 분석 + 정규식 적용)

        SDK로 레벨만 분석하고, 글머리 기호는 정규식 포맷터로 적용합니다.
        이렇게 하면 YAML 설정의 indent가 정확히 적용됩니다.

        Args:
            text: 변환할 텍스트 (줄바꿈으로 구분된 항목들)
            auto_detect: 자동 레벨 감지 (SDK에서는 항상 SDK 기반 분석 사용)

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

        # SDK로 레벨 분석 + 기존 글머리 제거
        levels, stripped_texts = self.analyze_and_strip(text)

        if levels and stripped_texts:
            # 정규식 포맷터로 글머리 기호 적용 (YAML 설정의 indent 포함)
            result = self._regex_formatter.format_with_levels(stripped_texts, levels)
            if result.success:
                print(f"    [BulletFormatter] SDK 성공")
                return result

        # SDK 실패 시 정규식 기반 변환 사용
        print(f"    [BulletFormatter] SDK FALLBACK → 정규식 사용")
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
        text = text.rstrip()  # 앞 공백(indent) 유지

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

        # 앞 공백(indent)은 유지, 뒤 공백만 제거
        return text.rstrip()

    def analyze_and_strip(self, text: str, existing_format: Optional[str] = None) -> tuple:
        """
        SDK를 사용하여 텍스트의 레벨 분석 + 기존 글머리 제거

        Args:
            text: 분석할 텍스트
            existing_format: 기존 문서의 포맷 예시 (참고용)

        Returns:
            tuple: (레벨 리스트, 글머리 제거된 텍스트 리스트)
                예: ([0, 1, 2], ["안녕하세요", "세부내용", "상세항목"])
        """
        if not text or not text.strip():
            return [], []

        lines = [l for l in text.split('\n') if l.strip()]
        if not lines:
            return [], []

        # 문맥에 따른 프롬프트 사용 (본문/테이블)
        prompt = self._analyze_strip_instruction.format(text=text)
        result = self.sdk.call(prompt)

        if result.success and result.output:
            return self._parse_analyze_strip_response(result.output, len(lines))
        else:
            # SDK 실패 시 정규식 fallback
            levels = self._regex_formatter._detect_levels(lines)
            stripped_texts = []
            for line in lines:
                clean_text, _ = self._regex_formatter._remove_existing_bullet(line)
                stripped_texts.append(clean_text.strip())
            return levels, stripped_texts

    def _parse_analyze_strip_response(self, response: str, expected_count: int) -> tuple:
        """SDK 응답에서 레벨과 텍스트 파싱"""
        levels = []
        stripped_texts = []

        # 코드 블록 제거
        text = response.strip()
        if '```' in text:
            # 코드 블록 내용만 추출
            import re
            code_match = re.search(r'```\n?(.*?)\n?```', text, re.DOTALL)
            if code_match:
                text = code_match.group(1)
            else:
                # 코드 블록 마커만 제거
                text = text.replace('```', '')

        for line in text.split('\n'):
            line = line.strip()
            if '|' in line:
                parts = line.split('|', 1)
                try:
                    level = int(parts[0].strip())
                    level = max(0, min(level, 2))
                    levels.append(level)
                    stripped_texts.append(parts[1].strip() if len(parts) > 1 else '')
                except ValueError:
                    continue

        # 개수 맞추기
        while len(levels) < expected_count:
            levels.append(0)
            stripped_texts.append('')

        return levels[:expected_count], stripped_texts[:expected_count]

    def analyze_levels(self, text: str, existing_format: Optional[str] = None) -> List[int]:
        """
        SDK를 사용하여 텍스트의 각 줄에 대한 계층 레벨 분석

        Args:
            text: 분석할 텍스트
            existing_format: 기존 문서의 포맷 예시 (참고용)
                예: " □ 주요 내용\\n   ○세부 설명\\n    - 상세 항목"

        Returns:
            각 줄의 레벨 리스트 [0, 1, 2, ...]
        """
        if not text or not text.strip():
            return []

        lines = [l for l in text.split('\n') if l.strip()]
        if not lines:
            return []

        # 기존 포맷이 있으면 컨텍스트 포함 프롬프트 사용
        if existing_format:
            prompt = LEVEL_ANALYSIS_WITH_CONTEXT_INSTRUCTION.format(
                existing_format=existing_format,
                text=text
            )
        else:
            prompt = LEVEL_ANALYSIS_INSTRUCTION.format(text=text)

        result = self.sdk.call(prompt)

        if result.success and result.output:
            levels = self._parse_level_response(result.output, len(lines))
            print(f"    [BulletFormatter] analyze_levels SDK 성공")
            return levels
        else:
            # SDK 실패 시 정규식 기반 분석
            print(f"    [BulletFormatter] analyze_levels FALLBACK → 정규식 (사유: {result.error})")
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

    # BaseFormatter 인터페이스 구현
    def get_style_name(self) -> str:
        """현재 스타일 이름 반환"""
        return self.style

    def has_format(self, text: str) -> bool:
        """텍스트에 이미 글머리 포맷이 적용되어 있는지 확인"""
        return self.has_bullet(text)
