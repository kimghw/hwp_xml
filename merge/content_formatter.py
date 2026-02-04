# -*- coding: utf-8 -*-
"""
개요 내용 양식 변환 모듈 (SDK 레벨 분석 래퍼)

이 모듈은 SDK 기반 레벨 분석 기능을 제공합니다.
- 기본 포맷팅: formatters/bullet_formatter.py (BaseFormatter 상속)
- SDK 레벨 분석: agent.BulletFormatter

SDK 없이 포맷팅만 필요하면 formatters.BulletFormatter를 직접 사용하세요:
    from merge.formatters import BulletFormatter
    formatter = BulletFormatter(style="default")
    result = formatter.format_text(text)

SDK 레벨 분석이 필요하면 이 모듈을 사용:
    from merge.content_formatter import ContentFormatter
    formatter = ContentFormatter(use_sdk=True)
    result = formatter.format_with_analyzed_levels(text)
"""

import re
from typing import List, Dict, Optional, Tuple, Any

# 정규식 기반 포맷터
from .formatters.bullet_formatter import (
    BulletFormatter as RegexBulletFormatter,
    FormatResult,
    BulletItem,
    BULLET_STYLES,
)

# SDK 기반 포맷터 (선택적)
try:
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from agent import BulletFormatter as SDKBulletFormatter
    SDK_AVAILABLE = True
except ImportError:
    SDK_AVAILABLE = False
    SDKBulletFormatter = None


class ContentFormatter:
    """
    개요 내용 양식 변환기

    정규식 기반 포맷터를 기본으로 사용하고,
    use_sdk=True일 때 SDK 기반 포맷터를 사용합니다.
    """

    def __init__(self, style: str = "default", use_sdk: bool = True):
        """
        Args:
            style: 글머리 스타일 ("default", "filled", "numbered")
            use_sdk: Claude Code SDK 사용 여부
        """
        self.style = style
        self.use_sdk = use_sdk and SDK_AVAILABLE

        # 정규식 기반 포맷터 (항상 사용 가능)
        self._regex_formatter = RegexBulletFormatter(style=style)

        # SDK 기반 포맷터 (선택적)
        if self.use_sdk and SDKBulletFormatter:
            self._sdk_formatter = SDKBulletFormatter(style=style)
        else:
            self._sdk_formatter = None

    def format_as_bullet_list(
        self,
        text: str,
        levels: Optional[List[int]] = None,
        auto_detect_levels: bool = True
    ) -> FormatResult:
        """
        텍스트를 글머리 기호 목록으로 변환 (정규식 기반)

        Args:
            text: 변환할 텍스트 (줄바꿈으로 구분된 항목들)
            levels: 각 줄의 레벨 지정 [0, 1, 2, ...], None이면 자동 감지
            auto_detect_levels: 자동 레벨 감지 여부

        Returns:
            FormatResult: 변환 결과
        """
        return self._regex_formatter.format_text(text, levels, auto_detect_levels)

    def format_with_sdk(
        self,
        text: str,
        instruction: Optional[str] = None
    ) -> FormatResult:
        """
        Claude Code SDK를 사용하여 텍스트 양식 변환

        Args:
            text: 변환할 텍스트
            instruction: 추가 지시사항 (선택)

        Returns:
            FormatResult: 변환 결과
        """
        if not self._sdk_formatter:
            return self.format_as_bullet_list(text)

        # instruction이 있으면 임시로 formatter에 설정
        if instruction:
            original_instruction = self._sdk_formatter.instruction
            self._sdk_formatter.instruction = instruction
            result = self._sdk_formatter.format_text(text)
            self._sdk_formatter.instruction = original_instruction
            return result

        return self._sdk_formatter.format_text(text)

    def analyze_levels_with_sdk(self, text: str, existing_format: Optional[str] = None) -> List[int]:
        """
        SDK를 사용하여 내용 문단의 계층 레벨 분석

        개요 아래 내용 문단들의 의미적 계층 구조를 분석합니다.

        Args:
            text: 분석할 텍스트 (줄바꿈으로 구분)
            existing_format: 기존 문서의 포맷 예시 (참고용)

        Returns:
            각 줄의 레벨 리스트 [0, 1, 2, ...]
        """
        if not self._sdk_formatter:
            # SDK 없으면 정규식 기반 분석
            lines = [l for l in text.split('\n') if l.strip()]
            return self._regex_formatter._detect_levels(lines)

        return self._sdk_formatter.analyze_levels(text, existing_format)

    def format_with_analyzed_levels(self, text: str, existing_format: Optional[str] = None) -> FormatResult:
        """
        SDK로 레벨 분석 + 글머리 제거 후 정규식으로 글머리 기호 적용

        1. SDK로 레벨 분석 및 기존 글머리 제거
        2. 순수 텍스트에 분석된 레벨로 글머리 기호 적용

        Args:
            text: 변환할 텍스트
            existing_format: 기존 문서의 포맷 예시 (참고용)

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

        # SDK로 레벨 분석 + 글머리 제거
        if self._sdk_formatter and hasattr(self._sdk_formatter, 'analyze_and_strip'):
            levels, stripped_texts = self._sdk_formatter.analyze_and_strip(text, existing_format)
            if stripped_texts:
                # 정규식으로 글머리 기호 적용 (format_with_levels 사용)
                result = self._regex_formatter.format_with_levels(stripped_texts, levels)
                if result.success:
                    result.changes = ["SDK로 레벨 분석 및 글머리 제거 후 재적용"]
                return result

        # SDK 없으면 기존 방식 (레벨만 분석)
        levels = self.analyze_levels_with_sdk(text, existing_format)
        result = self._regex_formatter.format_text(text, levels=levels, auto_detect=False)

        if result.success:
            result.changes = ["SDK로 레벨 분석 후 글머리 기호 적용"]

        return result

    def auto_format(self, text: str, use_sdk_for_levels: bool = True) -> FormatResult:
        """
        자동으로 레벨을 분석하고 글머리 기호 적용

        Args:
            text: 변환할 텍스트
            use_sdk_for_levels: SDK로 레벨 분석 여부

        Returns:
            FormatResult: 변환 결과
        """
        if use_sdk_for_levels and self._sdk_formatter:
            return self.format_with_analyzed_levels(text)
        else:
            return self._regex_formatter.auto_format(text)

    def parse_bullet_list(self, text: str) -> List[BulletItem]:
        """글머리 기호 목록 파싱"""
        return self._regex_formatter.parse_items(text)


class OutlineContentFormatter:
    """
    개요 내용 전용 포맷터

    개요 트리의 내용 문단(개요로 선택되지 않은 문단)을 글머리 양식으로 변환합니다.
    SDK를 사용하면 내용의 의미적 계층 구조를 분석하여 적절한 레벨을 지정합니다.
    """

    def __init__(self, formatter: Optional[ContentFormatter] = None):
        self.formatter = formatter or ContentFormatter()

    def format_outline_content(
        self,
        outline_tree: List[Any],
        use_sdk: bool = False,
        use_sdk_for_levels: bool = True
    ) -> Tuple[List[Any], List[Dict]]:
        """
        개요 트리의 내용을 글머리 양식으로 변환

        Args:
            outline_tree: 개요 트리 (OutlineNode 리스트)
            use_sdk: Claude SDK로 전체 변환 수행 여부
            use_sdk_for_levels: SDK로 레벨만 분석 (use_sdk=False일 때 적용)

        Returns:
            (변환된 트리, 변경 내역)
        """
        changes = []

        for node in outline_tree:
            # 내용 문단 처리 (개요 자체 제외)
            content_paras = node.get_content_paragraphs() if hasattr(node, 'get_content_paragraphs') else []

            # 여러 내용 문단을 모아서 한 번에 레벨 분석
            non_bullet_paras = []
            for i, para in enumerate(content_paras):
                if not para.text or not para.text.strip():
                    continue
                # 이미 글머리 양식인지 확인
                if not self._has_bullet(para.text):
                    non_bullet_paras.append((i, para))

            if non_bullet_paras:
                # 양식 변환
                if use_sdk:
                    # SDK로 전체 변환
                    for i, para in non_bullet_paras:
                        result = self.formatter.format_with_sdk(para.text)
                        if result.success and result.formatted_text != para.text:
                            changes.append({
                                'outline': node.name if hasattr(node, 'name') else '',
                                'para_index': para.index if hasattr(para, 'index') else i,
                                'original': para.text,
                                'formatted': result.formatted_text
                            })
                            para.text = result.formatted_text
                elif use_sdk_for_levels:
                    # SDK로 레벨만 분석, 정규식으로 적용
                    # 모든 내용 문단을 합쳐서 레벨 분석
                    combined_text = '\n'.join(para.text for _, para in non_bullet_paras)
                    levels = self.formatter.analyze_levels_with_sdk(combined_text)

                    # 각 문단에 분석된 레벨 적용
                    level_idx = 0
                    for i, para in non_bullet_paras:
                        para_lines = [l for l in para.text.split('\n') if l.strip()]
                        para_levels = levels[level_idx:level_idx + len(para_lines)]
                        level_idx += len(para_lines)

                        # 레벨이 부족하면 0으로 채움
                        while len(para_levels) < len(para_lines):
                            para_levels.append(0)

                        result = self.formatter.format_as_bullet_list(
                            para.text,
                            levels=para_levels
                        )

                        if result.success and result.formatted_text != para.text:
                            changes.append({
                                'outline': node.name if hasattr(node, 'name') else '',
                                'para_index': para.index if hasattr(para, 'index') else i,
                                'original': para.text,
                                'formatted': result.formatted_text,
                                'analyzed_levels': para_levels
                            })
                            para.text = result.formatted_text
                else:
                    # 정규식만 사용
                    base_level = min(node.level, 2) if hasattr(node, 'level') else 0
                    for i, para in non_bullet_paras:
                        result = self.formatter.format_as_bullet_list(
                            para.text,
                            levels=[base_level]
                        )
                        if result.success and result.formatted_text != para.text:
                            changes.append({
                                'outline': node.name if hasattr(node, 'name') else '',
                                'para_index': para.index if hasattr(para, 'index') else i,
                                'original': para.text,
                                'formatted': result.formatted_text
                            })
                            para.text = result.formatted_text

            # 하위 개요 재귀 처리
            if hasattr(node, 'children') and node.children:
                _, child_changes = self.format_outline_content(
                    node.children, use_sdk, use_sdk_for_levels
                )
                changes.extend(child_changes)

        return outline_tree, changes

    def _has_bullet(self, text: str) -> bool:
        """텍스트에 글머리 기호가 있는지 확인"""
        return self.formatter._regex_formatter.has_bullet(text)


def format_text_interactive(text: str) -> str:
    """
    대화형으로 텍스트 양식 변환 (Claude Code SDK 사용)

    Args:
        text: 변환할 텍스트

    Returns:
        변환된 텍스트
    """
    formatter = ContentFormatter(use_sdk=True)
    result = formatter.format_with_sdk(text)

    if result.success:
        return result.formatted_text
    else:
        print(f"변환 실패: {result.errors}")
        return text


if __name__ == "__main__":
    # 테스트
    test_text = """첫 번째 항목입니다
두 번째 항목입니다
세 번째 항목입니다"""

    formatter = ContentFormatter(style="default", use_sdk=False)

    # 기본 변환 테스트
    result = formatter.format_as_bullet_list(test_text, levels=[0, 1, 2])
    print("=== 기본 변환 결과 ===")
    print(result.formatted_text)
    print()
    print("변경 내역:", result.changes)
