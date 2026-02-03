# -*- coding: utf-8 -*-
"""
포맷터 기본 인터페이스

모든 포맷터가 상속받아야 하는 추상 클래스를 정의합니다.
SDK에서 레벨/텍스트를 받아서 스타일만 적용하는 구조입니다.

사용 예:
    class MyFormatter(BaseFormatter):
        def format_with_levels(self, texts, levels):
            # 스타일 적용 로직
            pass
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class FormatResult:
    """형식 변환 결과"""
    success: bool = True
    original_text: str = ""
    formatted_text: str = ""
    changes: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class BaseFormatter(ABC):
    """
    포맷터 기본 인터페이스

    SDK에서 레벨 분석된 텍스트를 받아서 스타일을 적용합니다.
    개요(outline)와 본문(add_)에 각각 다른 포맷터를 적용할 수 있습니다.
    """

    @abstractmethod
    def format_with_levels(
        self,
        texts: List[str],
        levels: List[int]
    ) -> FormatResult:
        """
        레벨이 지정된 텍스트에 스타일 적용

        SDK에서 분석한 레벨과 텍스트를 받아서 포맷팅합니다.

        Args:
            texts: 각 줄의 텍스트 리스트 (글머리 제거된 순수 텍스트)
            levels: 각 줄의 레벨 리스트 [0, 1, 2, ...]

        Returns:
            FormatResult: 변환 결과
        """
        pass

    @abstractmethod
    def format_text(
        self,
        text: str,
        levels: Optional[List[int]] = None,
        auto_detect: bool = True
    ) -> FormatResult:
        """
        텍스트를 포맷팅 (레벨 자동 감지 가능)

        Args:
            text: 변환할 텍스트 (줄바꿈으로 구분)
            levels: 각 줄의 레벨, None이면 자동 감지
            auto_detect: 자동 레벨 감지 여부

        Returns:
            FormatResult: 변환 결과
        """
        pass

    def has_format(self, text: str) -> bool:
        """
        텍스트에 이미 포맷이 적용되어 있는지 확인

        Args:
            text: 확인할 텍스트

        Returns:
            포맷 적용 여부
        """
        return False

    def get_style_name(self) -> str:
        """현재 스타일 이름 반환"""
        return "base"
