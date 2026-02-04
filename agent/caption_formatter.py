# -*- coding: utf-8 -*-
"""
캡션 양식 변환 모듈 (Claude SDK 기반)

Claude Code SDK를 사용하여 캡션의 제목(이름)을 추출합니다.
포맷 적용은 정규식 기반으로 수행합니다.

SDK 역할:
- 캡션에서 제목 추출 (복잡한 패턴 처리)
- 캡션 유형 분석

정규식 역할:
- 실제 포맷 적용 ("그림 1. 제목", "[표 제목]" 등)
- 번호 재정렬

사용 예:
    from agent import CaptionFormatter

    formatter = CaptionFormatter()

    # SDK로 제목 추출, 정규식으로 포맷 적용
    title = formatter.extract_title_with_sdk("Figure3-Sample Data")  # "Sample Data"
    result = formatter.format_caption("그림 1. 테스트", caption_type="figure")
"""

from typing import List, Optional, Tuple, Dict
import yaml

from .sdk import ClaudeSDK

# 정규식 기반 포맷터 import
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from merge.formatters.caption_formatter import (
    CaptionFormatter as RegexCaptionFormatter,
    CaptionInfo,
    FormatResult,
    get_captions,
    print_captions,
    renumber_captions,
    CAPTION_PATTERNS,
)


# YAML에서 프롬프트 로드
def _load_prompts() -> dict:
    """caption_formatter.yaml에서 프롬프트 로드"""
    yaml_path = Path(__file__).parent / "caption_formatter.yaml"
    if yaml_path.exists():
        with open(yaml_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    return {}

_PROMPTS = _load_prompts()

# SDK 프롬프트 - 분석용
CAPTION_PROMPTS = {
    "extract_title": _PROMPTS.get('extract_title', ''),
    "analyze_type": _PROMPTS.get('analyze_type', ''),
}


class CaptionFormatter:
    """
    캡션 양식 변환기 (Claude SDK 기반)

    SDK 역할: 캡션에서 제목 추출 (복잡한 패턴 처리)
    정규식 역할: 실제 포맷 적용 및 번호 재정렬

    SDK 호출 실패 시 정규식 기반 처리로 fallback합니다.
    """

    def __init__(self, custom_prompts: Optional[Dict[str, str]] = None):
        """
        Args:
            custom_prompts: 커스텀 프롬프트 (기본 프롬프트 오버라이드)
        """
        self.sdk = ClaudeSDK(timeout=30)
        self.prompts = {**CAPTION_PROMPTS}
        if custom_prompts:
            self.prompts.update(custom_prompts)
        # 정규식 기반 포맷터 (실제 포맷 적용용)
        self._regex_formatter = RegexCaptionFormatter()

    # ========== SDK 분석 기능 (핵심) ==========

    def extract_title_with_sdk(self, text: str) -> str:
        """
        SDK를 사용하여 캡션에서 제목만 추출

        복잡한 패턴도 처리 가능:
        - "Figure3-Sample Data" → "Sample Data"
        - "Table1:Results" → "Results"

        Args:
            text: 캡션 텍스트

        Returns:
            추출된 제목 (실패 시 정규식으로 fallback)
        """
        if not text or not text.strip():
            return ""

        prompt = f"""{self.prompts["extract_title"]}

캡션:
{text}

제목:"""

        result = self.sdk.call(prompt)

        if result.success and result.output:
            title = self._clean_response(result.output)
            if title:
                return title

        # SDK 실패 시 정규식으로 fallback
        return self._regex_formatter.extract_title(text)

    def analyze_type_with_sdk(self, text: str) -> str:
        """
        SDK를 사용하여 캡션 유형 분석

        Args:
            text: 캡션 텍스트

        Returns:
            캡션 유형 ("figure", "table", "equation")
        """
        if not text or not text.strip():
            return ""

        prompt = self.prompts["analyze_type"].format(text=text)
        result = self.sdk.call(prompt)

        if result.success and result.output:
            type_str = self._clean_response(result.output).lower()
            if type_str in ("figure", "table", "equation"):
                return type_str

        # SDK 실패 시 정규식으로 fallback
        return self._regex_formatter._detect_caption_type(text, "")

    # ========== 포맷 적용 기능 (정규식 기반) ==========

    def format_caption(
        self,
        text: str,
        caption_type: str = "",
        use_sdk_for_title: bool = True
    ) -> FormatResult:
        """
        캡션을 표준 양식으로 변환

        SDK로 제목 추출 → 정규식으로 포맷 적용

        Args:
            text: 변환할 캡션 텍스트
            caption_type: 캡션 유형 (빈 문자열이면 자동 감지)
            use_sdk_for_title: SDK로 제목 추출 여부

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

        # 제목 추출
        if use_sdk_for_title:
            title = self.extract_title_with_sdk(text)
        else:
            title = self._regex_formatter.extract_title(text)

        # 유형 감지
        if not caption_type:
            caption_type = self._regex_formatter._detect_caption_type(text, "")

        # 번호 추출
        number = self._regex_formatter._extract_caption_number(text)

        # 정규식으로 포맷 적용
        formatted = self._regex_formatter._build_standard_caption(caption_type, number, title)

        changes = []
        if formatted != text:
            changes.append(f"표준 양식 변환 (SDK 제목 추출)" if use_sdk_for_title else "표준 양식 변환")

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=changes
        )

    def to_bracket_format(
        self,
        text: str,
        caption_type: str = "",
        use_sdk_for_title: bool = True
    ) -> FormatResult:
        """
        캡션을 대괄호 형식으로 변환
        예: "표 1. 연구 결과" → "[연구 결과]"

        대괄호 안에는 제목만 들어감 (유형 접두어 제외)

        SDK로 제목 추출 → 정규식으로 대괄호 포맷 적용

        Args:
            text: 캡션 텍스트
            caption_type: 캡션 유형
            use_sdk_for_title: SDK로 제목 추출 여부

        Returns:
            FormatResult
        """
        if not text or not text.strip():
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=text,
                changes=[]
            )

        # 제목 추출
        if use_sdk_for_title:
            title = self.extract_title_with_sdk(text)
        else:
            title = self._regex_formatter.extract_title(text)

        # 대괄호 형식 생성 (제목만)
        if title:
            formatted = f"[{title}]"
        else:
            formatted = text

        changes = []
        if formatted != text:
            changes.append("대괄호 형식 변환")

        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=changes
        )

    # ========== 일괄 처리 ==========

    def format_all_captions(
        self,
        captions: List[CaptionInfo],
        use_sdk_for_title: bool = True
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 일괄 변환

        Args:
            captions: CaptionInfo 리스트
            use_sdk_for_title: SDK로 제목 추출 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []
        for caption in captions:
            result = self.format_caption(
                caption.text,
                caption.caption_type,
                use_sdk_for_title
            )
            results.append((caption, result))
        return results

    def format_all_to_bracket(
        self,
        captions: List[CaptionInfo],
        use_sdk_for_title: bool = True
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 대괄호 형식으로 일괄 변환

        Args:
            captions: CaptionInfo 리스트
            use_sdk_for_title: SDK로 제목 추출 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []
        for caption in captions:
            result = self.to_bracket_format(
                caption.text,
                caption.caption_type,
                use_sdk_for_title
            )
            results.append((caption, result))
        return results

    # ========== 유틸리티 ==========

    def _clean_response(self, text: str) -> str:
        """SDK 응답 정리"""
        text = text.strip()

        # 코드 블록 제거
        if text.startswith('```') and text.endswith('```'):
            lines = text.split('\n')
            if len(lines) >= 2:
                lines = lines[1:-1]
            text = '\n'.join(lines)

        while '```' in text:
            text = text.replace('```', '')
        while '`' in text:
            text = text.replace('`', '')

        # 따옴표 제거
        text = text.strip('"\'')

        # 첫 줄만 사용 (캡션은 보통 한 줄)
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if lines:
            text = lines[0]

        return text.strip()

    def get_all_captions(self, hwpx_path: str) -> List[CaptionInfo]:
        """HWPX 파일에서 모든 캡션 조회"""
        return self._regex_formatter.get_all_captions(hwpx_path)

    # 하위 호환성을 위한 extract_title (기존 인터페이스 유지)
    def extract_title(self, text: str, use_sdk: bool = True) -> str:
        """캡션에서 제목 추출 (하위 호환용)"""
        if use_sdk:
            return self.extract_title_with_sdk(text)
        return self._regex_formatter.extract_title(text)

    # ========== 정규식 포맷터 위임 ==========

    def renumber_captions(
        self,
        captions: List[CaptionInfo],
        by_type: bool = True,
        start_number: int = 1
    ) -> List[CaptionInfo]:
        """캡션 번호 재정렬 (정규식)"""
        return self._regex_formatter.renumber_captions(captions, by_type, start_number)

    def apply_new_numbers(
        self,
        captions: List[CaptionInfo],
        use_sdk_for_title: bool = False
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        재정렬된 번호를 캡션 텍스트에 적용

        Args:
            captions: renumber_captions()로 번호가 재정렬된 CaptionInfo 리스트
            use_sdk_for_title: SDK로 제목 추출 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []

        for caption in captions:
            if caption.new_number is None:
                results.append((caption, FormatResult(
                    success=True,
                    original_text=caption.text,
                    formatted_text=caption.text,
                    changes=[]
                )))
                continue

            # 제목 추출
            if use_sdk_for_title:
                title = self.extract_title_with_sdk(caption.text)
            else:
                title = self._regex_formatter.extract_title(caption.text)

            # 정규식으로 새 번호 적용
            formatted = self._regex_formatter._build_standard_caption(
                caption.caption_type,
                caption.new_number,
                title
            )

            result = FormatResult(
                success=True,
                original_text=caption.text,
                formatted_text=formatted,
                changes=[f"번호 변경: {caption.number} → {caption.new_number}"] if caption.number != caption.new_number else []
            )

            results.append((caption, result))

        return results

    def renumber_and_format(
        self,
        hwpx_path: str,
        by_type: bool = True,
        start_number: int = 1,
        use_sdk_for_title: bool = False
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        HWPX 파일의 캡션을 조회하고 번호를 재정렬하여 적용

        Args:
            hwpx_path: HWPX 파일 경로
            by_type: True면 유형별로 따로 번호 매김
            start_number: 시작 번호
            use_sdk_for_title: SDK로 제목 추출 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        captions = self.get_all_captions(hwpx_path)
        if not captions:
            return []

        renumbered = self.renumber_captions(captions, by_type, start_number)
        return self.apply_new_numbers(renumbered, use_sdk_for_title)

    def apply_to_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        to_bracket: bool = False,
        use_sdk_for_title: bool = False,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 변환하여 저장

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            to_bracket: True면 대괄호 형식으로 변환
            use_sdk_for_title: SDK로 제목 추출 여부
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬

        Returns:
            저장된 파일 경로
        """
        if use_sdk_for_title:
            # SDK로 제목 추출 후 정규식으로 포맷 적용
            def sdk_transform(text: str, caption_type: str) -> str:
                if to_bracket:
                    result = self.to_bracket_format(text, caption_type, use_sdk_for_title=True)
                else:
                    result = self.format_caption(text, caption_type, use_sdk_for_title=True)
                return result.formatted_text

            return self._regex_formatter.apply_to_hwpx(
                hwpx_path, output_path,
                transform_func=sdk_transform,
                to_bracket=False,
                keep_auto_num=keep_auto_num,
                renumber=renumber
            )
        else:
            # 정규식만 사용
            return self._regex_formatter.apply_to_hwpx(
                hwpx_path, output_path,
                to_bracket=to_bracket,
                keep_auto_num=keep_auto_num,
                renumber=renumber
            )

    def apply_bracket_format(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        use_sdk_for_title: bool = False,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 대괄호 형식으로 변환하여 저장

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            use_sdk_for_title: SDK로 제목 추출 여부
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬

        Returns:
            저장된 파일 경로
        """
        return self.apply_to_hwpx(
            hwpx_path, output_path,
            to_bracket=True,
            use_sdk_for_title=use_sdk_for_title,
            keep_auto_num=keep_auto_num,
            renumber=renumber
        )

    def renumber_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """HWPX 파일의 캡션 번호만 재정렬"""
        return self._regex_formatter.renumber_hwpx(hwpx_path, output_path)


# 편의 함수들은 정규식 기반 모듈에서 재export
__all__ = [
    'CaptionFormatter',
    'CaptionInfo',
    'FormatResult',
    'CAPTION_PROMPTS',
    'CAPTION_PATTERNS',
    'get_captions',
    'print_captions',
    'renumber_captions',
]
