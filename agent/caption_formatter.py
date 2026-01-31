# -*- coding: utf-8 -*-
"""
캡션 양식 변환 모듈 (Claude SDK 기반)

Claude Code SDK를 사용하여 캡션을 특정 양식으로 변환합니다.
정규식만 사용하려면 merge.formatters.CaptionFormatter를 사용하세요.

사용 예:
    from agent import CaptionFormatter

    formatter = CaptionFormatter()

    # 캡션 양식 변환 (SDK 사용)
    result = formatter.format_caption("그림 1. 테스트 이미지", caption_type="figure")

    # 대괄호 형식 변환 (SDK 사용)
    result = formatter.to_bracket_format("표 1. 연구 결과")
"""

from typing import List, Optional, Tuple, Dict

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


# 캡션 유형별 SDK 프롬프트
CAPTION_PROMPTS = {
    "figure": """다음 그림 캡션을 표준 양식으로 변환해주세요.

양식 규칙:
- 형식: "그림 [번호]. [설명]" 또는 "Figure [번호]. [설명]"
- 번호는 아라비아 숫자 사용
- 설명은 간결하게 작성
- 마침표로 끝나지 않음 (제목이므로)

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.
3. 원본 내용의 의미를 유지하세요.""",

    "table": """다음 표 캡션을 표준 양식으로 변환해주세요.

양식 규칙:
- 형식: "표 [번호]. [설명]" 또는 "Table [번호]. [설명]"
- 번호는 아라비아 숫자 사용
- 설명은 간결하게 작성
- 마침표로 끝나지 않음 (제목이므로)

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.
3. 원본 내용의 의미를 유지하세요.""",

    "equation": """다음 수식 캡션을 표준 양식으로 변환해주세요.

양식 규칙:
- 형식: "식 ([번호])" 또는 "Equation ([번호])"
- 번호는 아라비아 숫자 또는 장.절 형식 (예: 2.1)
- 수식 설명이 있다면 유지

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.""",

    "default": """다음 캡션을 표준 양식으로 변환해주세요.

양식 규칙:
- 형식: "[유형] [번호]. [설명]"
- 번호는 아라비아 숫자 사용
- 설명은 간결하게 작성

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.""",

    # 번호 없이 제목만 추출하는 프롬프트
    "title_only": """다음 캡션에서 제목(설명)만 추출해주세요.

규칙:
- "그림 1. 테스트 이미지" → "테스트 이미지"
- "표 2. 연구 결과" → "연구 결과"
- "Figure 3. Sample Data" → "Sample Data"
- 번호와 유형 키워드(그림, 표, Figure, Table 등)를 제거
- 설명 부분만 반환

중요:
1. 제목만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.""",

    # 대괄호 형식 프롬프트
    "bracket_title": """다음 캡션을 대괄호 제목 형식으로 변환해주세요.

규칙:
- "그림 1. 테스트 이미지" → "[그림 테스트 이미지]"
- "표 2. 연구 결과" → "[표 연구 결과]"
- "Figure 3. Sample Data" → "[Figure Sample Data]"
- 번호를 제거하고 대괄호로 감싸기
- 유형 키워드는 유지

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요.
3. 원본 내용의 의미를 유지하세요.""",
}


class CaptionFormatter:
    """
    캡션 양식 변환기 (Claude SDK 기반)

    Claude Code SDK를 사용하여 캡션 양식을 변환합니다.
    SDK 호출 실패 시 정규식 기반 변환으로 fallback합니다.
    """

    def __init__(self, custom_prompts: Optional[Dict[str, str]] = None):
        """
        Args:
            custom_prompts: 캡션 유형별 커스텀 프롬프트 (기본 프롬프트 오버라이드)
        """
        self.sdk = ClaudeSDK(timeout=30)
        self.prompts = {**CAPTION_PROMPTS}
        if custom_prompts:
            self.prompts.update(custom_prompts)
        # 정규식 기반 포맷터 (fallback 및 조회용)
        self._regex_formatter = RegexCaptionFormatter()

    def get_all_captions(self, hwpx_path: str) -> List[CaptionInfo]:
        """
        HWPX 파일에서 모든 캡션 조회

        Args:
            hwpx_path: HWPX 파일 경로

        Returns:
            CaptionInfo 리스트
        """
        return self._regex_formatter.get_all_captions(hwpx_path)

    def format_caption(
        self,
        text: str,
        caption_type: str = "default",
        custom_instruction: Optional[str] = None
    ) -> FormatResult:
        """
        캡션을 지정된 양식으로 변환 (SDK 사용)

        Args:
            text: 변환할 캡션 텍스트
            caption_type: 캡션 유형 ("figure", "table", "equation", "default")
            custom_instruction: 커스텀 프롬프트 (지정 시 기본 프롬프트 대체)

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

        # 프롬프트 선택
        instruction = custom_instruction or self.prompts.get(caption_type, self.prompts["default"])

        prompt = f"""{instruction}

변환할 캡션:
{text}
"""

        result = self.sdk.call(prompt)

        if result.success and result.output:
            cleaned = self._clean_response(result.output)
            return FormatResult(
                success=True,
                original_text=text,
                formatted_text=cleaned,
                changes=[f"캡션 양식 변환 완료 (유형: {caption_type})"]
            )
        else:
            return FormatResult(
                success=False,
                original_text=text,
                formatted_text=text,
                errors=[result.error or "SDK 호출 실패"]
            )

    def format_all_captions(
        self,
        captions: List[CaptionInfo],
        custom_instruction: Optional[str] = None
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 일괄 변환 (SDK 사용)

        Args:
            captions: CaptionInfo 리스트
            custom_instruction: 커스텀 프롬프트 (모든 캡션에 적용)

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []

        for caption in captions:
            result = self.format_caption(
                caption.text,
                caption.caption_type,
                custom_instruction
            )
            results.append((caption, result))

        return results

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

        # 첫 줄만 사용 (캡션은 보통 한 줄)
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if lines:
            text = lines[0]

        return text.strip()

    def set_prompt(self, caption_type: str, prompt: str):
        """
        특정 캡션 유형의 프롬프트 설정

        Args:
            caption_type: 캡션 유형
            prompt: 프롬프트 텍스트
        """
        self.prompts[caption_type] = prompt

    def get_prompt(self, caption_type: str) -> str:
        """
        특정 캡션 유형의 프롬프트 조회

        Args:
            caption_type: 캡션 유형

        Returns:
            프롬프트 텍스트
        """
        return self.prompts.get(caption_type, self.prompts["default"])

    def extract_title(self, text: str, use_sdk: bool = True) -> str:
        """
        캡션에서 제목(설명)만 추출 (번호 제거)

        Args:
            text: 캡션 텍스트
            use_sdk: SDK 사용 여부

        Returns:
            제목만 추출된 텍스트
        """
        if use_sdk:
            result = self.format_caption(text, "title_only")
            if result.success:
                return result.formatted_text

        # SDK 미사용 또는 실패 시 정규식으로 처리
        return self._regex_formatter.extract_title(text)

    def to_bracket_format(
        self,
        text: str,
        caption_type: str = "default",
        use_sdk: bool = True
    ) -> FormatResult:
        """
        캡션을 대괄호 형식으로 변환 (번호 제거)
        예: "표 1. 연구 결과" → "[표 연구 결과]"

        Args:
            text: 캡션 텍스트
            caption_type: 캡션 유형
            use_sdk: SDK 사용 여부

        Returns:
            FormatResult
        """
        if use_sdk:
            result = self.format_caption(text, "bracket_title")
            if result.success:
                return result

        # SDK 미사용 또는 실패 시 정규식으로 처리
        return self._regex_formatter.to_bracket_format(text, caption_type)

    def format_all_to_bracket(
        self,
        captions: List[CaptionInfo],
        use_sdk: bool = True
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        여러 캡션을 대괄호 형식으로 일괄 변환

        Args:
            captions: CaptionInfo 리스트
            use_sdk: SDK 사용 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        results = []
        for caption in captions:
            result = self.to_bracket_format(caption.text, caption.caption_type, use_sdk)
            results.append((caption, result))
        return results

    # 정규식 포맷터 메서드 위임
    def renumber_captions(
        self,
        captions: List[CaptionInfo],
        by_type: bool = True,
        start_number: int = 1
    ) -> List[CaptionInfo]:
        """캡션 번호 재정렬"""
        return self._regex_formatter.renumber_captions(captions, by_type, start_number)

    def apply_new_numbers(
        self,
        captions: List[CaptionInfo],
        use_sdk: bool = False
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        재정렬된 번호를 캡션 텍스트에 적용

        Args:
            captions: renumber_captions()로 번호가 재정렬된 CaptionInfo 리스트
            use_sdk: True면 SDK로 양식 변환, False면 단순 번호 교체

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

            if use_sdk:
                # SDK로 번호 포함 양식 변환
                instruction = f"""다음 캡션의 번호를 {caption.new_number}로 변경하고 표준 양식으로 변환해주세요.

양식 규칙:
- 그림: "그림 {caption.new_number}. [설명]"
- 표: "표 {caption.new_number}. [설명]"
- 식: "식 ({caption.new_number})"

중요:
1. 변환된 캡션만 반환하세요.
2. 설명이나 코멘트를 추가하지 마세요."""

                result = self.format_caption(caption.text, caption.caption_type, instruction)
            else:
                # 단순 번호 교체 (정규식)
                new_text = self._regex_formatter.replace_number(
                    caption.text, caption.new_number, caption.caption_type
                )
                result = FormatResult(
                    success=True,
                    original_text=caption.text,
                    formatted_text=new_text,
                    changes=[f"번호 변경: {caption.number} → {caption.new_number}"] if caption.number != caption.new_number else []
                )

            results.append((caption, result))

        return results

    def renumber_and_format(
        self,
        hwpx_path: str,
        by_type: bool = True,
        start_number: int = 1,
        use_sdk: bool = False
    ) -> List[Tuple[CaptionInfo, FormatResult]]:
        """
        HWPX 파일의 캡션을 조회하고 번호를 재정렬하여 적용 (일괄 처리)

        Args:
            hwpx_path: HWPX 파일 경로
            by_type: True면 유형별로 따로 번호 매김
            start_number: 시작 번호
            use_sdk: SDK로 양식 변환 여부

        Returns:
            (CaptionInfo, FormatResult) 튜플 리스트
        """
        # 캡션 조회
        captions = self.get_all_captions(hwpx_path)

        if not captions:
            return []

        # 번호 재정렬
        renumbered = self.renumber_captions(captions, by_type, start_number)

        # 새 번호 적용
        return self.apply_new_numbers(renumbered, use_sdk)

    def apply_to_hwpx(
        self,
        hwpx_path: str,
        output_path: Optional[str] = None,
        to_bracket: bool = False,
        use_sdk: bool = False,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 변환하여 저장

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            to_bracket: True면 대괄호 형식으로 변환 "[표 제목]"
            use_sdk: SDK 사용 여부
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬 (1부터 시작)

        Returns:
            저장된 파일 경로
        """
        if use_sdk and to_bracket:
            # SDK 사용 시 커스텀 transform_func 생성
            def sdk_transform(text: str, caption_type: str) -> str:
                result = self.to_bracket_format(text, caption_type, use_sdk=True)
                return result.formatted_text

            return self._regex_formatter.apply_to_hwpx(
                hwpx_path, output_path,
                transform_func=sdk_transform,
                to_bracket=False,  # transform_func 사용하므로 False
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
        use_sdk: bool = False,
        keep_auto_num: bool = False,
        renumber: bool = False
    ) -> str:
        """
        HWPX 파일의 캡션을 대괄호 형식으로 변환하여 저장
        예: "표 1. 연구 결과" → "[표 연구 결과]"

        Args:
            hwpx_path: 원본 HWPX 파일 경로
            output_path: 저장 경로 (None이면 원본 덮어쓰기)
            use_sdk: SDK 사용 여부
            keep_auto_num: True면 자동 번호(autoNum) 유지
            renumber: True면 자동 번호 재정렬 (1부터 시작)

        Returns:
            저장된 파일 경로
        """
        return self.apply_to_hwpx(
            hwpx_path, output_path,
            to_bracket=True, use_sdk=use_sdk,
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
