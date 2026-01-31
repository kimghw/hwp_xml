# -*- coding: utf-8 -*-
"""
Claude Code SDK Agent 모듈

Claude Code SDK를 사용한 텍스트 처리 기능을 제공합니다.
정규식만 사용하려면 merge.formatters 모듈을 사용하세요.

사용 예:
    from agent import ClaudeSDK, BulletFormatter, CaptionFormatter

    # 기본 SDK 호출
    sdk = ClaudeSDK()
    result = sdk.call("텍스트를 분석해주세요")

    # 글머리 기호 변환 (SDK 사용)
    formatter = BulletFormatter()
    formatted = formatter.format_text("항목1\\n항목2\\n항목3")

    # 캡션 조회 및 변환 (SDK 사용)
    caption_formatter = CaptionFormatter()
    captions = caption_formatter.get_all_captions("document.hwpx")
    result = caption_formatter.format_caption("그림 1. 테스트", caption_type="figure")

정규식만 사용:
    from merge.formatters import BulletFormatter, CaptionFormatter
"""

from .sdk import ClaudeSDK
from .bullet_formatter import BulletFormatter, BULLET_STYLES
from .caption_formatter import (
    CaptionFormatter,
    CaptionInfo,
    get_captions,
    print_captions,
    renumber_captions,
    CAPTION_PROMPTS,
)

# 공통 데이터 타입 (정규식 모듈에서 가져옴)
from merge.formatters.bullet_formatter import FormatResult, BulletItem
from merge.formatters.caption_formatter import CAPTION_PATTERNS

__all__ = [
    # SDK
    'ClaudeSDK',

    # 글머리 기호 변환 (SDK 기반)
    'BulletFormatter',
    'BULLET_STYLES',
    'FormatResult',
    'BulletItem',

    # 캡션 변환 (SDK 기반)
    'CaptionFormatter',
    'CaptionInfo',
    'get_captions',
    'print_captions',
    'renumber_captions',
    'CAPTION_PROMPTS',
    'CAPTION_PATTERNS',
]
