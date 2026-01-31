# -*- coding: utf-8 -*-
"""
Claude Code SDK Agent 모듈

SDK 역할: 분석만 수행 (레벨 분석, 제목 추출)
정규식 역할: 실제 포맷 적용

사용 예:
    from agent import ClaudeSDK, BulletFormatter, CaptionFormatter

    # 글머리 레벨 분석 (SDK) → 포맷 적용은 사용자 결정
    bullet = BulletFormatter()
    levels = bullet.analyze_levels("항목1\\n항목2\\n항목3")  # [0, 1, 2]

    # 캡션 제목 추출 (SDK) → 포맷 적용은 정규식
    caption = CaptionFormatter()
    title = caption.extract_title_with_sdk("Figure3-Sample Data")  # "Sample Data"
    result = caption.format_caption("그림 1. 테스트")  # 정규식으로 포맷

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
