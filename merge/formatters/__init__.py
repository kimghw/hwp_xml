# -*- coding: utf-8 -*-
"""
정규식 기반 포맷터 모듈

글머리 기호, 캡션 등의 양식 변환을 정규식으로 처리합니다.
SDK 사용이 필요한 경우 agent 모듈을 사용하세요.

YAML 설정 파일 기반 포맷팅:
    from merge.formatters import load_config, BulletFormatter, CaptionFormatter

    # 설정 로드
    config = load_config("formatter_config.yaml")

    # 글머리 포맷터
    bullet_fmt = BulletFormatter(style=config.bullet.style)

    # 캡션 포맷터
    caption_fmt = CaptionFormatter()
"""

from .bullet_formatter import (
    BulletFormatter,
    BulletItem,
    FormatResult,
    BULLET_STYLES,
)
from .caption_formatter import (
    CaptionFormatter,
    CaptionInfo,
    CAPTION_PATTERNS,
    get_captions,
    print_captions,
    renumber_captions,
)
from .config_loader import (
    ConfigLoader,
    FormatterConfig,
    BulletConfig,
    CaptionConfig,
    TableCaptionConfig,
    ImageCaptionConfig,
    load_config,
    create_default_config,
    save_default_config,
)

__all__ = [
    # 글머리 기호 변환
    'BulletFormatter',
    'BulletItem',
    'FormatResult',
    'BULLET_STYLES',

    # 캡션 변환
    'CaptionFormatter',
    'CaptionInfo',
    'CAPTION_PATTERNS',
    'get_captions',
    'print_captions',
    'renumber_captions',

    # YAML 설정 로더
    'ConfigLoader',
    'FormatterConfig',
    'BulletConfig',
    'CaptionConfig',
    'TableCaptionConfig',
    'ImageCaptionConfig',
    'load_config',
    'create_default_config',
    'save_default_config',
]
