# -*- coding: utf-8 -*-
"""
정규식 기반 포맷터 모듈

글머리 기호, 캡션 등의 양식 변환을 정규식으로 처리합니다.
SDK 사용이 필요한 경우 agent 모듈을 사용하세요.

포맷터 구조:
    BaseFormatter (추상 인터페이스)
        ├── BulletFormatter (글머리 기호)
        └── (확장 가능: NumberedFormatter, CustomFormatter 등)

    개요(outline)와 본문(add_)에 각각 다른 formatter를 적용할 수 있습니다.

YAML 설정 파일 기반 포맷팅:
    from merge.formatters import load_config, BulletFormatter, CaptionFormatter

    # 설정 로드
    config = load_config("formatter_config.yaml")

    # 글머리 포맷터
    bullet_fmt = BulletFormatter(style=config.bullet.style)

    # 캡션 포맷터
    caption_fmt = CaptionFormatter()
"""

from .base_formatter import (
    BaseFormatter,
    FormatResult,
)
from .bullet_formatter import (
    BulletFormatter,
    BulletItem,
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
from .object_formatter import (
    ObjectFormatter,
)
from .config_loader import (
    ConfigLoader,
    FormatterConfig,
    BulletConfig,
    CaptionConfig,
    CaptionFormatPreset,
    TableCaptionConfig,
    ImageCaptionConfig,
    load_config,
    create_default_config,
    save_default_config,
)

__all__ = [
    # 기본 인터페이스
    'BaseFormatter',
    'FormatResult',

    # 글머리 기호 변환
    'BulletFormatter',
    'BulletItem',
    'BULLET_STYLES',

    # 캡션 변환
    'CaptionFormatter',
    'CaptionInfo',
    'CAPTION_PATTERNS',
    'get_captions',
    'print_captions',
    'renumber_captions',

    # 객체 서식 설정
    'ObjectFormatter',

    # YAML 설정 로더
    'ConfigLoader',
    'FormatterConfig',
    'BulletConfig',
    'CaptionConfig',
    'CaptionFormatPreset',
    'TableCaptionConfig',
    'ImageCaptionConfig',
    'load_config',
    'create_default_config',
    'save_default_config',
]
