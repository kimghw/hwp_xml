# -*- coding: utf-8 -*-
"""
YAML 설정 로더 유틸리티

포맷터 설정을 YAML 파일에서 로드하고 관리합니다.
"""

import yaml
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List, Tuple
from pathlib import Path


@dataclass
class BulletConfig:
    """글머리 기호 설정 (BulletFormatter 옵션)"""
    # 스타일: default, filled, numbered, arrow, checkbox, star, diamond, dash, korean, roman, alpha
    style: str = "default"

    # 자동 레벨 감지 여부
    auto_detect: bool = True

    # 미리 정의된 스타일들 (YAML에서 로드)
    # 레벨별 (기호, 들여쓰기) 튜플
    styles: Optional[Dict[str, Dict[int, Tuple[str, str]]]] = None

    def get_bullets(self) -> Dict[int, Tuple[str, str]]:
        """현재 스타일의 글머리 기호 반환"""
        if self.styles and self.style in self.styles:
            return self.styles[self.style]
        # 기본값
        return {
            0: ("□ ", " "),
            1: ("○", "   "),
            2: ("- ", "    "),
        }


@dataclass
class CaptionFormatPreset:
    """캡션 형식 프리셋"""
    pattern: str = "{type} {number}{separator}{title}"
    include_type: bool = True
    include_number: bool = True


@dataclass
class CaptionConfig:
    """캡션 설정 (CaptionFormatter 옵션)"""
    # 캡션 형식: title_only, type_title, type_num_title, standard, parenthesis, bracket_num
    format: str = "standard"

    # 캡션 위치: TOP, BOTTOM, LEFT, RIGHT
    position: str = "TOP"

    # 번호 재정렬 여부
    renumber: bool = True

    # 유형별 따로 번호 매김 (그림 1,2... 표 1,2...)
    renumber_by_type: bool = True

    # 시작 번호
    start_number: int = 1

    # 구분자 (번호와 제목 사이): ". ", ": ", " - " 등
    separator: str = ". "

    # autoNum 유지 여부
    keep_auto_num: bool = False

    # 캡션 형식 프리셋들 (YAML에서 로드)
    formats: Optional[Dict[str, CaptionFormatPreset]] = None

    def get_format_preset(self) -> CaptionFormatPreset:
        """현재 형식의 프리셋 반환"""
        if self.formats and self.format in self.formats:
            return self.formats[self.format]
        # 기본 프리셋
        default_formats = {
            "title_only": CaptionFormatPreset("[{title}]", False, False),
            "type_title": CaptionFormatPreset("[{type} {title}]", True, False),
            "type_num_title": CaptionFormatPreset("[{type} {number} {title}]", True, True),
            "standard": CaptionFormatPreset("{type} {number}{separator}{title}", True, True),
            "parenthesis": CaptionFormatPreset("{type} ({number}) {title}", True, True),
            "bracket_num": CaptionFormatPreset("[{type} {number}] {title}", True, True),
        }
        return default_formats.get(self.format, default_formats["standard"])


@dataclass
class TableCaptionConfig(CaptionConfig):
    """테이블 캡션 설정"""
    position: str = "TOP"  # 테이블은 기본 위로
    type_prefix: str = "표"  # 유형 접두어


@dataclass
class ImageCaptionConfig(CaptionConfig):
    """이미지 캡션 설정"""
    position: str = "BOTTOM"  # 이미지는 기본 아래로
    type_prefix: str = "그림"  # 유형 접두어


@dataclass
class FormatterConfig:
    """통합 포맷터 설정"""
    bullet: BulletConfig = field(default_factory=BulletConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    table_caption: TableCaptionConfig = field(default_factory=TableCaptionConfig)
    image_caption: ImageCaptionConfig = field(default_factory=ImageCaptionConfig)


class ConfigLoader:
    """YAML 설정 로더"""

    DEFAULT_CONFIG_NAME = "formatter_config.yaml"

    def __init__(self, config_path: Optional[str] = None):
        self.config_path = Path(config_path) if config_path else None
        self._config: Optional[FormatterConfig] = None

    def load(self, config_path: Optional[str] = None) -> FormatterConfig:
        """YAML 설정 파일 로드"""
        path = Path(config_path) if config_path else self.config_path

        if path and path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f) or {}
            self._config = self._parse_config(data)
        else:
            self._config = FormatterConfig()

        return self._config

    def load_from_string(self, yaml_string: str) -> FormatterConfig:
        """YAML 문자열에서 설정 로드"""
        data = yaml.safe_load(yaml_string) or {}
        self._config = self._parse_config(data)
        return self._config

    def load_from_dict(self, data: Dict[str, Any]) -> FormatterConfig:
        """딕셔너리에서 설정 로드"""
        self._config = self._parse_config(data)
        return self._config

    def _parse_config(self, data: Dict[str, Any]) -> FormatterConfig:
        """설정 데이터 파싱"""
        return FormatterConfig(
            bullet=self._parse_bullet_config(data.get('bullet', {})),
            caption=self._parse_caption_config(data.get('caption', {})),
            table_caption=self._parse_table_caption_config(data.get('table_caption', {})),
            image_caption=self._parse_image_caption_config(data.get('image_caption', {}))
        )

    def _parse_bullet_config(self, data: Dict[str, Any]) -> BulletConfig:
        """글머리 설정 파싱"""
        config = BulletConfig()

        if 'style' in data:
            config.style = data['style']
        if 'auto_detect' in data:
            config.auto_detect = bool(data['auto_detect'])

        # 미리 정의된 스타일들 파싱
        if 'styles' in data and isinstance(data['styles'], dict):
            config.styles = {}
            for style_name, style_data in data['styles'].items():
                if isinstance(style_data, dict):
                    config.styles[style_name] = {}
                    for level, bullet_data in style_data.items():
                        level_int = int(level)
                        if isinstance(bullet_data, dict):
                            symbol = bullet_data.get('symbol', '-')
                            indent = bullet_data.get('indent', '')
                            config.styles[style_name][level_int] = (symbol, indent)
                        elif isinstance(bullet_data, list) and len(bullet_data) >= 2:
                            config.styles[style_name][level_int] = (bullet_data[0], bullet_data[1])

        return config

    def _parse_caption_config(self, data: Dict[str, Any], default_position: str = "TOP") -> CaptionConfig:
        """캡션 설정 파싱"""
        config = CaptionConfig()

        if 'format' in data:
            config.format = data['format']
        if 'position' in data:
            config.position = data['position'].upper()
        else:
            config.position = default_position
        if 'renumber' in data:
            config.renumber = bool(data['renumber'])
        if 'renumber_by_type' in data:
            config.renumber_by_type = bool(data['renumber_by_type'])
        if 'start_number' in data:
            config.start_number = int(data['start_number'])
        if 'separator' in data:
            config.separator = data['separator']
        if 'keep_auto_num' in data:
            config.keep_auto_num = bool(data['keep_auto_num'])

        # 캡션 형식 프리셋 파싱
        if 'formats' in data and isinstance(data['formats'], dict):
            config.formats = {}
            for format_name, format_data in data['formats'].items():
                if isinstance(format_data, dict):
                    preset = CaptionFormatPreset(
                        pattern=format_data.get('pattern', ''),
                        include_type=bool(format_data.get('include_type', True)),
                        include_number=bool(format_data.get('include_number', True))
                    )
                    config.formats[format_name] = preset

        return config

    def _parse_table_caption_config(self, data: Dict[str, Any]) -> TableCaptionConfig:
        """테이블 캡션 설정 파싱"""
        base = self._parse_caption_config(data, default_position="TOP")
        type_prefix = data.get('type_prefix', '표')
        return TableCaptionConfig(
            format=base.format,
            position=base.position,
            renumber=base.renumber,
            renumber_by_type=base.renumber_by_type,
            start_number=base.start_number,
            separator=base.separator,
            keep_auto_num=base.keep_auto_num,
            formats=base.formats,
            type_prefix=type_prefix
        )

    def _parse_image_caption_config(self, data: Dict[str, Any]) -> ImageCaptionConfig:
        """이미지 캡션 설정 파싱"""
        base = self._parse_caption_config(data, default_position="BOTTOM")
        type_prefix = data.get('type_prefix', '그림')
        return ImageCaptionConfig(
            format=base.format,
            position=base.position,
            renumber=base.renumber,
            renumber_by_type=base.renumber_by_type,
            start_number=base.start_number,
            separator=base.separator,
            keep_auto_num=base.keep_auto_num,
            formats=base.formats,
            type_prefix=type_prefix
        )

    def save(self, config: FormatterConfig, path: Optional[str] = None) -> str:
        """설정을 YAML 파일로 저장"""
        save_path = Path(path) if path else self.config_path
        if not save_path:
            save_path = Path(self.DEFAULT_CONFIG_NAME)

        data = self._config_to_dict(config)

        with open(save_path, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

        return str(save_path)

    def _config_to_dict(self, config: FormatterConfig) -> Dict[str, Any]:
        """FormatterConfig를 딕셔너리로 변환"""
        return {
            'bullet': self._bullet_config_to_dict(config.bullet),
            'caption': self._caption_config_to_dict(config.caption),
            'table_caption': self._caption_config_to_dict(config.table_caption),
            'image_caption': self._caption_config_to_dict(config.image_caption)
        }

    def _bullet_config_to_dict(self, config: BulletConfig) -> Dict[str, Any]:
        """BulletConfig를 딕셔너리로 변환"""
        result = {
            'style': config.style,
            'auto_detect': config.auto_detect,
        }
        if config.styles:
            result['styles'] = {
                style_name: {
                    level: {'symbol': symbol, 'indent': indent}
                    for level, (symbol, indent) in bullets.items()
                }
                for style_name, bullets in config.styles.items()
            }
        return result

    def _caption_config_to_dict(self, config: CaptionConfig) -> Dict[str, Any]:
        """CaptionConfig를 딕셔너리로 변환"""
        result = {
            'format': config.format,
            'position': config.position,
            'renumber': config.renumber,
            'renumber_by_type': config.renumber_by_type,
            'start_number': config.start_number,
            'separator': config.separator,
            'keep_auto_num': config.keep_auto_num,
        }
        if config.formats:
            result['formats'] = {
                name: {
                    'pattern': preset.pattern,
                    'include_type': preset.include_type,
                    'include_number': preset.include_number
                }
                for name, preset in config.formats.items()
            }
        # type_prefix 추가 (TableCaptionConfig, ImageCaptionConfig)
        if hasattr(config, 'type_prefix'):
            result['type_prefix'] = config.type_prefix
        return result

    def to_yaml_string(self, config: FormatterConfig) -> str:
        """설정을 YAML 문자열로 변환"""
        data = self._config_to_dict(config)
        return yaml.dump(data, default_flow_style=False, allow_unicode=True, sort_keys=False)

    @property
    def config(self) -> FormatterConfig:
        """현재 로드된 설정 반환"""
        if self._config is None:
            self._config = self.load()
        return self._config


def load_config(config_path: Optional[str] = None) -> FormatterConfig:
    """YAML 설정 파일 로드 (편의 함수)"""
    loader = ConfigLoader()
    return loader.load(config_path)


def create_default_config() -> FormatterConfig:
    """기본 설정 생성"""
    return FormatterConfig()


def save_default_config(path: str) -> str:
    """기본 설정을 YAML 파일로 저장"""
    loader = ConfigLoader()
    config = create_default_config()
    return loader.save(config, path)
