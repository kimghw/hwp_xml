# -*- coding: utf-8 -*-
"""
테이블 병합 시 add_ 필드용 포맷터 설정 로더

table_formatter_config.yaml 파일을 로드하여 필드별 포맷터 설정을 관리합니다.
"""

import re
import yaml
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List, Tuple
from pathlib import Path

from ..formatters.bullet_formatter import BulletFormatter


@dataclass
class FieldFormatterConfig:
    """필드별 포맷터 설정"""
    # 필드명 또는 패턴
    name: Optional[str] = None
    pattern: Optional[str] = None

    # 포맷터 유형: none, bullet, caption
    formatter: str = "none"

    # 구분자
    separator: str = " "

    # 포맷터 옵션
    options: Dict[str, Any] = field(default_factory=dict)


@dataclass
class TableFormatterConfig:
    """테이블 포맷터 통합 설정"""
    # 기본 설정
    default_formatter: str = "none"
    default_separator: str = " "

    # 필드별 설정
    field_configs: List[FieldFormatterConfig] = field(default_factory=list)

    # 글머리 기호 설정
    bullet_style: str = "default"
    bullet_auto_detect: bool = True
    bullet_styles: Optional[Dict[str, Dict[int, Tuple[str, str]]]] = None


class TableFormatterConfigLoader:
    """테이블 포맷터 설정 로더"""

    DEFAULT_CONFIG_PATH = Path(__file__).parent.parent / "formatters" / "table_formatter_config.yaml"

    def __init__(self, config_path: Optional[str] = None):
        self.config_path = Path(config_path) if config_path else self.DEFAULT_CONFIG_PATH
        self._config: Optional[TableFormatterConfig] = None
        self._formatters: Dict[str, BulletFormatter] = {}

    def load(self, config_path: Optional[str] = None) -> TableFormatterConfig:
        """YAML 설정 파일 로드"""
        path = Path(config_path) if config_path else self.config_path

        if path and path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f) or {}
            self._config = self._parse_config(data)
        else:
            self._config = TableFormatterConfig()

        return self._config

    def _parse_config(self, data: Dict[str, Any]) -> TableFormatterConfig:
        """설정 데이터 파싱"""
        config = TableFormatterConfig()

        # 기본 설정
        default_data = data.get('default', {})
        config.default_formatter = default_data.get('formatter', 'none')
        config.default_separator = default_data.get('separator', ' ')

        # 필드별 설정
        fields_data = data.get('fields', [])
        for field_data in fields_data:
            if isinstance(field_data, dict):
                field_config = FieldFormatterConfig(
                    name=field_data.get('name'),
                    pattern=field_data.get('pattern'),
                    formatter=field_data.get('formatter', config.default_formatter),
                    separator=field_data.get('separator', config.default_separator),
                    options=field_data.get('options', {})
                )
                config.field_configs.append(field_config)

        # 글머리 기호 설정
        bullet_data = data.get('bullet', {})
        config.bullet_style = bullet_data.get('style', 'default')
        config.bullet_auto_detect = bullet_data.get('auto_detect', True)

        # 글머리 스타일들 파싱
        if 'styles' in bullet_data and isinstance(bullet_data['styles'], dict):
            config.bullet_styles = {}
            for style_name, style_data in bullet_data['styles'].items():
                if isinstance(style_data, dict):
                    config.bullet_styles[style_name] = {}
                    for level, bullet_data_item in style_data.items():
                        level_int = int(level)
                        if isinstance(bullet_data_item, dict):
                            symbol = bullet_data_item.get('symbol', '-')
                            indent = bullet_data_item.get('indent', '')
                            config.bullet_styles[style_name][level_int] = (symbol, indent)

        return config

    def get_config_for_field(self, field_name: str) -> FieldFormatterConfig:
        """필드명에 해당하는 설정 반환"""
        if self._config is None:
            self.load()

        # 1. 정확한 필드명 매칭
        for field_config in self._config.field_configs:
            if field_config.name and field_config.name == field_name:
                return field_config

        # 2. 패턴 매칭
        for field_config in self._config.field_configs:
            if field_config.pattern:
                if re.match(field_config.pattern, field_name):
                    return field_config

        # 3. 기본값 반환
        return FieldFormatterConfig(
            formatter=self._config.default_formatter,
            separator=self._config.default_separator
        )

    def get_formatter(self, field_name: str) -> Optional[BulletFormatter]:
        """필드명에 해당하는 포맷터 반환"""
        field_config = self.get_config_for_field(field_name)

        if field_config.formatter == "none":
            return None

        if field_config.formatter == "bullet":
            style = field_config.options.get('style', self._config.bullet_style)

            # 캐시된 포맷터 사용
            if style not in self._formatters:
                self._formatters[style] = BulletFormatter(style=style)

            return self._formatters[style]

        return None

    def format_value(self, field_name: str, value: str) -> str:
        """필드값에 포맷터 적용"""
        formatter = self.get_formatter(field_name)

        if formatter is None:
            return value

        field_config = self.get_config_for_field(field_name)
        auto_detect = field_config.options.get('auto_detect', self._config.bullet_auto_detect)

        result = formatter.format_text(value, auto_detect=auto_detect)

        if result.success:
            return result.formatted_text

        return value

    @property
    def config(self) -> TableFormatterConfig:
        """현재 로드된 설정"""
        if self._config is None:
            self.load()
        return self._config


# 편의 함수
def load_table_formatter_config(config_path: Optional[str] = None) -> TableFormatterConfig:
    """테이블 포맷터 설정 로드"""
    loader = TableFormatterConfigLoader(config_path)
    return loader.load()


def format_add_field_value(field_name: str, value: str, config_path: Optional[str] = None) -> str:
    """add_ 필드값에 포맷터 적용"""
    loader = TableFormatterConfigLoader(config_path)
    loader.load()
    return loader.format_value(field_name, value)
