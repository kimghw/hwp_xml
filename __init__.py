# -*- coding: utf-8 -*-
"""
hwp_xml 패키지

HWPX 파일 처리 및 Excel 변환 도구

모듈:
- xml: HWPX XML 파싱 (테이블, 페이지 속성 추출)
- excel: HWPX → Excel 변환
- win32: 한글 COM API 연동 (Windows 전용)
- core: 공통 유틸리티 (단위 변환 등)
"""

from .config import (
    PROJECT_ROOT,
    HWPXML_MODULE_DIR,
    WIN32_MODULE_DIR,
    EXCEL_MODULE_DIR,
    CORE_MODULE_DIR,
    is_windows,
    is_wsl,
    get_test_hwpx_path,
    setup_logging,
)

__version__ = '0.1.0'

__all__ = [
    'PROJECT_ROOT',
    'HWPXML_MODULE_DIR',
    'WIN32_MODULE_DIR',
    'EXCEL_MODULE_DIR',
    'CORE_MODULE_DIR',
    'is_windows',
    'is_wsl',
    'get_test_hwpx_path',
    'setup_logging',
]
