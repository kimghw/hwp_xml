# -*- coding: utf-8 -*-
"""
프로젝트 설정 및 경로 관리

환경변수 또는 기본값을 통해 경로를 설정합니다.
"""

import os
import platform
from pathlib import Path


# ============================================================
# 기본 경로 설정
# ============================================================

# 프로젝트 루트 디렉토리
PROJECT_ROOT = Path(__file__).parent.resolve()

# 모듈 디렉토리
HWPXML_MODULE_DIR = PROJECT_ROOT / 'hwpxml'
WIN32_MODULE_DIR = PROJECT_ROOT / 'win32'
EXCEL_MODULE_DIR = PROJECT_ROOT / 'excel'
CORE_MODULE_DIR = PROJECT_ROOT / 'core'
TESTS_DIR = PROJECT_ROOT / 'tests'

# 외부 의존성 경로 (환경변수로 설정 가능)
WIN32HWP_DIR = Path(os.environ.get('WIN32HWP_DIR', r'C:\win32hwp'))


# ============================================================
# 플랫폼별 경로 변환
# ============================================================

def get_windows_path(path: Path) -> str:
    """WSL 경로를 Windows 경로로 변환"""
    path_str = str(path)
    if path_str.startswith('/mnt/'):
        # /mnt/c/... -> C:\...
        parts = path_str.split('/')
        drive = parts[2].upper()
        rest = '\\'.join(parts[3:])
        return f"{drive}:\\{rest}"
    return path_str


def get_wsl_path(path: str) -> Path:
    """Windows 경로를 WSL 경로로 변환"""
    if len(path) >= 2 and path[1] == ':':
        drive = path[0].lower()
        rest = path[2:].replace('\\', '/')
        return Path(f"/mnt/{drive}{rest}")
    return Path(path)


def is_windows() -> bool:
    """Windows 환경인지 확인"""
    return platform.system() == "Windows"


def is_wsl() -> bool:
    """WSL 환경인지 확인"""
    return 'microsoft' in platform.uname().release.lower()


# ============================================================
# 테스트 데이터 경로
# ============================================================

def get_test_hwpx_path() -> Path:
    """테스트용 HWPX 파일 경로"""
    return PROJECT_ROOT / 'table.hwpx'


def get_test_output_path(filename: str = 'output.xlsx') -> Path:
    """테스트 출력 파일 경로"""
    return EXCEL_MODULE_DIR / filename


# ============================================================
# 모듈 경로 설정 (import용)
# ============================================================

def setup_module_paths():
    """모듈 경로를 sys.path에 추가"""
    import sys

    paths_to_add = [
        str(PROJECT_ROOT),
        str(HWPXML_MODULE_DIR),
        str(CORE_MODULE_DIR),
    ]

    # Windows 환경에서는 win32hwp 경로도 추가
    if is_windows() and WIN32HWP_DIR.exists():
        paths_to_add.append(str(WIN32HWP_DIR))

    for path in paths_to_add:
        if path not in sys.path:
            sys.path.insert(0, path)


# ============================================================
# 단위 변환 상수 (HWPUNIT 기준)
# ============================================================

class Units:
    """단위 변환 상수"""
    HWPUNIT_PER_PT = 100
    HWPUNIT_PER_INCH = 7200
    HWPUNIT_PER_CM = 7200 / 2.54
    HWPUNIT_PER_MM = 7200 / 25.4

    # Excel 변환용
    HWPUNIT_TO_EXCEL_WIDTH = 632
    HWPUNIT_TO_EXCEL_PT = 100


# ============================================================
# 로깅 설정
# ============================================================

import logging

def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """로깅 설정"""
    logger = logging.getLogger('hwp_xml')

    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    logger.setLevel(level)
    return logger


# 모듈 로드 시 자동 경로 설정
setup_module_paths()
