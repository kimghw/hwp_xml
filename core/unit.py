# -*- coding: utf-8 -*-
"""
HWPUNIT 단위 변환 유틸리티

HWPUNIT은 한글 문서에서 사용하는 내부 단위입니다.
- 1 inch = 7200 HWPUNIT
- 1 pt = 100 HWPUNIT
- 1 cm ≈ 2834.6 HWPUNIT
- 1 mm ≈ 283.46 HWPUNIT
"""


class Unit:
    """단위 변환 상수 및 메서드"""

    # 기본 변환 상수
    HWPUNIT_PER_PT = 100
    HWPUNIT_PER_INCH = 7200
    HWPUNIT_PER_CM = 7200 / 2.54  # ≈ 2834.6
    HWPUNIT_PER_MM = 7200 / 25.4  # ≈ 283.46

    # 포인트 변환
    PT_PER_INCH = 72
    PT_PER_CM = 72 / 2.54  # ≈ 28.35

    # Excel 변환용
    EXCEL_CHAR_TO_PT = 7
    HWPUNIT_TO_EXCEL_WIDTH = 632  # HWPUNIT → Excel 열 너비
    HWPUNIT_TO_EXCEL_PT = 100     # HWPUNIT → Excel 행 높이 (pt)

    # ========================================
    # HWPUNIT ↔ 포인트
    # ========================================

    @staticmethod
    def hwpunit_to_pt(hwpunit: int) -> float:
        """HWPUNIT -> 포인트"""
        return hwpunit / Unit.HWPUNIT_PER_PT

    @staticmethod
    def pt_to_hwpunit(pt: float) -> int:
        """포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    # ========================================
    # HWPUNIT ↔ cm
    # ========================================

    @staticmethod
    def hwpunit_to_cm(hwpunit: int) -> float:
        """HWPUNIT -> cm"""
        return hwpunit / Unit.HWPUNIT_PER_CM

    @staticmethod
    def cm_to_hwpunit(cm: float) -> int:
        """cm -> HWPUNIT"""
        return int(cm * Unit.HWPUNIT_PER_CM)

    # ========================================
    # HWPUNIT ↔ mm
    # ========================================

    @staticmethod
    def hwpunit_to_mm(hwpunit: int) -> float:
        """HWPUNIT -> mm"""
        return hwpunit / Unit.HWPUNIT_PER_MM

    @staticmethod
    def mm_to_hwpunit(mm: float) -> int:
        """mm -> HWPUNIT"""
        return int(mm * Unit.HWPUNIT_PER_MM)

    # ========================================
    # HWPUNIT ↔ inch
    # ========================================

    @staticmethod
    def hwpunit_to_inch(hwpunit: int) -> float:
        """HWPUNIT -> inch"""
        return hwpunit / Unit.HWPUNIT_PER_INCH

    @staticmethod
    def inch_to_hwpunit(inch: float) -> int:
        """inch -> HWPUNIT"""
        return int(inch * Unit.HWPUNIT_PER_INCH)

    # ========================================
    # Excel 변환
    # ========================================

    @staticmethod
    def hwpunit_to_excel_width(hwpunit: int) -> float:
        """HWPUNIT -> Excel 열 너비 (문자 단위)"""
        return hwpunit / Unit.HWPUNIT_TO_EXCEL_WIDTH

    @staticmethod
    def hwpunit_to_excel_height(hwpunit: int) -> float:
        """HWPUNIT -> Excel 행 높이 (pt)"""
        return hwpunit / Unit.HWPUNIT_TO_EXCEL_PT

    @staticmethod
    def excel_width_to_hwpunit(width: float) -> int:
        """Excel 열 너비 -> HWPUNIT"""
        return int(width * Unit.HWPUNIT_TO_EXCEL_WIDTH)

    @staticmethod
    def excel_pt_to_hwpunit(pt: float) -> int:
        """Excel 포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    @staticmethod
    def excel_char_to_hwpunit(chars: float) -> int:
        """Excel 문자 단위(열 너비) -> HWPUNIT"""
        pt = chars * Unit.EXCEL_CHAR_TO_PT
        return int(pt * Unit.HWPUNIT_PER_PT)
