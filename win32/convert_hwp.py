# -*- coding: utf-8 -*-
"""
HWP ↔ HWPX 변환 모듈

한글 COM API를 사용하여 HWP 파일을 HWPX로 변환합니다.
Windows 환경에서만 실행 가능합니다.
"""

import os
import sys
import tempfile
from pathlib import Path

try:
    import win32com.client as win32
    from pywintypes import com_error
except ImportError:
    print("pywin32가 설치되지 않았습니다.")
    sys.exit(1)

try:
    from hwp_file_manager import get_or_create_hwp
except ImportError:
    from win32.hwp_file_manager import get_or_create_hwp


def convert_hwp_to_hwpx(hwp_path: str, output_path: str = None) -> str:
    """
    HWP 파일을 HWPX로 변환

    Args:
        hwp_path: 원본 HWP 파일 경로
        output_path: 출력 HWPX 파일 경로 (None이면 같은 위치에 확장자만 변경)

    Returns:
        변환된 HWPX 파일 경로
    """
    hwp_path = Path(hwp_path)
    if not hwp_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwp_path}")

    if output_path is None:
        output_path = hwp_path.with_suffix(".hwpx")
    else:
        output_path = Path(output_path)

    hwp, is_new = get_or_create_hwp(visible=False)

    try:
        # 파일 열기
        hwp.Open(str(hwp_path), "HWP", "forceopen:true")

        # HWPX로 저장
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(output_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        # 문서 닫기
        hwp.Clear(1)

        return str(output_path)

    except com_error as e:
        raise RuntimeError(f"변환 중 오류: {e}")
    finally:
        if is_new:
            hwp.Quit()


def convert_hwpx_to_hwp(hwpx_path: str, output_path: str = None) -> str:
    """
    HWPX 파일을 HWP로 변환

    Args:
        hwpx_path: 원본 HWPX 파일 경로
        output_path: 출력 HWP 파일 경로

    Returns:
        변환된 HWP 파일 경로
    """
    hwpx_path = Path(hwpx_path)
    if not hwpx_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {hwpx_path}")

    if output_path is None:
        output_path = hwpx_path.with_suffix(".hwp")
    else:
        output_path = Path(output_path)

    hwp, is_new = get_or_create_hwp(visible=False)

    try:
        hwp.Open(str(hwpx_path), "HWPX", "forceopen:true")

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(output_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.Clear(1)

        return str(output_path)

    except com_error as e:
        raise RuntimeError(f"변환 중 오류: {e}")
    finally:
        if is_new:
            hwp.Quit()


def convert_to_hwpx_temp(hwp_path: str) -> str:
    """
    HWP 파일을 임시 HWPX 파일로 변환

    Args:
        hwp_path: 원본 HWP 파일 경로

    Returns:
        임시 HWPX 파일 경로
    """
    hwp_path = Path(hwp_path)
    temp_dir = tempfile.gettempdir()
    output_path = Path(temp_dir) / f"{hwp_path.stem}_temp.hwpx"

    return convert_hwp_to_hwpx(str(hwp_path), str(output_path))


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="HWP ↔ HWPX 변환")
    parser.add_argument("input", help="입력 파일 경로")
    parser.add_argument("-o", "--output", help="출력 파일 경로")
    parser.add_argument("--to-hwp", action="store_true", help="HWPX → HWP 변환")

    args = parser.parse_args()

    try:
        if args.to_hwp:
            result = convert_hwpx_to_hwp(args.input, args.output)
        else:
            result = convert_hwp_to_hwpx(args.input, args.output)
        print(f"변환 완료: {result}")
    except Exception as e:
        print(f"오류: {e}")
        sys.exit(1)
