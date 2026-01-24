# -*- coding: utf-8 -*-
"""
파일 탐색기 대화상자 모듈

WSL 환경에서 Windows 파일 탐색기를 열어 파일을 선택합니다.
"""

import subprocess
import os
from pathlib import Path
from typing import Optional, List, Tuple


def open_file_dialog(
    title: str = "파일 선택",
    filetypes: Optional[List[Tuple[str, str]]] = None,
    initialdir: Optional[str] = None
) -> Optional[str]:
    """
    Windows 파일 탐색기를 열어 파일 선택

    Args:
        title: 대화상자 제목
        filetypes: 파일 유형 필터 리스트 [("설명", "*.확장자"), ...]
        initialdir: 초기 디렉토리 (Windows 경로)

    Returns:
        선택된 파일의 Windows 경로 (취소 시 None)

    Example:
        >>> path = open_file_dialog(
        ...     title="HWPX 파일 선택",
        ...     filetypes=[("HWPX 파일", "*.hwpx"), ("HWP 파일", "*.hwp")],
        ...     initialdir="C:\\Documents"
        ... )
    """
    if filetypes is None:
        filetypes = [("모든 파일", "*.*")]

    # filetypes를 PowerShell 필터 문자열로 변환
    # 예: "HWPX 파일 (*.hwpx)|*.hwpx|HWP 파일 (*.hwp)|*.hwp"
    filter_parts = []
    for desc, pattern in filetypes:
        filter_parts.append(f"{desc} ({pattern})|{pattern}")
    filter_str = "|".join(filter_parts)

    # 초기 디렉토리 설정
    initial_dir_cmd = ""
    if initialdir:
        initial_dir_cmd = f'$dialog.InitialDirectory = "{initialdir}";'

    # PowerShell 스크립트 (UTF-8 인코딩 강제)
    ps_script = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "{title}"
$dialog.Filter = "{filter_str}"
{initial_dir_cmd}
$dialog.ShowDialog() | Out-Null
$dialog.FileName
'''

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            timeout=300
        )
        # PowerShell 출력 인코딩 처리
        try:
            filepath = result.stdout.decode('utf-8').strip()
        except UnicodeDecodeError:
            try:
                filepath = result.stdout.decode('cp949').strip()
            except UnicodeDecodeError:
                filepath = result.stdout.decode('utf-16-le', errors='ignore').strip()

        # 빈 문자열이 아니면 반환 (한글 경로는 os.path.exists가 실패할 수 있음)
        if filepath:
            return filepath
        return None
    except subprocess.TimeoutExpired:
        return None
    except Exception as e:
        print(f"파일 대화상자 오류: {e}")
        return None


def open_hwp_dialog(initialdir: Optional[str] = None) -> Optional[str]:
    """HWP/HWPX 파일 선택 대화상자"""
    return open_file_dialog(
        title="한글 파일 선택",
        filetypes=[
            ("한글 파일", "*.hwp;*.hwpx"),
            ("HWPX 파일", "*.hwpx"),
            ("HWP 파일", "*.hwp"),
            ("모든 파일", "*.*")
        ],
        initialdir=initialdir
    )


def open_hwpx_dialog(initialdir: Optional[str] = None) -> Optional[str]:
    """HWPX 파일 선택 대화상자"""
    return open_file_dialog(
        title="HWPX 파일 선택",
        filetypes=[
            ("HWPX 파일", "*.hwpx"),
            ("모든 파일", "*.*")
        ],
        initialdir=initialdir
    )


def open_excel_dialog(initialdir: Optional[str] = None) -> Optional[str]:
    """Excel 파일 선택 대화상자"""
    return open_file_dialog(
        title="Excel 파일 선택",
        filetypes=[
            ("Excel 파일", "*.xlsx;*.xls"),
            ("모든 파일", "*.*")
        ],
        initialdir=initialdir
    )


def save_file_dialog(
    title: str = "파일 저장",
    filetypes: Optional[List[Tuple[str, str]]] = None,
    initialdir: Optional[str] = None,
    defaultext: Optional[str] = None
) -> Optional[str]:
    """
    Windows 파일 저장 대화상자

    Args:
        title: 대화상자 제목
        filetypes: 파일 유형 필터
        initialdir: 초기 디렉토리
        defaultext: 기본 확장자

    Returns:
        저장할 파일의 Windows 경로 (취소 시 None)
    """
    if filetypes is None:
        filetypes = [("모든 파일", "*.*")]

    filter_parts = []
    for desc, pattern in filetypes:
        filter_parts.append(f"{desc} ({pattern})|{pattern}")
    filter_str = "|".join(filter_parts)

    initial_dir_cmd = ""
    if initialdir:
        initial_dir_cmd = f'$dialog.InitialDirectory = "{initialdir}";'

    default_ext_cmd = ""
    if defaultext:
        default_ext_cmd = f'$dialog.DefaultExt = "{defaultext}";'

    ps_script = f'''
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.SaveFileDialog
$dialog.Title = "{title}"
$dialog.Filter = "{filter_str}"
{initial_dir_cmd}
{default_ext_cmd}
$dialog.ShowDialog() | Out-Null
$dialog.FileName
'''

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            timeout=300
        )
        try:
            filepath = result.stdout.decode('utf-8').strip()
        except UnicodeDecodeError:
            filepath = result.stdout.decode('cp949', errors='ignore').strip()

        return filepath if filepath else None
    except subprocess.TimeoutExpired:
        return None
    except Exception as e:
        print(f"파일 대화상자 오류: {e}")
        return None


def wsl_to_windows_path(wsl_path: str) -> str:
    """WSL 경로를 Windows 경로로 변환"""
    if wsl_path.startswith("/mnt/"):
        # /mnt/c/path -> C:\path
        parts = wsl_path[5:].split("/", 1)
        drive = parts[0].upper()
        rest = parts[1].replace("/", "\\") if len(parts) > 1 else ""
        return f"{drive}:\\{rest}"
    return wsl_path


def windows_to_wsl_path(win_path: str) -> str:
    """Windows 경로를 WSL 경로로 변환"""
    if len(win_path) >= 2 and win_path[1] == ":":
        # C:\path -> /mnt/c/path
        drive = win_path[0].lower()
        rest = win_path[2:].replace("\\", "/")
        return f"/mnt/{drive}{rest}"
    return win_path


if __name__ == "__main__":
    print("파일 탐색기 테스트")
    print("=" * 40)

    filepath = open_hwp_dialog()
    if filepath:
        print(f"선택된 파일: {filepath}")
        print(f"WSL 경로: {windows_to_wsl_path(filepath)}")
    else:
        print("파일 선택 취소")
