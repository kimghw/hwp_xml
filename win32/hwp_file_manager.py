# -*- coding: utf-8 -*-
"""
한글 COM API 공통 유틸리티

HWP 인스턴스 연결, 파일 대화상자 등 공통 기능
"""

from typing import Optional


def get_hwp_instance():
    """
    열린 한글 인스턴스 가져오기

    Returns:
        hwp: 한글 COM 객체 (없으면 None)
    """
    try:
        import win32com.client as win32
        hwp = win32.GetActiveObject("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        return hwp
    except:
        return None


def create_hwp_instance(visible: bool = True):
    """
    새 한글 인스턴스 생성

    Args:
        visible: 창 표시 여부

    Returns:
        hwp: 한글 COM 객체
    """
    import win32com.client as win32
    # DispatchEx로 항상 새 프로세스 생성 (기존 인스턴스와 분리)
    hwp = win32.DispatchEx("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
    return hwp


def get_or_create_hwp(visible: bool = True):
    """
    열린 한글 인스턴스 가져오기, 없으면 새로 생성

    Args:
        visible: 새로 생성 시 창 표시 여부

    Returns:
        tuple: (hwp, is_new) - hwp 객체와 새로 생성 여부
    """
    hwp = get_hwp_instance()
    if hwp:
        return hwp, False
    return create_hwp_instance(visible), True


def get_active_filepath(hwp) -> Optional[str]:
    """
    열린 문서의 파일 경로 가져오기

    Args:
        hwp: 한글 COM 객체

    Returns:
        파일 경로 (없으면 None)
    """
    try:
        path = hwp.XHwpDocuments.Active_XHwpDocument.Path
        return path if path else None
    except:
        return None


def open_file_dialog(
    title: str = "한글 파일 선택",
    filter_str: str = "한글 파일 (*.hwp;*.hwpx)\0*.hwp;*.hwpx\0모든 파일 (*.*)\0*.*\0\0"
) -> Optional[str]:
    """
    Windows 파일 선택 대화상자

    Args:
        title: 대화상자 제목
        filter_str: 파일 필터 (Win32 형식)

    Returns:
        선택된 파일 경로 (취소 시 None)
    """
    try:
        import win32gui
        import win32con

        filename, _, _ = win32gui.GetOpenFileNameW(
            Filter=filter_str,
            Title=title,
            Flags=win32con.OFN_FILEMUSTEXIST
        )
        return filename
    except Exception as e:
        # 취소 또는 오류
        return None


def save_hwp(hwp, filepath: str, format: str = "HWP") -> bool:
    """
    한글 문서 저장 (편집 가능하게)

    Args:
        hwp: 한글 COM 객체
        filepath: 저장 경로
        format: 저장 형식 ("HWP" 또는 "HWPX")

    Returns:
        성공 여부
    """
    try:
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = filepath
        hwp.HParameterSet.HFileOpenSave.Format = format
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        return True
    except:
        return False


def close_document(hwp, save: bool = False) -> bool:
    """
    현재 문서 닫기

    Args:
        hwp: 한글 COM 객체
        save: 저장 여부 (True=저장, False=저장 안 함)

    Returns:
        성공 여부
    """
    try:
        hwp.Clear(0 if save else 1)
        return True
    except:
        return False
