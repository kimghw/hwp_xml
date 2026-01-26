# -*- coding: utf-8 -*-
"""SecurityModule 등록 테스트 - SetMessageBoxMode 사용"""

import win32com.client as win32

print("=" * 50)
print("SecurityModule 등록 테스트")
print("=" * 50)

# 방법: SetMessageBoxMode로 대화상자 자동 처리
print("\n[방법] SetMessageBoxMode 사용")
try:
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

    # 메시지 박스 자동 처리 (예/확인 자동 클릭)
    hwp.SetMessageBoxMode(0x00010000)  # MB_YESNO_IDYES

    hwp.XHwpWindows.Item(0).Visible = True
    hwp.Open(r"C:\hwp_xml\test_step5.hwp", "HWP", "forceopen:true")
    print("  파일 열기 성공")
    hwp.Quit()
except Exception as e:
    print(f"  오류: {e}")

print("\n" + "=" * 50)
print("테스트 완료")
print("=" * 50)
