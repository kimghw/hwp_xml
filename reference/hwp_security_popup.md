# 한글 COM API 보안 팝업 해결 방법

---

## 문제 상황

한글 COM API를 사용하여 파일을 열거나 저장할 때 다음과 같은 보안 팝업이 나타날 수 있습니다:

- "한글을 이용하여 위 파일에 접근하려는 시도 (파일의 손상 또는 유출의 위험 등)가 있습니다."
- "접근 허용" 버튼을 클릭해야 진행됨

이 팝업은 자동화 스크립트 실행을 방해합니다.

---

## 해결 방법

### 1. 보안 모듈 등록 (필수)

**두 가지 모듈을 모두 등록해야 합니다:**

```python
# 1. SetMessageBoxMode 먼저 설정 (순서 중요!)
hwp.SetMessageBoxMode(0x7FFFFFFF)

# 2. 보안 모듈 등록 (두 가지 모두 필요)
hwp.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
```

| 모듈 | 설명 |
|------|------|
| `FilePathCheckerModule` | 파일 경로 접근 허용 |
| `SecurityModule` | 모든 보안 경고 자동 허용 (스크립트, 매크로, 개인정보 등) |

### 2. 파일 열기 방식 (HAction 사용)

**`hwp.Open()` 대신 `HAction.Execute("FileOpen")`을 사용합니다:**

```python
# 팝업이 나타날 수 있음 (비권장)
hwp.Open(filepath, "HWP", "forceopen:true")

# 팝업 없이 파일 열기 (권장)
hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = filepath
hwp.HParameterSet.HFileOpenSave.Format = "HWP"  # 또는 "HWPX"
hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
```

### 3. 파일 저장 방식 (HAction 사용)

```python
# 팝업 없이 파일 저장 (권장)
hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = filepath
hwp.HParameterSet.HFileOpenSave.Format = "HWP"  # 또는 "HWPX"
hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
```

---

## 전체 코드 예시

### hwp_file_manager.py

```python
# -*- coding: utf-8 -*-
"""한글 COM API 공통 유틸리티"""

from typing import Optional


def get_hwp_instance():
    """열린 한글 인스턴스 가져오기"""
    try:
        import win32com.client as win32
        hwp = win32.GetActiveObject("HWPFrame.HwpObject")
        # 메시지박스 모드 먼저 설정 (순서 중요!)
        hwp.SetMessageBoxMode(0x7FFFFFFF)
        # 보안 모듈 등록 (두 가지 모두 필요)
        hwp.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        return hwp
    except:
        return None


def create_hwp_instance(visible: bool = True):
    """새 한글 인스턴스 생성"""
    import win32com.client as win32
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    # 메시지박스 모드 먼저 설정 (순서 중요!)
    hwp.SetMessageBoxMode(0x7FFFFFFF)
    # 보안 모듈 등록 (두 가지 모두 필요)
    hwp.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    if visible:
        hwp.XHwpWindows.Item(0).Visible = True
    else:
        hwp.XHwpWindows.Item(0).Visible = False
    return hwp


def open_hwp_file(hwp, filepath: str, format: str = "HWP"):
    """파일 열기 (팝업 없이)"""
    hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = filepath
    hwp.HParameterSet.HFileOpenSave.Format = format
    hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)


def save_hwp_file(hwp, filepath: str, format: str = "HWP"):
    """파일 저장 (팝업 없이)"""
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = filepath
    hwp.HParameterSet.HFileOpenSave.Format = format
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
```

### 사용 예시

```python
from hwp_file_manager import create_hwp_instance, open_hwp_file, save_hwp_file

# 한글 인스턴스 생성 (숨김 모드)
hwp = create_hwp_instance(visible=False)

# 파일 열기 (팝업 없이)
open_hwp_file(hwp, r"C:\path\to\input.hwp", "HWP")

# 작업 수행...

# 파일 저장 (팝업 없이)
save_hwp_file(hwp, r"C:\path\to\output.hwp", "HWP")

# 종료
hwp.Quit()
```

---

## 핵심 포인트 정리

| 항목 | 내용 |
|------|------|
| **SetMessageBoxMode 순서** | `RegisterModule` 전에 먼저 호출해야 함 |
| **SetMessageBoxMode 값** | `0x7FFFFFFF` (최대 signed 32-bit) |
| **보안 모듈** | `FilePathCheckerModule` + `SecurityModule` 둘 다 등록 |
| **파일 열기** | `hwp.Open()` 대신 `HAction.Execute("FileOpen")` 사용 |
| **파일 저장** | `HAction.Execute("FileSaveAs_S")` 사용 |
| **Python 캐시** | 코드 수정 후 `__pycache__` 삭제 필요할 수 있음 |

---

## 주의사항

### Python 캐시 문제

코드 수정 후에도 팝업이 계속 나타나면 Python 캐시를 삭제합니다:

```bash
# Windows CMD
del /s /q __pycache__\*.pyc
rmdir /s /q __pycache__
```

### 한글 프로그램 보안 설정

한글 프로그램 자체의 보안 설정도 확인합니다:
1. 한글 실행
2. **도구 → 환경 설정 → 기타 → 개인 정보 보호** (또는 **보안**)
3. 보안 수준을 **"낮음"**으로 설정

### SetMessageBoxMode 플래그 값

| 값 | 설명 |
|----|------|
| `0x00000001` | MB_OK 자동 처리 |
| `0x00010000` | 보안 경고 자동 허용 |
| `0x00020000` | 스크립트 경고 자동 허용 |
| `0x00040000` | 개인정보 경고 자동 허용 |
| `0x00080000` | 기타 경고 자동 허용 |
| `0x00100000` | 파일 손상/유출 위험 경고 |
| `0x7FFFFFFF` | 모든 플래그 (권장) |

> **주의**: `0xFFFFFFFF`는 Python에서 signed 32-bit 오버플로우 오류 발생

---

## 테스트 방법

```bash
# WSL에서 실행
cmd.exe /c "cd /d C:\hwp_xml\win32 && python test_popup.py" 2>&1
```

### test_popup.py

```python
# -*- coding: utf-8 -*-
import win32com.client as win32

print("한글 인스턴스 생성...")
hwp = win32.Dispatch("HWPFrame.HwpObject")

print("SetMessageBoxMode 설정...")
hwp.SetMessageBoxMode(0x7FFFFFFF)

print("RegisterModule 설정...")
hwp.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

print("창 숨김...")
hwp.XHwpWindows.Item(0).Visible = False

print("파일 열기 (HAction 방식)...")
hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = r"C:\hwp_xml\test.hwp"
hwp.HParameterSet.HFileOpenSave.Format = "HWP"
hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

print("파일 열기 완료")

print("저장 중...")
hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = r"C:\hwp_xml\test_output.hwp"
hwp.HParameterSet.HFileOpenSave.Format = "HWP"
hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

print("저장 완료")
hwp.Quit()
print("종료")
```

---

*작성일: 2026-01-28*
