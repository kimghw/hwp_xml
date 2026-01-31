# Formatters 모듈

텍스트 양식 변환을 위한 정규식 기반 포맷터 모듈입니다.

## 구조

```
merge/formatters/          # 정규식 기반 (포맷 적용)
├── bullet_formatter.py    # 글머리 기호 변환
├── caption_formatter.py   # 캡션 변환
└── __init__.py

agent/                     # SDK 기반 (분석만)
├── bullet_formatter.py    # 글머리 레벨 분석
├── caption_formatter.py   # 캡션 제목 추출
└── sdk.py                 # Claude SDK 래퍼
```

## 역할 분리

| 구분 | SDK (agent/) | 정규식 (merge/formatters/) |
|------|-------------|---------------------------|
| **BulletFormatter** | `analyze_levels()` → `List[int]` | `format_text(levels=...)` |
| **CaptionFormatter** | `extract_title_with_sdk()` → 제목 | `format_caption()`, `to_bracket_format()` |

**핵심 원칙**: SDK는 분석만, 포맷 적용은 정규식으로

---

## BulletFormatter (글머리 기호)

### 기본 사용법

```python
from merge.formatters import BulletFormatter

formatter = BulletFormatter(style="default")

# 레벨 지정하여 변환
result = formatter.format_text("항목1\n항목2\n항목3", levels=[0, 1, 2])
print(result.formatted_text)
# □ 항목1
#    ○항목2
#     - 항목3

# 자동 레벨 감지
result = formatter.auto_format("1. 첫번째\n  a. 두번째\n    - 세번째")
```

### 글머리 스타일

```python
BULLET_STYLES = {
    "default": {
        0: ("□ ", " "),      # 레벨 0: 빈 네모
        1: ("○", "   "),     # 레벨 1: 빈 원
        2: ("- ", "    "),   # 레벨 2: 대시
    },
    "filled": {
        0: ("■ ", " "),      # 채운 네모
        1: ("● ", "   "),    # 채운 원
        2: ("- ", "    "),
    },
    "numbered": {
        0: ("1. ", " "),
        1: ("1) ", "   "),
        2: ("- ", "    "),
    },
}
```

### SDK로 레벨 분석

```python
from agent import BulletFormatter as SDKBulletFormatter
from merge.formatters import BulletFormatter

# 1. SDK로 레벨 분석
sdk = SDKBulletFormatter()
levels = sdk.analyze_levels("주요 내용\n세부 설명\n추가 사항")
# → [0, 1, 1]

# 2. 사용자가 레벨 확인/수정 후 정규식으로 포맷 적용
formatter = BulletFormatter()
result = formatter.format_text(text, levels=levels)
```

### 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `format_text(text, levels)` | 지정된 레벨로 글머리 적용 |
| `auto_format(text)` | 자동 레벨 감지 후 적용 |
| `normalize_style(text)` | 기존 글머리를 현재 스타일로 통일 |
| `convert_style(text, style)` | 다른 스타일로 변환 |
| `parse_items(text)` | 글머리 목록 파싱 |

---

## CaptionFormatter (캡션)

### 기본 사용법

```python
from merge.formatters import CaptionFormatter

formatter = CaptionFormatter()

# 표준 양식 변환
result = formatter.auto_format("그림1 테스트")
print(result.formatted_text)  # "그림 1. 테스트"

# 대괄호 형식
result = formatter.to_bracket_format("표 1. 연구 결과")
print(result.formatted_text)  # "[표 연구 결과]"

# 번호 재정렬
result = formatter.to_parenthesis_format("그림 1. 테스트")
print(result.formatted_text)  # "그림 (1) 테스트"
```

### HWPX 파일 처리

```python
# 캡션 조회
captions = formatter.get_all_captions("document.hwpx")
for cap in captions:
    print(f"{cap.caption_type}: {cap.text}")

# 번호 재정렬
renumbered = formatter.renumber_captions(captions, by_type=True)

# HWPX 파일에 적용
formatter.apply_to_hwpx("input.hwpx", "output.hwpx", to_bracket=True)
```

### SDK로 제목 추출

```python
from agent import CaptionFormatter as SDKCaptionFormatter

# 복잡한 패턴도 처리 가능
sdk = SDKCaptionFormatter()
title = sdk.extract_title_with_sdk("Figure3-Sample Data")
# → "Sample Data"

title = sdk.extract_title_with_sdk("Table1:Results")
# → "Results"
```

### 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `auto_format(text)` | 표준 양식으로 변환 |
| `to_standard_format(text)` | "그림 1. 제목" 형식 |
| `to_bracket_format(text)` | "[표 제목]" 형식 |
| `to_parenthesis_format(text)` | "그림 (1) 제목" 형식 |
| `remove_number(text)` | 번호 제거 |
| `extract_title(text)` | 제목만 추출 |
| `renumber_captions(captions)` | 번호 재정렬 |
| `get_all_captions(hwpx_path)` | HWPX에서 캡션 조회 |
| `apply_to_hwpx(...)` | HWPX 파일에 변환 적용 |

---

## 새 포맷터 추가 방법

### 1. 정규식 기반 포맷터 (merge/formatters/)

```python
# merge/formatters/my_formatter.py
from dataclasses import dataclass
from typing import List, Optional
from .bullet_formatter import FormatResult  # 공통 결과 클래스 재사용

@dataclass
class MyItem:
    """포맷터 항목 데이터"""
    text: str
    level: int = 0

class MyFormatter:
    """새 포맷터"""

    def __init__(self, style: str = "default"):
        self.style = style

    def format_text(self, text: str, levels: Optional[List[int]] = None) -> FormatResult:
        """텍스트 변환"""
        # 구현
        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=["변환 완료"]
        )

    def auto_format(self, text: str) -> FormatResult:
        """자동 변환"""
        levels = self._detect_levels(text)
        return self.format_text(text, levels)

    def _detect_levels(self, text: str) -> List[int]:
        """레벨 자동 감지 (정규식)"""
        # 구현
        pass
```

### 2. __init__.py에 등록

```python
# merge/formatters/__init__.py
from .my_formatter import MyFormatter, MyItem

__all__ = [
    # 기존 exports...
    'MyFormatter',
    'MyItem',
]
```

### 3. SDK 분석 기능 추가 (agent/)

```python
# agent/my_formatter.py
from .sdk import ClaudeSDK
from merge.formatters.my_formatter import MyFormatter as RegexMyFormatter, FormatResult

MY_PROMPTS = {
    "analyze": """텍스트를 분석해주세요.

규칙:
- ...

결과만 반환하세요.""",
}

class MyFormatter:
    """SDK 기반 분석 + 정규식 포맷 적용"""

    def __init__(self):
        self.sdk = ClaudeSDK(timeout=30)
        self._regex_formatter = RegexMyFormatter()

    def analyze_with_sdk(self, text: str) -> List[int]:
        """SDK로 분석"""
        prompt = f"{MY_PROMPTS['analyze']}\n\n텍스트:\n{text}"
        result = self.sdk.call(prompt)

        if result.success and result.output:
            return self._parse_response(result.output)

        # fallback
        return self._regex_formatter._detect_levels(text)

    def format_text(self, text: str, use_sdk: bool = True) -> FormatResult:
        """분석 후 포맷 적용"""
        if use_sdk:
            levels = self.analyze_with_sdk(text)
        else:
            levels = self._regex_formatter._detect_levels(text)

        return self._regex_formatter.format_text(text, levels)
```

### 4. agent/__init__.py에 등록

```python
# agent/__init__.py
from .my_formatter import MyFormatter, MY_PROMPTS

__all__ = [
    # 기존 exports...
    'MyFormatter',
    'MY_PROMPTS',
]
```

---

## FormatResult 구조

모든 포맷터가 반환하는 공통 결과 클래스:

```python
@dataclass
class FormatResult:
    success: bool = True           # 변환 성공 여부
    original_text: str = ""        # 원본 텍스트
    formatted_text: str = ""       # 변환된 텍스트
    changes: List[str] = []        # 변경 내역
    errors: List[str] = []         # 오류 메시지
```

---

## 사용 예시

### ContentFormatter (통합 사용)

```python
from merge.content_formatter import ContentFormatter

formatter = ContentFormatter(style="default", use_sdk=True)

# SDK로 레벨 분석만
levels = formatter.analyze_levels_with_sdk("항목1\n항목2\n항목3")
# → [0, 1, 2]

# SDK 분석 + 정규식 포맷 적용
result = formatter.format_with_analyzed_levels("항목1\n항목2\n항목3")

# 정규식만 사용
result = formatter.format_as_bullet_list("항목1\n항목2", levels=[0, 1])
```

### 개요 트리 내용 변환

```python
from merge.content_formatter import OutlineContentFormatter

formatter = OutlineContentFormatter()

# 개요 트리의 내용 문단을 글머리 양식으로 변환
tree, changes = formatter.format_outline_content(
    outline_tree,
    use_sdk=False,           # SDK로 전체 변환
    use_sdk_for_levels=True  # SDK로 레벨만 분석
)
```
