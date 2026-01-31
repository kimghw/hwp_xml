# Formatters 모듈

텍스트 양식 변환을 위한 정규식 기반 포맷터 모듈입니다.

## 구조

```
merge/formatters/          # 정규식 기반 (포맷 적용)
├── bullet_formatter.py    # 글머리 기호 변환
├── caption_formatter.py   # 캡션 변환
├── config_loader.py       # YAML 설정 로더
├── formatter_config.yaml  # 기본 설정 템플릿
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

## YAML 설정 파일

포맷터 옵션을 YAML 파일로 설정할 수 있습니다.

### 설정 파일 구조

```yaml
# formatter_config.yaml

# 글머리 기호 설정 (BulletFormatter)
bullet:
  style: default        # default, filled, numbered
  auto_detect: true     # 자동 레벨 감지
  # 커스텀 글머리 (선택적)
  # custom_bullets:
  #   0: { symbol: "□ ", indent: " " }
  #   1: { symbol: "○", indent: "   " }
  #   2: { symbol: "- ", indent: "    " }

# 캡션 설정 (CaptionFormatter 공통)
caption:
  format: standard      # standard, bracket, parenthesis
  position: TOP         # TOP, BOTTOM, LEFT, RIGHT
  renumber: true        # 번호 재정렬
  renumber_by_type: true  # 유형별 따로 번호 (그림 1,2... 표 1,2...)
  start_number: 1
  separator: ". "       # 번호와 제목 사이 구분자
  keep_auto_num: false  # autoNum 유지 여부

# 테이블 캡션 설정
table_caption:
  format: standard
  position: TOP         # 테이블은 캡션이 위로
  renumber: true
  renumber_by_type: true
  start_number: 1
  separator: ". "
  keep_auto_num: false

# 이미지 캡션 설정
image_caption:
  format: standard
  position: BOTTOM      # 이미지는 캡션이 아래로
  renumber: true
  renumber_by_type: true
  start_number: 1
  separator: ". "
  keep_auto_num: false
```

### 설정 로더 사용

```python
from merge.formatters import load_config, save_default_config, ConfigLoader

# 설정 로드
config = load_config("formatter_config.yaml")
print(f"글머리 스타일: {config.bullet.style}")
print(f"캡션 형식: {config.caption.format}")
print(f"테이블 캡션 위치: {config.table_caption.position}")

# 기본 설정 파일 생성
save_default_config("my_config.yaml")

# 설정 수정 후 저장
loader = ConfigLoader()
config = loader.load("config.yaml")
config.caption.format = "bracket"
loader.save(config, "updated_config.yaml")
```

---

## BulletFormatter (글머리 기호)

### 기본 사용법

```python
from merge.formatters import BulletFormatter, load_config

# YAML 설정으로 초기화
config = load_config("formatter_config.yaml")
formatter = BulletFormatter(style=config.bullet.style)

# 또는 직접 스타일 지정
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
from merge.formatters import CaptionFormatter, load_config

formatter = CaptionFormatter()

# 표준 양식 변환
result = formatter.auto_format("그림1 테스트")
print(result.formatted_text)  # "그림 1. 테스트"

# 대괄호 형식 (제목만)
result = formatter.to_bracket_format("표 1. 연구 결과")
print(result.formatted_text)  # "[연구 결과]"

# 괄호 형식
result = formatter.to_parenthesis_format("그림 1. 테스트")
print(result.formatted_text)  # "그림 (1) 테스트"
```

### YAML 설정 적용

```python
from merge.formatters import CaptionFormatter, load_config

config = load_config("formatter_config.yaml")

formatter = CaptionFormatter()

# 설정에 따라 캡션 형식 변환
if config.caption.format == "bracket":
    result = formatter.to_bracket_format(text)
elif config.caption.format == "parenthesis":
    result = formatter.to_parenthesis_format(text)
else:
    result = formatter.to_standard_format(text, separator=config.caption.separator)
```

### HWPX 파일 처리

```python
# 캡션 조회
captions = formatter.get_all_captions("document.hwpx")
for cap in captions:
    print(f"{cap.caption_type}: {cap.text}")

# 번호 재정렬
renumbered = formatter.renumber_captions(
    captions,
    by_type=config.caption.renumber_by_type,
    start_number=config.caption.start_number
)

# HWPX 파일에 적용
formatter.apply_to_hwpx(
    "input.hwpx",
    "output.hwpx",
    to_bracket=(config.caption.format == "bracket"),
    keep_auto_num=config.caption.keep_auto_num,
    renumber=config.caption.renumber
)

# 캡션 위치 변경
formatter.set_caption_position(
    "input.hwpx",
    position=config.table_caption.position,
    output_path="output.hwpx"
)
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
| `to_bracket_format(text)` | "[제목]" 형식 |
| `to_parenthesis_format(text)` | "그림 (1) 제목" 형식 |
| `remove_number(text)` | 번호 제거 |
| `extract_title(text)` | 제목만 추출 |
| `renumber_captions(captions)` | 번호 재정렬 |
| `get_all_captions(hwpx_path)` | HWPX에서 캡션 조회 |
| `apply_to_hwpx(...)` | HWPX 파일에 변환 적용 |
| `set_caption_position(...)` | 캡션 위치 변경 |

---

## ContentFormatter (통합 사용)

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

## 새 포맷터 추가 방법

### 1. 정규식 기반 포맷터 (merge/formatters/)

```python
# merge/formatters/my_formatter.py
from dataclasses import dataclass
from typing import List, Optional
from .bullet_formatter import FormatResult

@dataclass
class MyItem:
    text: str
    level: int = 0

class MyFormatter:
    def __init__(self, style: str = "default"):
        self.style = style

    def format_text(self, text: str, levels: Optional[List[int]] = None) -> FormatResult:
        # 구현
        return FormatResult(success=True, formatted_text=formatted)
```

### 2. __init__.py에 등록

```python
from .my_formatter import MyFormatter, MyItem

__all__ = [
    # 기존 exports...
    'MyFormatter',
    'MyItem',
]
```

### 3. config_loader.py에 설정 추가

```python
@dataclass
class MyConfig:
    option1: str = "default"
    option2: bool = True
```
