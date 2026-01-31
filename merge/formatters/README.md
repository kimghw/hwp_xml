# Formatters 모듈

텍스트 양식 변환을 위한 정규식 기반 포맷터 모듈입니다.

## 파일 구조

```
merge/formatters/
├── __init__.py             # 모듈 진입점, 모든 클래스/함수 export
├── bullet_formatter.py     # 글머리 기호 변환
├── caption_formatter.py    # 캡션 변환
├── config_loader.py        # YAML 설정 로더
├── formatter_config.yaml   # 기본 설정 템플릿
└── README.md
```

---

## YAML 설정 사용법

### ConfigLoader

```python
from merge.formatters import ConfigLoader, load_config, save_default_config

# 방법 1: load_config 편의 함수
config = load_config("formatter_config.yaml")

# 방법 2: ConfigLoader 직접 사용
loader = ConfigLoader()
config = loader.load("formatter_config.yaml")

# YAML 문자열에서 로드
config = loader.load_from_string(yaml_string)

# 딕셔너리에서 로드
config = loader.load_from_dict(data_dict)

# 설정 저장
loader.save(config, "updated_config.yaml")

# 기본 설정 파일 생성
save_default_config("my_config.yaml")
```

### 설정 접근

```python
config = load_config("formatter_config.yaml")

# 글머리 설정
print(config.bullet.style)       # "default"
print(config.bullet.auto_detect) # True

# 캡션 설정
print(config.caption.format)     # "standard"
print(config.caption.position)   # "TOP"
print(config.caption.separator)  # ". "

# 테이블 캡션 설정
print(config.table_caption.position)     # "TOP"
print(config.table_caption.type_prefix)  # "표"

# 이미지 캡션 설정
print(config.image_caption.position)     # "BOTTOM"
print(config.image_caption.type_prefix)  # "그림"
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
#  □ 항목1
#    ○항목2
#     - 항목3

# 자동 레벨 감지
result = formatter.auto_format("1. 첫번째\n  a. 두번째\n    - 세번째")
```

### 글머리 스타일

| 스타일 | 레벨 0 | 레벨 1 | 레벨 2 |
|--------|--------|--------|--------|
| `default` | □ (빈 네모) | ○ (빈 원) | - (대시) |
| `filled` | ■ (채운 네모) | ● (채운 원) | - (대시) |
| `numbered` | 1. | 1) | - |
| `arrow` | ▶ | ▷ | - |

```python
from merge.formatters import BULLET_STYLES

BULLET_STYLES = {
    "default": {
        0: ("□ ", " "),      # (기호, 들여쓰기)
        1: ("○", "   "),
        2: ("- ", "    "),
    },
    "filled": {
        0: ("■ ", " "),
        1: ("● ", "   "),
        2: ("- ", "    "),
    },
    "numbered": {
        0: ("1. ", " "),
        1: ("1) ", "   "),
        2: ("- ", "    "),
    },
    "arrow": {
        0: ("▶ ", ""),
        1: ("▷ ", "  "),
        2: ("- ", "    "),
    },
}
```

### 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `format_text(text, levels)` | 지정된 레벨로 글머리 적용 |
| `auto_format(text)` | 자동 레벨 감지 후 적용 |
| `normalize_style(text)` | 기존 글머리를 현재 스타일로 통일 |
| `convert_style(text, target_style)` | 다른 스타일로 변환 |
| `set_style(style)` | 현재 스타일 변경 |
| `parse_items(text)` | 글머리 목록을 BulletItem 리스트로 파싱 |
| `has_bullet(text)` | 텍스트에 글머리 기호 존재 여부 |
| `apply_hierarchy(items, base_level)` | 계층적 글머리 적용 |
| `apply_flat(items, level)` | 동일 레벨 글머리 적용 |

---

## CaptionFormatter (캡션)

### 기본 사용법

```python
from merge.formatters import CaptionFormatter

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

# 표준 형식 (구분자 지정)
result = formatter.to_standard_format("그림1테스트", separator=": ")
print(result.formatted_text)  # "그림 1: 테스트"
```

### 캡션 형식 프리셋

| 프리셋 | 패턴 | 예시 |
|--------|------|------|
| `title_only` | `[{title}]` | `[연구결과]` |
| `type_title` | `[{type} {title}]` | `[표 연구결과]` |
| `type_num_title` | `[{type} {number} {title}]` | `[표 1 연구결과]` |
| `standard` | `{type} {number}{separator}{title}` | `표 1. 연구결과` |
| `parenthesis` | `{type} ({number}) {title}` | `표 (1) 연구결과` |
| `bracket_num` | `[{type} {number}] {title}` | `[표 1] 연구결과` |

### YAML 설정 적용

```python
from merge.formatters import CaptionFormatter, load_config

config = load_config("formatter_config.yaml")
formatter = CaptionFormatter()

# 설정에 따라 캡션 형식 변환
if config.caption.format == "title_only":
    result = formatter.to_bracket_format(text)
elif config.caption.format == "parenthesis":
    result = formatter.to_parenthesis_format(text)
else:
    result = formatter.to_standard_format(text, separator=config.caption.separator)
```

### HWPX 파일 처리

```python
from merge.formatters import CaptionFormatter, get_captions, print_captions, renumber_captions

formatter = CaptionFormatter()

# 캡션 조회
captions = formatter.get_all_captions("document.hwpx")
# 또는 편의 함수 사용
captions = get_captions("document.hwpx")

# 캡션 출력
print_captions(captions)

# 번호 재정렬
renumbered = formatter.renumber_captions(
    captions,
    by_type=True,      # 유형별 따로 번호 (그림 1,2... 표 1,2...)
    start_number=1
)
# 또는 편의 함수 사용
renumbered = renumber_captions("document.hwpx", by_type=True)

# HWPX 파일에 적용
formatter.apply_to_hwpx(
    "input.hwpx",
    "output.hwpx",
    to_bracket=True,
    keep_auto_num=False,
    renumber=True
)

# 대괄호 형식 적용 (편의 메서드)
formatter.apply_bracket_format("input.hwpx", "output.hwpx")

# 번호만 재정렬
formatter.renumber_hwpx("input.hwpx", "output.hwpx")

# 캡션 위치 변경
formatter.set_caption_position("input.hwpx", "TOP", "output.hwpx")
formatter.set_caption_to_top("input.hwpx")     # 위로
formatter.set_caption_to_bottom("input.hwpx")  # 아래로

# 글자처럼 취급 설정
formatter.set_table_as_char("input.hwpx")      # 테이블 글자처럼
formatter.set_image_as_anchor("input.hwpx")    # 이미지 어울림
formatter.set_all_as_char("input.hwpx")        # 전체 글자처럼
```

### 주요 메서드

| 메서드 | 설명 |
|--------|------|
| `auto_format(text)` | 표준 양식으로 자동 변환 |
| `to_standard_format(text, separator)` | "그림 1. 제목" 형식 |
| `to_bracket_format(text)` | "[제목]" 형식 (제목만) |
| `to_parenthesis_format(text)` | "그림 (1) 제목" 형식 |
| `normalize_format(text)` | 공백/구두점 정규화 |
| `remove_number(text)` | 번호 제거 |
| `extract_title(text)` | 제목만 추출 |
| `get_type_prefix(text)` | 유형 접두어 추출 |
| `get_all_captions(hwpx_path)` | HWPX에서 캡션 조회 |
| `renumber_captions(captions)` | 번호 재정렬 |
| `apply_to_hwpx(...)` | HWPX 파일에 변환 적용 |
| `set_caption_position(...)` | 캡션 위치 변경 |
| `set_treat_as_char(...)` | 글자처럼 취급 설정 |

---

## TableCaptionConfig, ImageCaptionConfig 설정

테이블과 이미지에 대해 별도 캡션 설정 가능:

```yaml
# formatter_config.yaml

table_caption:
  format: standard
  position: TOP           # 테이블은 캡션이 위로
  renumber: true
  renumber_by_type: true
  start_number: 1
  separator: ". "
  keep_auto_num: false
  type_prefix: "표"       # 유형 접두어

image_caption:
  format: standard
  position: BOTTOM        # 이미지는 캡션이 아래로
  renumber: true
  renumber_by_type: true
  start_number: 1
  separator: ". "
  keep_auto_num: false
  type_prefix: "그림"     # 유형 접두어
```

```python
from merge.formatters import load_config, TableCaptionConfig, ImageCaptionConfig

config = load_config("formatter_config.yaml")

# 테이블 캡션 설정
table_cfg: TableCaptionConfig = config.table_caption
print(table_cfg.position)     # "TOP"
print(table_cfg.type_prefix)  # "표"

# 이미지 캡션 설정
image_cfg: ImageCaptionConfig = config.image_caption
print(image_cfg.position)     # "BOTTOM"
print(image_cfg.type_prefix)  # "그림"
```

---

## FormatResult 구조

모든 포맷터가 반환하는 공통 결과 클래스:

```python
from dataclasses import dataclass, field
from typing import List

@dataclass
class FormatResult:
    success: bool = True           # 변환 성공 여부
    original_text: str = ""        # 원본 텍스트
    formatted_text: str = ""       # 변환된 텍스트
    changes: List[str] = field(default_factory=list)   # 변경 내역
    errors: List[str] = field(default_factory=list)    # 오류 메시지
```

사용 예시:
```python
result = formatter.auto_format("그림1 테스트")
if result.success:
    print(result.formatted_text)
    print(f"변경 사항: {result.changes}")
else:
    print(f"오류: {result.errors}")
```

---

## 새 포맷터 추가 방법

### 1. 포맷터 클래스 생성 (merge/formatters/)

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
        # 변환 로직 구현
        formatted = text  # 실제 변환 로직
        return FormatResult(
            success=True,
            original_text=text,
            formatted_text=formatted,
            changes=["변환 완료"]
        )
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

### 3. config_loader.py에 설정 추가 (선택)

```python
# merge/formatters/config_loader.py
from dataclasses import dataclass

@dataclass
class MyConfig:
    option1: str = "default"
    option2: bool = True

@dataclass
class FormatterConfig:
    # 기존 설정...
    my: MyConfig = field(default_factory=MyConfig)
```

### 4. formatter_config.yaml에 기본값 추가 (선택)

```yaml
# formatter_config.yaml
my:
  option1: default
  option2: true
```
