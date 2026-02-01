# Formatters 모듈

텍스트 양식 변환을 위한 정규식 기반 포맷터 모듈

## 파일 구조

```
merge/formatters/
├── __init__.py                      # 모듈 진입점
├── bullet_formatter.py              # 글머리 기호 변환
├── caption_formatter.py             # 캡션 변환
├── config_loader.py                 # YAML 설정 로더
├── content_formatter_config.yaml    # 콘텐츠 포맷터 설정
├── table_formatter_config.yaml      # 테이블 포맷터 설정
└── README.md
```

---

## 주요 클래스

### BulletFormatter (글머리 기호)
텍스트를 글머리 기호 양식으로 변환

```python
from merge.formatters import BulletFormatter

formatter = BulletFormatter(style="default")
result = formatter.format_text("항목1\n항목2", levels=[0, 1])
```

**스타일**: `default`, `filled`, `numbered`, `arrow`

### CaptionFormatter (캡션)
HWPX 파일에서 캡션 조회 및 양식 변환

```python
from merge.formatters import CaptionFormatter

formatter = CaptionFormatter()
captions = formatter.get_all_captions("document.hwpx")
result = formatter.to_bracket_format("표 1. 제목")  # "[제목]"
```

**주요 기능**:
- `get_all_captions()`: HWPX에서 캡션 조회
- `to_bracket_format()`: 대괄호 형식 변환
- `to_standard_format()`: 표준 형식 변환
- `renumber_captions()`: 번호 재정렬
- `set_caption_position()`: 캡션 위치 변경
- `set_treat_as_char()`: 글자처럼 취급 설정

### ConfigLoader (설정 로더)
YAML 설정 파일 로드

```python
from merge.formatters import load_config

config = load_config("content_formatter_config.yaml")
print(config.bullet.style)
print(config.caption.format)
```

---

## Export 목록

```python
from merge.formatters import (
    # 글머리 기호
    BulletFormatter, BulletItem, FormatResult, BULLET_STYLES,

    # 캡션
    CaptionFormatter, CaptionInfo, CAPTION_PATTERNS,
    get_captions, print_captions, renumber_captions,

    # 설정
    ConfigLoader, FormatterConfig, BulletConfig, CaptionConfig,
    CaptionFormatPreset, TableCaptionConfig, ImageCaptionConfig,
    load_config, create_default_config, save_default_config,
)
```
