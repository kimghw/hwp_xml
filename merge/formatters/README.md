# Formatters 모듈

텍스트 양식 변환을 위한 정규식 기반 포맷터 모듈

## 파일 구조

```
merge/formatters/
├── __init__.py                      # 모듈 진입점
├── base_formatter.py                # 포맷터 추상 인터페이스
├── bullet_formatter.py              # 글머리 기호 변환
├── caption_formatter.py             # 캡션 변환
├── object_formatter.py              # 테이블/이미지 정렬 설정
├── config_loader.py                 # YAML 설정 로더
├── content_formatter_config.yaml    # 콘텐츠 포맷터 설정
├── table_formatter_config.yaml      # 테이블 포맷터 설정
└── README.md
```

---

## 주요 클래스

### BaseFormatter (추상 인터페이스)

모든 포맷터가 상속받아야 하는 인터페이스

```python
class BaseFormatter(ABC):
    @abstractmethod
    def format_with_levels(texts: List[str], levels: List[int]) -> FormatResult
    @abstractmethod
    def format_text(text: str, levels=None, auto_detect=True) -> FormatResult
    def has_format(text: str) -> bool
    def get_style_name() -> str
```

### BulletFormatter (글머리 기호)

텍스트를 글머리 기호 양식으로 변환

```python
from merge.formatters import BulletFormatter

formatter = BulletFormatter(style="default")
result = formatter.format_text("항목1\n항목2", levels=[0, 1])
```

**내장 스타일**:

| 스타일 | Level 1 | Level 2 | Level 3 |
|--------|---------|---------|---------|
| `default` | □ | ○ | - |
| `filled` | ■ | ● | - |
| `numbered` | 1. | 1) | - |
| `arrow` | ▶ | ▷ | - |

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

### ObjectFormatter (테이블/이미지 정렬)

테이블/이미지 서식 설정 (글자처럼 취급, 정렬)

```python
from merge.formatters import ObjectFormatter

fmt = ObjectFormatter()
fmt.set_all_format("input.hwpx", "output.hwpx", treat_as_char=True)
```

**테이블/이미지 정렬 원리**:
```
HWP 가운데 정렬 = treatAsChar(1) + paraPrIDRef(20) + horzAlign(CENTER)
```

**중요**: 인라인 `<paraShape align="CENTER" />`는 HWP에서 인식 안됨!
→ 반드시 header.xml의 스타일(paraPrIDRef)을 참조해야 함

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
    # 기본 인터페이스
    BaseFormatter,

    # 글머리 기호
    BulletFormatter, BulletItem, FormatResult, BULLET_STYLES,

    # 캡션
    CaptionFormatter, CaptionInfo, CAPTION_PATTERNS,
    get_captions, print_captions, renumber_captions,

    # 객체 서식
    ObjectFormatter,

    # 설정
    ConfigLoader, FormatterConfig, BulletConfig, CaptionConfig,
    CaptionFormatPreset, TableCaptionConfig, ImageCaptionConfig,
    load_config, create_default_config, save_default_config,
)
```

---

## 파이프라인 통합 위치

| Step | 모듈 | 역할 |
|------|------|------|
| 3 | BulletFormatter | 글머리 기호 정규화 |
| 4 | CaptionFormatter | 캡션 정규화 |
| 6 | ObjectFormatter | 테이블/이미지 정렬 |

---

## 최근 변경사항 (3일 이내)

### 0b0209b: ObjectFormatter 추가
- 테이블/이미지 정렬 문제 해결
- `treatAsChar` + `paraPrIDRef` 기반 정렬
- 파이프라인 step 6에 통합
- **핵심 이슈 해결**: XML 기반 정렬이 HWP에서 인식 안 되는 문제 → paraPrIDRef 사용

### 649383f: YAML 프롬프트 외부화
- SDK 프롬프트를 YAML 파일로 분리
- 커스텀 스타일 지원
- context별(본문/테이블) instruction 구분
- 테이블 병합 신뢰성 향상

### 1fb8aee: BaseFormatter 도입
- 추상 인터페이스 추가
- 개요/본문에 별도 포맷터 주입 지원
- BulletFormatter가 BaseFormatter 상속

---

## 주의사항

1. **정규식 vs SDK**: formatters는 정규식 기반 (빠름), agent/는 AI 기반 (정확)
2. **XML 네임스페이스**: 등록 필수 (`ET.register_namespace`)
3. **paraPrIDRef 값**: 0=좌측, 20=중앙 (header.xml 확인 필요)
4. **임시 파일**: ZIP 작업 후 move/copy로 원본 보호

### 새 포맷터 추가 방법

```python
class CustomFormatter(BaseFormatter):
    def format_with_levels(self, texts, levels):
        # 커스텀 로직
        pass

    def format_text(self, text, levels=None, auto_detect=True):
        # 커스텀 로직
        pass
```
