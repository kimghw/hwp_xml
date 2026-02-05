# HWPX 병합 모듈

개요(Outline) 기준으로 HWPX 파일 병합 및 형식 검토

## 용어 정의

| 용어 | 설명 |
|------|------|
| **Template** | 기준 파일 (첫 번째 파일). 테이블 구조와 필드명을 정의 |
| **Addition** | 추가 파일 (두 번째 이후 파일). 템플릿에 병합될 데이터 |

## 모듈 구조

```
merge/
├── __init__.py             # 모듈 진입점, 공개 API
├── merge_pipeline.py       # 메인 파이프라인 (6단계)
├── merge_hwpx.py           # 핵심 병합 로직 (HwpxMerger)
├── merge_table.py          # 테이블 필드명 매칭 및 병합
├── run_merge.py            # CLI 실행 스크립트
├── parser.py               # HWPX 문단 파싱
├── models.py               # Paragraph, OutlineNode 등
├── outline.py              # 개요 트리 처리
├── content_formatter.py    # 개요 내용 양식 변환
├── format_validator.py     # 형식 검증/수정
├── merge_with_review.py    # 병합 + Agent 통합
│
├── field/                  # 필드명 관리 모듈
├── formatters/             # 정규식 기반 포맷터
└── table/                  # 테이블 병합 서브모듈
```

## 파이프라인 흐름

```
[HWPX 파일들]
    ↓
[1/6] HwpxParser.parse() → HwpxData
    ↓
[2/6] merge_outline_trees() → 병합된 개요 트리
    ↓
[3/6] FormatFixer → 글머리/캡션 자동 수정
    ↓
[4/6] TableMergeHandler.collect_and_merge() → 테이블 데이터 병합
    ↓
[5/6] HwpxMerger.merge_with_tree() → HWPX 파일 생성
    ↓
[6/6] ObjectFormatter → 테이블/이미지 가운데 정렬
    ↓
[output.hwpx]
```

## 사용법

### CLI

```bash
# 구조 확인
python merge_hwpx.py file1.hwpx file2.hwpx

# 개요 목록 출력
python merge_hwpx.py --list-outlines file1.hwpx file2.hwpx

# 병합
python merge_hwpx.py -o output.hwpx file1.hwpx file2.hwpx

# 특정 개요 제외
python merge_hwpx.py -o output.hwpx --exclude "1. 서론" file1.hwpx file2.hwpx
```

### Python

```python
from merge import MergePipeline

pipeline = MergePipeline()
result = pipeline.merge(
    ["file1.hwpx", "file2.hwpx"],
    "output.hwpx",
    auto_fix=True
)
print(f"성공: {result.success}")
print(f"테이블 병합: {len(result.table_merges)}개")
```

## 병합 규칙

| 조건 | 동작 |
|------|------|
| 같은 level + 같은 이름 | 내용 이어붙이기 |
| 같은 level + 다른 이름 | 개요 추가 |
| 일반 문단 | 위 개요에 종속 |

## 형식 검토 규칙

- **캡션**: `표 N. 제목`, `그림 N. 제목`
- **글머리**: □ (1단계) → ○ (2단계) → - (3단계)

---

## 최근 변경사항 (3일 이내)

### d3d89c1: 테이블 병합 단계 분리
- 테이블 병합을 파일 생성 전 [4/6] 단계로 분리
- `merge_with_tree()`에서 table_merge_data 파라미터 제거
- 파이프라인 순차 처리 명확화

### c9e498a: SDK 의존성 제거
- SDK formatter 의존성 제거
- formatters 모듈 직접 사용 (BulletFormatter, CaptionFormatter)
- 유지보수성 향상

### 0b0209b: ObjectFormatter 추가
- 테이블/이미지 `treatAsChar` + `paraPrIDRef` 기반 정렬
- 파이프라인 step 6에 통합

### 버그 수정 (ee3da26, b8c1417)
- add_ 필드 추출 버그 수정 (header row에서 추출 안됨)
- 중복 bullet formatting 방지
- multi-section 손실 수정
- _shift_rows_down: rowspan 셀의 end_row 업데이트 추가

---

## 주의사항

1. **테이블 병합 순서**: 반드시 파일 생성 전에 처리
2. **개요 제외**: 정확한 이름 또는 접두사 매칭 (예: `"3."`)
3. **중첩 테이블**: `table_idx`는 최상위 테이블만 참조
4. **다중 섹션 문서**: section2.xml+ 콘텐츠 복사 필수

---

## 서브모듈

### field/ (필드명 관리)

테이블 필드명 자동 생성 및 시각화 모듈.

```python
from merge.field import insert_auto_fields, fill_empty_fields, colorize_by_field

# 자동 필드명 생성
insert_auto_fields("template.hwpx")

# 빈 셀에 위 셀 필드명 복사
fill_empty_fields("template.hwpx")

# 필드별 배경색 설정
colorize_by_field("template.hwpx", "output.hwpx")
```

### formatters/ (정규식 기반 포맷터)

텍스트 양식 변환을 위한 정규식 기반 모듈. 자세한 내용은 [formatters/README.md](formatters/README.md) 참조.

```python
from merge.formatters import BulletFormatter, CaptionFormatter, load_config

# YAML 설정 로드
config = load_config("content_formatter_config.yaml")

# 글머리 기호 변환
bullet = BulletFormatter(style=config.bullet.style)

# 캡션 변환
caption = CaptionFormatter()
result = caption.to_bracket_format("표 1. 연구결과")  # → "[연구결과]"
```

**글머리 스타일**: `default` | `filled` | `numbered` | `arrow`

**캡션 프리셋**: `title_only` | `type_title` | `type_num_title` | `standard` | `parenthesis` | `bracket_num`

### table/ (테이블 병합)

HWPX 테이블 파싱 및 데이터 병합 모듈. 자세한 내용은 [table/README.md](table/README.md) 참조.

```python
from merge.table import TableParser, TableMerger

# 테이블 파싱
parser = TableParser()
tables = parser.parse_hwpx("document.hwpx")

# 테이블 병합
merger = TableMerger()
merger.merge(template_table, addition_table)
```

---

## 데이터 구조

```python
@dataclass
class OutlineNode:
    level: int              # 개요 레벨 (0~6)
    name: str               # 개요 텍스트
    paragraphs: List[Paragraph]
    children: List[OutlineNode]

@dataclass
class Paragraph:
    index: int
    is_outline: bool
    level: int
    text: str
    has_table: bool
    has_image: bool
```

## 필드명 접두사 규칙

| 접두사 | 설명 |
|--------|------|
| `header_` | 테이블 상단의 헤더 셀 |
| `stub_` | 테이블 좌측의 고정 값 (행 헤더), 병합 시 새 셀 생성 |
| `gstub_` | 테이블 좌측의 고정 값, 병합 시 셀 병합됨 (grouped stub) |
| `input_` | 입력이 필요한 셀, 병합 시 새 셀 생성 |
| `add_` | 병합 시 셀 생성 없이 내용만 합침 |

### nc_name 자동 설정 규칙

우선순위 순서대로 적용:

1. **header_**: 최상단 행 + 배경색 있음, 또는 헤더와 연결되고 배경색 동일
2. **add_**: 최상단 행 + 텍스트 30자 이상 + 배경색 없음
3. **stub_**: 텍스트 있음 + 오른쪽에 빈 셀 존재 + rowspan 없음
4. **gstub_**: 텍스트 있음 + 오른쪽에 빈 셀 존재 + rowspan 있음 (병합된 셀)
5. **input_**: 빈 셀 (기본값)

#### 동일 필드명 그룹화

빈 셀(`input_`)의 경우:
- 위로 이동 시 첫 번째 만나는 셀의 값이 동일하고
- 좌측으로 이동 시 텍스트 있는 셀이 없으면
- → 같은 nc_name 사용

```
┌─────────┬─────────┬─────────┐
│ header_ │ header_ │ header_ │  ← 배경색 있음
├─────────┼─────────┼─────────┤
│ stub_A  │ input_1 │ input_2 │  ← 빈 셀은 각각 다른 이름
├─────────┼─────────┼─────────┤
│ gstub_B │ input_1 │ input_2 │  ← 위 셀과 헤더 동일 + 좌측 빈 → 같은 이름
│ (병합)  │         │         │
└─────────┴─────────┴─────────┘
```

### 병합 동작 예시

```
┌─────────┬─────────┬─────────┐
│ header_ │ header_ │ header_ │  ← 헤더 (유지)
├─────────┼─────────┼─────────┤
│ stub_   │ input_  │ input_  │  ← stub_/input_: 새 행 추가
├─────────┼─────────┼─────────┤
│ gstub_  │ add_    │ add_    │  ← gstub_/add_: 기존 셀에 내용 추가 (병합)
└─────────┴─────────┴─────────┘
```

## 테이블 병합 (table_merger.py)

```python
from table_merger import TableMerger, HeaderConfig

merger = TableMerger()
table = merger.load_base_table("template.hwpx", 0)
merger.merge_data([{"필드명": "값"}], mode="smart")
merger.save("output.hwpx")
```

**HeaderConfig 옵션**:
- `action='extend'`: 기존 헤더 rowspan 확장
- `action='new'`: 새 헤더 생성
- `action='data'`: 데이터 셀

## 테이블 병합 흐름 (HwpxMerger)

```
Template 파일 (첫 번째)          Addition 파일 (두 번째 이후)
┌─────────────────────┐         ┌─────────────────────┐
│ 테이블 A (필드 있음) │         │ 테이블 A' (필드 있음)│
│ - header_col1       │         │ - header_col1       │
│ - input_data        │         │ - input_data = "값" │
├─────────────────────┤         ├─────────────────────┤
│ 테이블 B (필드 없음) │         │ 테이블 C (필드 없음) │
└─────────────────────┘         └─────────────────────┘
           │                              │
           └──────────┬───────────────────┘
                      ▼
              병합 결과 (Output)
┌─────────────────────────────────────────┐
│ 테이블 A (머지됨)                        │
│ - header_col1                           │
│ - input_data = "값" ← Addition에서 병합 │
├─────────────────────────────────────────┤
│ 테이블 B (Template에서 유지)             │
├─────────────────────────────────────────┤
│ 테이블 C (Addition에서 복사)             │
└─────────────────────────────────────────┘
```

### 테이블 처리 규칙

| 조건 | 동작 |
|------|------|
| Addition 테이블에 필드명 있음 + Template과 일치 | TableMerger로 머지 |
| Addition 테이블에 필드명 없음 | 테이블 그대로 복사 |
| Template 테이블 | 그대로 유지 |
