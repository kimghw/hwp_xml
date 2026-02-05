# merge/table 모듈

HWPX 테이블 파싱 및 데이터 병합 모듈

## 파일 구조

```
merge/table/
├── __init__.py           # 모듈 초기화 및 exports
├── models.py             # 데이터 모델 (CellInfo, TableInfo 등)
├── parser.py             # HWPX 테이블 파싱
├── merger.py             # 테이블 셀 내용 병합
├── row_extractor.py      # 테이블 행 데이터 추출
├── gstub_cell_splitter.py # gstub 셀 나누기
├── row_builder.py        # 테이블 행 자동 생성
└── formatter_config.py   # add_ 필드 포맷터 설정 로더
```

## 필드명 접두사 우선순위

```python
FIELD_PRIORITY = {
    "gstub_": 5,    # 그룹화 스텁 (rowspan 확장)
    "input_": 4,    # 입력 필드 (데이터 추가)
    "stub_": 3,     # 스텁 (new row marker)
    "data_": 2,     # 데이터 (불변)
    "header_": 1,   # 헤더 (불변)
}
```

## 모듈별 역할

### models.py

| 클래스 | 설명 |
|--------|------|
| `CellInfo` | 테이블 셀 정보 (위치, 크기, 텍스트, 필드명) |
| `TableInfo` | 테이블 정보 (행/열 수, 셀 딕셔너리, 필드 매핑) |
| `HeaderConfig` | 새 행 추가 시 헤더 열 설정 |
| `RowAddPlan` | 행 추가 계획 |

### parser.py

HWPX 파일에서 테이블 파싱

```python
from merge.table import TableParser

parser = TableParser(auto_field_names=True)
tables = parser.parse_tables("input.hwpx")

for table in tables:
    print(f"테이블: {table.row_count}행 x {table.col_count}열")
```

### merger.py

Base 파일에 데이터 병합. SDK 기반 글머리 포맷팅 지원.

**생성자 옵션:**

| 파라미터 | 기본값 | 설명 |
|----------|--------|------|
| `validate_format` | `False` | add_/input_ 필드 병합 전 형식 검증 |
| `sdk_validator` | `None` | Claude Code SDK 검증 함수 (외부 주입) |
| `formatter_config_path` | `None` | 포맷터 설정 YAML 파일 경로 |
| `use_formatter` | `True` | add_ 필드에 포맷터 적용 |
| `format_add_content` | `True` | add_ 필드에 글머리 기호 포맷팅 적용 |
| `use_sdk_for_levels` | `True` | SDK로 레벨 분석 (글머리 제거 + 레벨 분석) |

**필드명 접두사별 동작:**

| 접두사 | 동작 |
|--------|------|
| `header_` | 유지 (헤더 행) |
| `data_` | 유지 (데이터 표시용) |
| `add_` | 기존 셀에 내용 추가 (SDK로 글머리 포맷팅) |
| `stub_` | 새 행 생성 시 식별자 |
| `gstub_` | 같은 값이면 rowspan 확장 (그룹 stub) |
| `input_` | 빈 셀 채우기, 부족시 행 추가 |

**주요 메서드:**

| 메서드 | 설명 |
|--------|------|
| `load_base_table(hwpx_path, table_index)` | 기준 테이블 로드 |
| `merge_data(data_dict)` | 단순 필드-값 병합 |
| `merge_with_stub(data_list)` | stub/gstub/input 기반 병합 |
| `save(output_path)` | 결과 저장 |

### gstub_cell_splitter.py

gstub 범위 내 행 삽입 및 rowspan 확장

**주요 기능:**
- gstub 범위 내 새 행 삽입
- 기존 행들을 밀어내고 rowspan 확장
- 새 행의 셀 생성 (input_, stub_, data_ 등)
- 다중 gstub 열 지원

### row_builder.py

헤더 기반 자동 행 추가

**주요 메서드:**

| 메서드 | 설명 |
|--------|------|
| `add_rows_smart(data_list)` | 필드명 접두사 분석하여 자동 행 추가 |
| `add_rows_auto(data_list, header_names)` | 헤더 이름 기준 자동 행 추가 |
| `add_row_with_headers(header_configs)` | 헤더 설정에 따른 행 추가 |

### formatter_config.py

add_ 필드용 포맷터 설정 로드 (YAML)

| 클래스/함수 | 설명 |
|-------------|------|
| `TableFormatterConfigLoader` | YAML 설정 파일 로더 |
| `TableFormatterConfig` | 테이블 포맷터 통합 설정 |
| `FieldFormatterConfig` | 필드별 포맷터 설정 |
| `load_table_formatter_config()` | 설정 로드 편의 함수 |
| `format_add_field_value()` | add_ 필드값에 포맷터 적용 |

---

## 최근 변경사항 (3일 이내)

### b8c1417: 테이블 병합 버그 수정 및 리팩토링

**핵심 수정사항:**
- **multi-section 손실**: section1.xml+ 복사 누락 수정
- **multi-gstub 열 처리**: 모든 gstub 열 rowspan 확장
- **_shift_rows_down**: rowspan 셀의 end_row 업데이트 추가 ★
- **add_ 필드 구분자 누적**: 필드별 초기화

### 70ce1b4: 빈 입력 행 스킵
- data_ 필드만 있는 행 스킵
- 모든 input_ 값이 빈 행 스킵
- 불필요한 행 삽입 방지

### a2fc8b3: TableMergePlan 이동
- TableMergePlan을 merge_table.py로 이동
- `collect_table_data()` 메서드 추가
- 파이프라인 step 4를 4-1(body), 4-2(table)로 분리

---

## 개발자 주의사항

### rowspan 관련 (중요!)

```python
# _extend_rowspan에서 end_row도 함께 증가
cell.row_span += 1
cell.end_row += 1  # 필수!

# _shift_rows_down에서 rowspan 셀의 end_row 업데이트
if cell.row < from_row <= cell.end_row:
    cell.end_row += 1
```

### add_ 필드 처리
- add_data 딕셔너리는 필드별로 초기화 (누적 방지)
- 구분자(separator)는 YAML 설정에서 로드

### gstub 범위 계산
- 여러 gstub 열: 가장 작은 end_row 기준으로 삽입 위치 결정
- gstub 범위 내에만 행 삽입 가능

### 필드명 충돌
- 같은 필드명이 여러 테이블에 존재 가능
- `List[(table_idx, row, col)]` 구조 고려

### 다중 섹션 문서
- section1.xml 이후의 section2.xml+ 콘텐츠를 템플릿에서 복사 필수

---

## gstub 범위 내 행 삽입 알고리즘

```
1. 매칭되는 gstub 셀 찾기
   - gstub_values의 모든 field_name에 대해 같은 text 값인 셀 찾기

2. 삽입 위치 결정
   - insert_row_idx = min(cell.end_row for cell in matching_cells) + 1

3. 기존 행 밀어내기 (_shift_rows_down)
   - cellAddr rowAddr 업데이트
   - cells dict 업데이트 (row, end_row 조정)
   - field_to_cell 매핑 업데이트

4. 모든 gstub 셀의 rowspan 확장
   - 새 행이 범위 안에 들어가므로 모든 매칭 gstub 확장

5. 새 행의 셀 생성
   - gstub 열: 셀 없음 (rowspan으로 커버)
   - stub 열: stub_values의 값 사용
   - input 열: input_values의 값 사용
```

---

## 파일 복잡도

| 파일 | 라인 수 | 복잡도 | 비고 |
|------|--------|--------|------|
| merger.py | 903 | 높음 | 메인 로직, 잦은 수정 |
| row_builder.py | 747 | 높음 | 복잡한 헤더 처리 |
| gstub_cell_splitter.py | 343 | 중간 | _shift_rows_down 주의 |
| parser.py | 295 | 중간 | 안정적, 중첩 테이블 처리 |
| models.py | 152 | 낮음 | 안정적 |

---

## 사용 예시

### 기본 병합

```python
from merge.table import TableParser, TableMerger

# 테이블 파싱
parser = TableParser(auto_field_names=True)
tables = parser.parse_tables("report.hwpx")

# 데이터 병합
merger = TableMerger()
merger.load_base_table("template.hwpx")

data = [
    {"gstub_category": "분류1", "input_value": "100"},
    {"gstub_category": "분류1", "input_value": "200"},
    {"gstub_category": "분류2", "input_value": "300"},
]

merger.merge_with_stub(data, fill_empty_first=True)
merger.save("result.hwpx")
```

### add_ 필드 글머리 포맷팅

```python
from merge.table import TableMerger

merger = TableMerger(
    format_add_content=True,
    use_sdk_for_levels=True
)
merger.load_base_table("template.hwpx")

# add_ 필드에 글머리가 있는 텍스트
data = [
    {
        "add_notes": " 1 첫번째 항목\n   1.1 세부항목\n    1.1.1 상세"
    }
]

merger.merge_with_stub(data)
merger.save("result.hwpx")

# 결과: SDK가 글머리 제거 후 레벨 분석 → □, ○, - 스타일로 변환
# " □ 첫번째 항목"
# "   ○세부항목"
# "    - 상세"
```

### SDK 없이 정규식만 사용

```python
merger = TableMerger(
    format_add_content=True,
    use_sdk_for_levels=False  # 정규식만 사용
)
```
