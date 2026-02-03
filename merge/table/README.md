# merge/table 모듈

HWPX 테이블 파싱 및 데이터 병합 모듈

## 파일 구조

```
merge/table/
├── __init__.py           # 모듈 초기화 및 exports
├── models.py             # 데이터 모델 (CellInfo, TableInfo 등)
├── parser.py             # HWPX 테이블 파싱
├── merger.py             # 테이블 셀 내용 병합
├── cell_splitter.py      # gstub 셀 나누기
├── row_builder.py        # 테이블 행 자동 생성
└── formatter_config.py   # add_ 필드 포맷터 설정 로더
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

```python
from merge.table import TableMerger

merger = TableMerger(
    format_add_content=True,   # add_ 필드 글머리 포맷팅
    use_sdk_for_levels=True    # SDK로 레벨 분석
)
merger.load_base_table("base.hwpx", table_index=0)
merger.merge_with_stub(data_list)
merger.save("output.hwpx")
```

### cell_splitter.py

gstub 범위 내 행 삽입 및 rowspan 확장

**주요 기능:**
- gstub 범위 내 새 행 삽입
- 기존 행들을 밀어내고 rowspan 확장
- 새 행의 셀 생성 (input_, stub_, data_ 등)
- 다중 gstub 열 지원

```python
# GstubCellSplitter는 TableMerger 내부에서 사용됨
splitter = GstubCellSplitter(table_info)
splitter.insert_row_in_gstub_range(gstub_field, insert_row, data)
```

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

## 병합 흐름

```
1. merge_with_stub(data_list) 호출
   │
2. 각 데이터 행 처리
   ├─ gstub/stub 값으로 매칭되는 빈 행 찾기
   ├─ 빈 행 있음 → input 셀 채우기
   └─ 빈 행 없음 → gstub 범위 내 새 행 삽입
   │
3. add_ 필드 처리 (_process_add_fields)
   ├─ SDK로 글머리 제거 + 레벨 분석 (analyze_and_strip)
   ├─ 정규식으로 새 글머리 적용 (□, ○, - 스타일)
   └─ 기존 셀에 내용 추가 (여러 문단 지원)
   │
4. 결과 저장
```

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
