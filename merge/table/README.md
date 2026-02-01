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

Base 파일에 데이터 병합

**필드명 접두사별 동작:**

| 접두사 | 동작 |
|--------|------|
| `header_` | 유지 |
| `data_` | 유지 |
| `add_` | 기존 셀에 내용 추가 |
| `stub_` | 새 행 생성 |
| `gstub_` | 같은 값이면 rowspan 확장 |
| `input_` | 빈 셀 채우기, 부족시 행 추가 |

```python
from merge.table import TableMerger

merger = TableMerger()
merger.load_base_table("base.hwpx", table_index=0)
merger.merge_with_stub(data_list)
merger.save("output.hwpx")
```

### cell_splitter.py

gstub 범위 내 행 삽입 및 rowspan 확장

- `GstubCellSplitter`: gstub 셀 나누기 처리

### row_builder.py

헤더 기반 자동 행 추가

- `RowBuilder.add_rows_smart()`: 필드명 접두사 분석하여 자동 행 추가
- `RowBuilder.add_rows_auto()`: 헤더 이름 기준 자동 행 추가
- `RowBuilder.add_row_with_headers()`: 헤더 설정에 따른 행 추가

### formatter_config.py

add_ 필드용 포맷터 설정 로드 (YAML)

| 클래스/함수 | 설명 |
|-------------|------|
| `TableFormatterConfigLoader` | YAML 설정 파일 로더 |
| `TableFormatterConfig` | 테이블 포맷터 통합 설정 |
| `FieldFormatterConfig` | 필드별 포맷터 설정 |
| `load_table_formatter_config()` | 설정 로드 편의 함수 |
| `format_add_field_value()` | add_ 필드값에 포맷터 적용 |

## 사용 예시

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
