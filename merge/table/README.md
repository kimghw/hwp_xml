# merge/table 모듈

HWPX 테이블 파싱, 필드명 자동 생성, 데이터 병합을 위한 모듈입니다.

## 파일 구조

```
merge/table/
├── __init__.py           # 모듈 초기화 및 exports
├── models.py             # 데이터 모델 (CellInfo, TableInfo 등)
├── parser.py             # HWPX 테이블 파싱
├── merger.py             # 테이블 셀 내용 병합
├── field_name_generator.py  # 필드명 자동 생성
├── insert_auto_field.py  # 자동 필드명을 tc 태그 속성에 삽입
└── insert_field_text.py  # 필드명을 셀 내 텍스트로 삽입 (테스트용)
```

## 모듈별 역할

### models.py

테이블 관련 데이터 모델을 정의합니다.

| 클래스 | 설명 |
|--------|------|
| `CellInfo` | 테이블 셀 정보 (위치, 크기, 텍스트, 필드명 등) |
| `TableInfo` | 테이블 정보 (행/열 수, 셀 딕셔너리, 필드 매핑) |
| `HeaderConfig` | 새 행 추가 시 헤더 열 설정 |
| `RowAddPlan` | 행 추가 계획 |

### parser.py

HWPX 파일에서 테이블을 파싱합니다.

```python
from merge.table import TableParser

parser = TableParser(auto_field_names=True)
tables = parser.parse_tables("input.hwpx")

for table in tables:
    print(f"테이블: {table.row_count}행 x {table.col_count}열")
    for (row, col), cell in table.cells.items():
        print(f"  ({row},{col}): {cell.text[:20]}, 필드={cell.field_name}")
```

**주요 옵션:**
- `auto_field_names=True`: 필드명 없는 셀에 자동 생성
- `regenerate=True`: 기존 필드명 무시하고 새로 생성

### merger.py

Base 파일에 Add 데이터를 병합합니다.

**필드명 접두사별 동작:**

| 접두사 | 동작 |
|--------|------|
| `header_` | 유지 (변경 없음) |
| `data_` | 유지 (변경 없음) |
| `add_` | 기존 셀에 내용 추가 (행 추가 없음) |
| `stub_` | 새 행 생성 |
| `gstub_` | 같은 값이면 rowspan 확장, 다른 값이면 새 셀 |
| `input_` | 빈 셀 채우기, 없으면 새 행 추가 |

```python
from merge.table import TableMerger

merger = TableMerger()
merger.load_base_table("base.hwpx", table_index=0)

add_data = [
    {"gstub_abc": "그룹A", "input_123": "값1"},
    {"gstub_abc": "그룹A", "input_123": "값2"},  # gstub rowspan 확장
    {"gstub_abc": "그룹B", "input_123": "값3"},  # 새 gstub 셀
]
merger.merge_with_stub(add_data)
merger.save("output.hwpx")
```

### field_name_generator.py

테이블 셀에 자동으로 필드명(nc_name)을 생성합니다.

**접두사 규칙:**

| 접두사 | 조건 |
|--------|------|
| `header_` | 행 전체에 배경색 있음 |
| `add_` | 최상단 행 + 텍스트 30자 이상 + 배경색 없음, 또는 1x1 단일 셀 |
| `stub_` | 텍스트 있음 + 오른쪽에 빈 셀 + rowspan 없음 |
| `gstub_` | 텍스트 있음 + 오른쪽에 빈 셀 + rowspan 있음 |
| `input_` | 빈 셀 (기본값) |

```python
from merge.table import FieldNameGenerator, CellForNaming

cells = [
    CellForNaming(row=0, col=0, text='헤더', bg_color='#CCCCCC'),
    CellForNaming(row=1, col=0, text='분류A', row_span=2),
    CellForNaming(row=1, col=1, text=''),  # 빈 셀
]

generator = FieldNameGenerator()
generator.generate(cells)

for cell in cells:
    print(f"({cell.row},{cell.col}): {cell.nc_name}")
```

### insert_auto_field.py

자동 생성된 필드명을 HWPX 파일의 `tc` 태그 `name` 속성에 삽입합니다.

```python
from merge.table.insert_auto_field import insert_auto_fields

tables = insert_auto_fields("input.hwpx", "output.hwpx")
# 또는 원본 덮어쓰기
tables = insert_auto_fields("input.hwpx")
```

**CLI 사용:**
```bash
python -m merge.table.insert_auto_field input.hwpx output.hwpx
python -m merge.table.insert_auto_field input.hwpx --regenerate  # 기존 필드명 재생성
```

### insert_field_text.py

필드명을 셀 내 파란색 텍스트로 삽입합니다 (테스트/확인용).

```python
from merge.table.insert_field_text import insert_field_text

tables = insert_field_text("input.hwpx", "output_with_field.hwpx")
```

## 중첩 테이블 처리

HWPX에서 테이블 셀 안에 또 다른 테이블이 포함될 수 있습니다.

### 문제점

`root.iter('tc')`로 모든 셀을 가져오면 부모/자식 테이블 셀이 문서 순서로 섞여서 데이터가 어긋납니다.

### 해결 방법

`TableParser`는 재귀적으로 테이블을 탐색하여 각 테이블의 직접 셀만 파싱합니다:

1. `_find_tables_recursive()`: 문서 순서대로 테이블 탐색
2. `_parse_table()`: 테이블의 직접 자식 셀만 파싱
3. 중첩 테이블은 별도 `TableInfo`로 분리

```python
parser = TableParser()
tables = parser.parse_tables("nested_tables.hwpx")

# tables[0]: 부모 테이블
# tables[1]: 첫 번째 중첩 테이블
# tables[2]: 두 번째 중첩 테이블 (문서 순서)
```

## 사용 예시

### 테이블 파싱 및 필드명 자동 생성

```python
from merge.table import TableParser

parser = TableParser(auto_field_names=True)
tables = parser.parse_tables("report.hwpx")

for table in tables:
    print(f"테이블: {table.row_count}x{table.col_count}")
    for field_name, (row, col) in table.field_to_cell.items():
        print(f"  {field_name} -> ({row}, {col})")
```

### 데이터 병합

```python
from merge.table import TableMerger

merger = TableMerger()
merger.load_base_table("template.hwpx")

# 데이터 추가
data = [
    {"header_category": "분류1", "input_value": "100", "add_memo": "비고1"},
    {"header_category": "분류1", "input_value": "200", "add_memo": "비고2"},
    {"header_category": "분류2", "input_value": "300", "add_memo": "비고3"},
]

merger.merge_with_stub(data, fill_empty_first=True)
merger.save("result.hwpx")
```
