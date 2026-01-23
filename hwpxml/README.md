# HWPX XML 처리 모듈

## extract_cell_index.py

HWPX 파일에서 `[index:##{숫자}]` 패턴을 추출하여 셀 좌표와 매핑합니다.

### 사용법

```python
from extract_cell_index import extract_indexes_from_hwpx, get_index_mapping

# 방법 1: JSON 저장 안함 (기본)
result = extract_indexes_from_hwpx("table.hwpx")

# 방법 2: JSON 저장 (자동 경로: table_index.json)
result = extract_indexes_from_hwpx("table.hwpx", save_json=True)

# 방법 3: JSON 저장 (경로 지정)
result = extract_indexes_from_hwpx("table.hwpx", save_json=True, json_path="output.json")

# 방법 4: 딕셔너리로 바로 받기
mapping = get_index_mapping("table.hwpx", save_json=True)
```

### JSON 형식

```json
{
  "테이블ID": {
    "인덱스번호": {"row": 행, "col": 열}
  }
}
```

## get_cell_detail.py

HWPX 파일에서 셀의 상세 스타일 정보를 추출합니다.

### 사용법

```python
from get_cell_detail import CellDetail

cd = CellDetail("document.hwpx")
details = cd.get_cell_details(table_id="table_0")

for cell in details:
    print(cell.font, cell.alignment, cell.border, cell.background)
```

### 반환 정보

- **font**: 폰트 정보 (이름, 크기, 굵기 등)
- **alignment**: 정렬 (가로, 세로)
- **border**: 테두리 스타일
- **background**: 배경색

## get_table_property.py

HWPX 파일에서 테이블 및 셀 속성을 추출합니다.

### 사용법

```python
from get_table_property import GetTableProperty

parser = GetTableProperty()
tables = parser.from_hwpx("document.hwpx")

for table in tables:
    print(table.id, table.row_count, table.col_count)
    for row in table.cells:
        for cell in row:
            print(cell.to_dict())
```

## get_page_property.py

HWPX 파일에서 페이지 속성(크기, 여백)을 추출합니다.

### 사용법

```python
from get_page_property import get_page_property, get_all_page_properties, Unit

# 첫 번째 섹션의 페이지 속성
page = get_page_property("document.hwpx")
print(page.page_size.width, page.page_size.height)
print(page.margin.left, page.margin.right)
print(page.content_width, page.content_height)

# 모든 섹션의 페이지 속성
pages = get_all_page_properties("document.hwpx")
for page in pages:
    print(page.to_dict())

# 단위 변환
print(Unit.hwpunit_to_mm(59528))  # HWPUNIT → mm
print(Unit.mm_to_hwpunit(210))    # mm → HWPUNIT
```

### 반환 정보

- **page_size**: 페이지 크기 (width, height, orientation)
- **margin**: 여백 (left, right, top, bottom, header, footer, gutter)
- **content_width/height**: 본문 영역 크기 (페이지 - 여백)
