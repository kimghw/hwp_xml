# HWPX XML 처리 모듈

HWPX 파일에서 테이블, 페이지, 셀 속성을 XML 파싱으로 추출

## 파일 목록

| 파일 | 설명 |
|------|------|
| `get_table_property.py` | 테이블/셀 속성 추출 (위치, 병합, 크기) |
| `get_cell_detail.py` | 셀 상세 스타일 (폰트, 정렬, 테두리, 배경색) |
| `get_page_property.py` | 페이지 속성 (크기, 여백), Unit 단위 변환 |
| `extract_cell_index.py` | `[index:##{숫자}]` 패턴 추출/매핑 |
| `export_meta_yaml.py` | 셀 메타데이터 YAML 추출 (list_id, para_id, field_name) |
| `get_para_property.py` | 문단 속성 추출 (미완성) |
| `set_field_by_header.py` | 헤더 셀 기준 필드명 자동 설정 |

## 주요 사용법

### 테이블 속성 추출
```python
from hwpxml import GetTableProperty

parser = GetTableProperty()
tables = parser.from_hwpx("document.hwpx")
for table in tables:
    for row in table.cells:
        for cell in row:
            print(cell.row_index, cell.col_index, cell.text)
```

### 셀 상세 정보 (중첩 테이블 분리)
```python
from hwpxml import GetCellDetail

parser = GetCellDetail()
table_cells = parser.from_hwpx_by_table("document.hwpx")
# table_cells[0]: 첫 번째 테이블 셀들
# table_cells[1]: 두 번째 테이블 (중첩 포함)
```

### 페이지 속성/단위 변환
```python
from hwpxml import get_page_property, Unit

page = get_page_property("document.hwpx")
print(page.content_width, page.content_height)
print(Unit.hwpunit_to_mm(59528))  # HWPUNIT -> mm
```

### 인덱스 패턴 추출
```python
from hwpxml import get_index_mapping

mapping = get_index_mapping("table.hwpx")
# {"테이블ID": {"인덱스번호": {"row": 행, "col": 열}}}
```
