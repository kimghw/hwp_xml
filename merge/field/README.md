# merge/field - 테이블 필드 관리 모듈

## 파일 구조

| 파일 | 기능 |
|------|------|
| `auto_insert_field_template.py` | 필드명 자동 생성 (FieldNameGenerator, AutoFieldInserter) |
| `insert_auto_field.py` | 자동 필드명을 HWPX 셀에 삽입 |
| `fill_empty.py` | 빈 셀에 위 셀 필드명 복사 |
| `check_empty_field.py` | 필드 시각화 (빨간 배경/파란 텍스트) |
| `insert_field_background_color.py` | 필드명별 배경색 설정 |
| `insert_field_text.py` | nc_name을 셀에 파란 텍스트로 삽입 |

## 사용법

```python
from merge.field import (
    insert_auto_fields,      # 자동 필드명 생성
    fill_empty_fields,       # 빈 셀에 위 셀 필드명 복사
    highlight_empty_fields,  # 필드 없는 셀 빨간 배경
    insert_field_text,       # 필드명 파란 텍스트 표시
    colorize_by_field,       # 필드명별 배경색 설정
)
```

## 필드명 접두사 규칙

| 접두사 | 조건 |
|--------|------|
| `header_` | 행 전체가 배경색 있음 |
| `add_` | 최상단 데이터 행 + 30자 이상 텍스트 |
| `stub_` | 텍스트 + 오른쪽 빈 셀 (rowspan=1) |
| `gstub_` | 텍스트 + 오른쪽 빈 셀 (rowspan>1) |
| `input_` | 빈 셀 |
| `data_` | 텍스트 있음 + stub 조건 미충족 |

## 접두사별 배경색 (insert_field_background_color.py)

| 접두사 | 색상 |
|--------|------|
| `header_` | 회색 (#D1D1D1) |
| `add_` | 연한 파랑 (#E6F3FF) |
| `stub_` | 연한 노랑 (#FFFFD0) |
| `gstub_` | 연한 주황 (#FFE4C4) |
| `input_` | 연한 초록 (#E8FFE8) |
| `data_` | 연한 보라 (#F0E6FF) |
