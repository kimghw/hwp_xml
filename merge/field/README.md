# merge/field - 테이블 필드 관리 모듈

## 기능

| 모듈 | 기능 |
|------|------|
| `auto_field.py` | 테이블 구조 분석 → 필드명 자동 생성 |
| `fill_empty.py` | 빈 셀에 위 셀 필드명 복사 |
| `visualizer.py` | 빨간 배경 강조 / 파란 텍스트 표시 |

## 사용법

```python
from merge.field import (
    insert_auto_fields,      # 자동 필드명 생성
    fill_empty_fields,       # 빈 셀에 위 셀 필드명 복사
    highlight_empty_fields,  # 필드 없는 셀 빨간 배경
    insert_field_text,       # 필드명 파란 텍스트 표시
)

# 1. 자동 필드명 생성
insert_auto_fields("template.hwpx")

# 2. 행 추가 후 빈 셀 채우기
fill_empty_fields("template.hwpx")

# 3. 시각화
highlight_empty_fields("template.hwpx", "highlighted.hwpx")
insert_field_text("template.hwpx", "with_field.hwpx")
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

## fill_empty_fields 동작

행 추가 후 필드명이 없는 셀 처리:

```
조건:
1. 행의 모든 셀에 필드명 없음
2. 셀 개수와 너비가 위 행과 동일

→ 위 셀의 필드명 복사
```

## CLI 실행

```bash
# 자동 필드 생성
python3 -m merge.field.auto_field input.hwpx output.hwpx

# 빈 셀 채우기
python3 -m merge.field.fill_empty input.hwpx output.hwpx

# 시각화
python3 -m merge.field.visualizer highlight input.hwpx output.hwpx
python3 -m merge.field.visualizer text input.hwpx output.hwpx
```
