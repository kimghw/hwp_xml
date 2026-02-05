# merge/field - 테이블 필드 관리 모듈

HWPX 테이블 셀에 필드명 자동 생성/삽입/시각화

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

---

## 빈 셀 채우기 (EmptyFieldFiller) 처리 케이스

### Case 1: 위 행 복사
- 필드명 없는 행이 모두 비어있음
- 위 행이 모두 `data_`, `input_`
- 컬럼 수/너비 동일 → 위 셀의 필드명 복사

### Case 2: gstub 처리 (rowspan > 1)
- `gstub_{랜덤}` 할당
- 우측 빈 셀: 첫 행은 `input_{랜덤}`, 이후 행은 첫 행 복사

### Case 3: 연속 gstub 처리
- gstub 2개 이상이 동일 행 범위 → 같은 nc_name 공유
- gstub rowspan 일치 조건 필수

---

## 최근 변경사항 (3일 이내)

### ee3da26: add_ 필드 추출 버그 수정
- header row(row 0)에서 add_ 필드가 추출되지 않는 문제 해결
- `_extract_addition_table_data`에 add_ prefix 필터 조건 추가
- 중복 bullet formatting 방지
- 다중 문단 텍스트 지원

### d9eb027: 파일 재조직
- `auto_field.py` + `field_name_generator.py` → `auto_insert_field_template.py` 통합
- `visualizer.py` → `check_empty_field.py` 이름 변경
- fill_empty.py Case 1, 2, 3 gstub 처리 로직 완성
- 순환 참조 해결 (lazy import 적용)

---

## 개발자 주의사항

1. **로거**: `field_name_log.txt`에 상세 로그 기록
2. **HWPX 처리**: ZIP 기반 → 임시 해제 후 수정 → 재압축
3. **네임스페이스**: XML 네임스페이스 등록 필수 (`ET.register_namespace`)
4. **borderFill ID 관리**: 기존 ID와 충돌 방지 필요
5. **gstub rowspan**: 매칭 조건이 엄격함 (Case 3)
6. **순환 참조**: fill_empty.py에서 auto_insert_field_template import 시 주의
