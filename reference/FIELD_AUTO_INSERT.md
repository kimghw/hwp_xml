# 필드 자동 삽입 가이드

## 개요

HWP 테이블 셀에 필드명(메타데이터)을 자동 삽입하는 방식들.

| 방식 | 위치 | 용도 |
|------|------|------|
| **구조 기반 자동 생성** | `merge/table/` | 테이블 구조 분석 → 필드명 자동 생성 |
| JSON 메타데이터 | `win32/insert_table_field.py` | 테이블 구조 전체 매핑 |
| 색상 기반 | `win32/insert_field.py` | 사용자 지정 필드명 |
| list_id 텍스트 | `win32/insert_listid_on_hwp.py` | 셀 위치 추적용 |

---

## 1. merge/table/ - 구조 기반 자동 생성 (권장)

### 핵심 모듈

| 파일 | 역할 |
|------|------|
| `field_name_generator.py` | 테이블 구조 분석 → 필드명 자동 생성 |
| `insert_auto_field.py` | 생성된 필드명을 HWPX에 저장 |
| `insert_field_text.py` | 필드명을 파란색 텍스트로 표시 (확인용) |
| `parser.py` | HWPX 테이블 파싱 |
| `merger.py` | 데이터 병합 엔진 |

### 필드명 접두사 규칙

| 접두사 | 조건 |
|--------|------|
| `header_` | 행 전체가 배경색 있음 |
| `add_` | 최상단 데이터 행 + 30자 이상 텍스트 |
| `stub_` | 텍스트 있음 + 오른쪽 빈 셀 (rowspan=1) |
| `gstub_` | 텍스트 있음 + 오른쪽 빈 셀 (rowspan>1) |
| `input_` | 빈 셀 (기본값) |
| `data_` | 텍스트 있음 + stub 조건 미충족 |

### 처리 흐름

```
HWPX → TableParser → FieldNameGenerator (4단계 분석)
    ↓
AutoFieldInserter (tc.name 속성에 저장)
    ↓
수정된 HWPX
```

### 사용 예시

```python
# 필드명 자동 생성 + XML 저장
from merge.table.insert_auto_field import insert_auto_fields
insert_auto_fields("template.hwpx")

# 파란색 텍스트로 확인 (선택)
from merge.table.insert_field_text import insert_field_text
insert_field_text("template.hwpx", "template_check.hwpx")
```

### 데이터 병합

```python
from merge.table import TableMerger

merger = TableMerger()
merger.load_base_table("template.hwpx")

data = [
    {"gstub_category": "A", "input_value": "100"},
    {"gstub_category": "A", "input_value": "200"},  # rowspan 확장
    {"gstub_category": "B", "input_value": "300"},  # 새 셀
]

merger.merge_with_stub(data)
merger.save("output.hwpx")
```

---

## 1. insert_table_field.py

### 기능
- 모든 셀에 JSON 필드명 설정
- 캡션 필드 자동 삽입
- Excel/YAML 출력

### 흐름
```
HWP → HWPX → XML 수정 → 캡션 삽입 → HWP + Excel + YAML
```

### JSON 형식
```json
{
  "tblIdx": 0,
  "rowAddr": 1,
  "colAddr": 2,
  "rowSpan": 1,
  "colSpan": 2,
  "type": "parent",
  "parentTbl": null,
  "parentCell": null
}
```

---

## 2. insert_field.py

### 색상 규칙
| 색상 | 조건 | 필드명 |
|------|------|--------|
| 빨간색 | R>180, G<80, B<80 | `[L:좌측][T:상단]` 자동 생성 |
| 노란색 | R>200, G>200, B<100 | 셀 텍스트 사용 |

### 빨간색 셀 필드명 생성
```
← 왼쪽 최대 3개 텍스트
↑ 위쪽 최대 3개 텍스트
→ [L:항목][T:헤더] 형식
```

---

## 3. insert_listid_on_hwp.py

- 열린 한글 문서에 `[list_id:값]` 텍스트 삽입
- COM API 직접 조작 (가장 빠름)

---

## 비교표

| 항목 | insert_table_field | insert_field | insert_listid_on_hwp |
|------|-------------------|--------------|---------------------|
| 입력 | HWP 파일 | HWP/HWPX | 열린 한글 |
| 방식 | HWPX XML | HWPX XML | COM API |
| 색상 필요 | 없음 | 필요 | 없음 |
| 캡션 | 자동 | 없음 | 없음 |
| 출력 | HWP, Excel, YAML | HWPX, YAML | HWP |

---

## 비교: merge/table vs win32

| 항목 | merge/table | win32 |
|------|-------------|-------|
| 필드명 생성 | 자동 (구조 분석) | 수동/색상 기반 |
| 데이터 병합 | 지원 (merger.py) | 미지원 |
| 중첩 테이블 | 재귀 처리 | 지원 |
| 출력 | HWPX | HWP, Excel, YAML |

---

## 실행 (WSL)

```bash
# merge/table - 필드명 자동 생성
python -m merge.table.insert_auto_field input.hwpx output.hwpx

# win32/insert_table_field.py
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_table_field.py" 2>&1

# win32/insert_field.py
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_field.py" 2>&1

# win32/insert_listid_on_hwp.py (한글 열린 상태)
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_listid_on_hwp.py" 2>&1
```
