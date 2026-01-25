# Workflow 1: 테이블 메타데이터 추출

1. HWP 열기 - `insert_table_field.py`
2. HWPX 변환 - `convert_hwp.py` → `temp.hwpx`
3. 필드명 삽입 (tc.name에 JSON) - `insert_table_field.py`
4. 캡션 삽입 - `insert_table_field.py`
5. HWP 저장 - `insert_table_field.py` → `파일.hwp`
6. Excel 변환 - `hwpx_to_excel.py` → `파일.xlsx`
7. COM API로 셀 순회 - `extract_cell_meta.py`
8. GetCurFieldName으로 필드명 읽기 - `extract_cell_meta.py`
9. GetPos로 list_id 추출 - `extract_cell_meta.py`
10. YAML 저장 - `extract_cell_meta.py` → `파일_meta.yaml`
11. 필드명 삭제 (tc.name 속성 제거) - `extract_cell_meta.py`
12. HWP 저장 - `extract_cell_meta.py` → `파일.hwp`

---

# Workflow 2: 문단 스타일 추출

1. 열린 한글 연결 - `get_para_style.py`
2. list_id 순차 조회 (0~1000) - `SetPos(list_id, para_id, 0)`
3. 문단 정보 추출 (text, line_count, char_id, 스타일) - `get_para_style.py`
4. YAML 저장 - `get_para_style.py` → `파일_para.yaml`

---

# Workflow 3: HWPX → Excel 변환 (문단 분할)

1. HWPX 파싱 - `get_cell_detail.py`
2. 테이블별 셀/문단 정보 추출 - `get_cell_detail.py`
3. Excel 변환 (문단별 행 분할) - `hwpx_to_excel.py`
4. 메타데이터 시트 생성 - `cell_info_sheet.py` → `파일.xlsx`

---

# Workflow 4: 통합 Excel 생성 (예정)

workflow1 + workflow2 + workflow3 조합하여 완전한 메타데이터가 포함된 Excel 생성

1. HWP 열기
2. workflow1 실행 → `_meta.yaml`
3. workflow2 실행 → `_para.yaml`
4. workflow3 실행 (YAML 연동) → `파일.xlsx`
5. 정리 (임시 파일 삭제)
