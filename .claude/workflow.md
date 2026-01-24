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
