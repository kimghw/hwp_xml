# Workflow 1: HWPX → Excel 변환

## HwpxToExcel 변환 메서드 (`excel/hwpx_to_excel.py`)

| 메서드 | 설명 |
|--------|------|
| `convert()` | 1개 테이블 → 1개 시트 |
| `convert_all()` | N개 테이블 → N개 시트 (테이블당 1시트) |
| `convert_all_to_single_sheet()` | N개 테이블 → 1개 시트 (통합) |
| `convert_by_bookmark()` | 특정 북마크 테이블만 → 1개 시트 |
| `convert_by_bookmark_with_body()` | 특정 북마크 본문+테이블 → 1개 시트 |
| `convert_all_by_bookmark()` | 전체 문서 → 북마크별 시트 |

### 옵션
- `split_by_para=True`: 셀 내 문단별 행 분할
- `include_body=True`: 본문 문단 포함 (북마크 변환 시)

### 유틸리티
- `get_bookmarks()`: 북마크 목록
- `get_bookmark_table_mapping()`: 북마크별 테이블 인덱스
- `get_body_elements()`: 본문 요소 (문단/테이블) 순서대로
- `get_bookmark_body_mapping()`: 북마크별 본문 요소

---

# Workflow 1-Legacy: 테이블 메타데이터 추출

## 필드 → list_id 연결 원리

HWPX의 tc.name 속성을 임시 저장소로 활용하여 셀 위치(row/col)와 list_id를 매핑:

1. **필드 삽입** (`insert_table_field.py`): 각 셀에 JSON 필드명 설정
   ```python
   # tc.name에 저장되는 JSON
   {"tblIdx": 0, "rowAddr": 0, "colAddr": 0, "type": "parent"}
   ```

2. **list_id 추출** (`extract_cell_meta.py`): COM API로 셀 순회하며 매핑
   ```python
   field_name = hwp.GetCurFieldName(0)  # JSON 필드명 읽기
   fd = json.loads(field_name)
   row, col = fd['rowAddr'], fd['colAddr']
   list_id = hwp.GetPos()[0]  # 현재 위치의 list_id
   cells[(row, col)] = (list_id, para_id)
   ```

3. **필드 삭제**: 작업 완료 후 tc.name 속성 제거하여 원본 복원

## 전체 프로세스

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

# Workflow 4: 통합 Excel 생성

`workflow4_integrated.py` - workflow1 + workflow2 + workflow3 조합

## 실행 방법

```bash
# Windows에서 직접 실행
python workflow4_integrated.py [파일경로]

# WSL에서 실행
cmd.exe /c "cd /d C:\hwp_xml\win32 && python workflow4_integrated.py" 2>&1
```

## 프로세스

1. HWP 열기 - 기존 인스턴스 연결 또는 새로 생성
2. HWPX 임시 변환
3. workflow1 실행 → `{파일}_meta.yaml`
4. workflow2 실행 → `{파일}_para.yaml`
5. workflow3 실행 → `{파일}.xlsx`
6. 임시 파일 삭제
7. 종료 확인 (새로 생성한 경우만)

## 출력 파일

- `{base}_meta.yaml` : 테이블/셀 메타데이터 (list_id, 병합정보 등)
- `{base}_para.yaml` : 문단 스타일 정보 (폰트, 정렬, 줄간격 등)
- `{base}.xlsx` : Excel 변환 결과 (모든 테이블)

---

# Workflow 5: 북마크별 Excel 생성

`workflow5_integrated.py` - workflow1 + workflow2 + 북마크별 시트 분리

## 실행 방법

```bash
# Windows에서 직접 실행
python workflow5_integrated.py [파일경로]

# WSL에서 실행
cmd.exe /c "cd /d C:\hwp_xml\win32 && python workflow5_integrated.py" 2>&1
```

## 프로세스

1. HWP 열기 - 기존 인스턴스 연결 또는 새로 생성
2. HWPX 임시 변환
3. workflow1 실행 → `{파일}_meta.yaml`
4. workflow2 실행 → `{파일}_para.yaml`
5. 북마크별 시트 분리 + 문단 분할 → `{파일}_by_bookmark.xlsx`
6. 임시 파일 삭제
7. 종료 확인 (새로 생성한 경우만)

## 출력 파일

- `{base}_meta.yaml` : 테이블/셀 메타데이터 (list_id, 병합정보 등)
- `{base}_para.yaml` : 문단 스타일 정보 (폰트, 정렬, 줄간격 등)
- `{base}_by_bookmark.xlsx` : 북마크별 시트 분리 Excel

## Workflow 4 vs 5 차이점

| 항목 | Workflow 4 | Workflow 5 |
|------|------------|------------|
| 시트 구조 | 테이블당 1시트 (tbl_0, tbl_1...) | 북마크당 1시트 |
| 본문 포함 | X | O (include_body=True) |
| 문단 분할 | O | O |
| 출력 파일 | `{base}.xlsx` | `{base}_by_bookmark.xlsx` |

---

# Workflow 6: 빨간색 셀 필드 자동 설정

`set_red_field.py` - 빨간색 배경 빈 셀에 필드명 자동 설정

## 실행 방법

```bash
python set_red_field.py <입력.hwp> <출력.hwp>
```

## 프로세스

1. HWP → HWPX 임시 변환
2. 빨간색 배경(`#FF0000`, `#CF2741`) 빈 셀 찾기
3. 왼쪽/위쪽 셀 텍스트로 필드명 생성: `[왼쪽텍스트][위쪽텍스트]`
4. HWPX XML의 `tc.name` 속성에 필드명 설정
5. HWP로 저장

## 특징

- 병합 셀 지원: 왼쪽/위쪽 검색 시 병합된 부모 셀 텍스트 참조
- 빈 셀만 처리: 텍스트가 있는 셀은 스킵
