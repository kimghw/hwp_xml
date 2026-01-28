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

# Workflow 4: 통합 Excel 생성

`workflow4_integrated.py` - workflow1 + workflow2 + workflow3 조합

## 실행 방법

\`\`\`bash
# Windows에서 직접 실행
python workflow4_integrated.py [파일경로]

# WSL에서 실행
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow4_integrated.py" 2>&1
\`\`\`

## 프로세스

1. HWP 열기 - 기존 인스턴스 연결 또는 새로 생성
2. HWPX 임시 변환
3. workflow1 실행 → `{파일}_meta.yaml`
4. workflow2 실행 → `{파일}_para.yaml`
5. workflow3 실행 → `{파일}.xlsx`
6. 임시 파일 삭제
7. 종료 확인 (새로 생성한 경우만)

## 출력 파일

`data/파일명/` 폴더에 저장:
- `{파일명}_meta.yaml` : 테이블/셀 메타데이터 (list_id, 병합정보 등)
- `{파일명}_para.yaml` : 문단 스타일 정보 (폰트, 정렬, 줄간격 등)
- `{파일명}.xlsx` : Excel 변환 결과 (모든 테이블)

---

# Workflow 5: 북마크별 Excel 생성

`workflow5_integrated.py` - workflow1 + workflow2 + 북마크별 시트 분리

## 실행 방법

\`\`\`bash
# Windows에서 직접 실행
python workflow5_integrated.py [파일경로]

# WSL에서 실행
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow5_integrated.py" 2>&1
\`\`\`

## 프로세스

1. HWP 열기 - 기존 인스턴스 연결 또는 새로 생성
2. HWPX 임시 변환
3. workflow1 실행 → `{파일}_meta.yaml`
4. workflow2 실행 → `{파일}_para.yaml`
5. 북마크별 시트 분리 + 문단 분할 → `{파일}_by_bookmark.xlsx`
6. 임시 파일 삭제
7. 종료 확인 (새로 생성한 경우만)

## 출력 파일

`data/파일명/` 폴더에 저장:
- `{파일명}_meta.yaml` : 테이블/셀 메타데이터 (list_id, 병합정보 등)
- `{파일명}_para.yaml` : 문단 스타일 정보 (폰트, 정렬, 줄간격 등)
- `{파일명}_by_bookmark.xlsx` : 북마크별 시트 분리 Excel

---

# Workflow 6: 색상 기반 셀 필드 자동 설정

`insert_field.py` - 색상 기반으로 테이블 셀에 필드명 자동 설정

## 실행 방법

\`\`\`bash
# 파일 선택 대화상자
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_field.py" 2>&1

# 직접 경로 지정
python insert_field.py <입력.hwp>
python insert_field.py <입력_insert_field.hwpx>
\`\`\`

## 프로세스

### HWP 파일 선택 시 (자동 진행)
1. HWP → `_insert_field.hwpx` 변환
2. 한글에서 파일 열림
3. 사용자가 색상 블럭 작업 후 저장
4. **문서 닫으면 자동 감지** (`XHwpDocuments.Count` 폴링)
5. tc.name에 필드명 설정
6. YAML 출력

### HWPX 파일 선택 시
1. tc.name에 필드명 설정
2. YAML 출력

## 출력 파일

- `{원본}_insert_field.hwpx` : 필드 설정된 HWPX (원본 폴더)
- `data/{원본}/{원본}_field.yaml` : 필드 설정 정보

## 색상별 처리 규칙

| 색상 | 조건 | 필드명 형식 |
|------|------|-------------|
| 노란색 | 텍스트 있는 셀 | 셀 텍스트 (20자 제한) |
| 빨간색 | 빈 셀만 | `[L:왼쪽][T:위쪽]` |

## 빨간색 셀 탐색 규칙

- 왼쪽/위쪽으로 최대 3개 텍스트 수집
- 빨간색 범위 내에서만 탐색 (경계 벗어나면 중단)
- 병합 셀은 시작 위치로 점프하여 중복 방지
- `L:` 접두사 = 왼쪽, `T:` 접두사 = 위쪽

## YAML 출력 형식

\`\`\`yaml
- table_idx: 1
  table_id: '2047609131'
  row: 26
  col: 3
  field_name: '[L:연구시설·장비][T:구입기관]'
  type: red
\`\`\`

---

# 공통: 출력 폴더 구조

모든 workflow (4, 5, 6)는 `data/파일명/` 폴더에 출력 파일을 저장합니다.

\`\`\`
입력파일.hwp
data/
  └── 입력파일/
      ├── 입력파일_meta.yaml
      ├── 입력파일_para.yaml
      ├── 입력파일.xlsx
      ├── 입력파일_by_bookmark.xlsx
      └── 입력파일_field.yaml
\`\`\`
