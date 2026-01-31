# Workflow 정리

## 개요

| Workflow | 파일 | 설명 |
|----------|------|------|
| 5 | workflow5_integrated.py | 북마크별 Excel + 메타데이터 |
| 6 | insert_field.py | 빨간/노란 셀에 필드 자동 설정 |

---

## Workflow 5

**입력**: HWP/HWPX
**출력**: data/파일명/

### 프로세스
1. 북마크 확인
2. 기존 필드 추출 → _field.yaml
3. HWPX 변환
4. 메타데이터 삽입 + 캡션 삽입 → _meta.yaml
5. 북마크별 시트 분리 → _by_bookmark.xlsx
6. Excel에 meta 시트 추가 (_meta.yaml + _field.yaml 통합)

### 출력 파일
- 파일명_meta.yaml : 테이블/셀 메타데이터
- 파일명_field.yaml : 기존 필드 이름 (있는 경우)
- 파일명_by_bookmark.xlsx : 북마크별 시트 + meta 시트

### meta 시트 컬럼
tbl_idx, table_id, type, size, row, col, row_span, col_span, list_id, field_name, field_type

---

## Workflow 6

**입력**: HWP/HWPX
**출력**: 수정된 HWPX, data/파일명/파일명_field.yaml

### 프로세스
1. HWPX 변환
2. 빨간/노란 배경 빈 셀 탐색
3. 필드명 생성: [L:왼쪽텍스트][T:위쪽텍스트]
4. tc.name에 필드명 설정
5. _field.yaml 저장

---

## 유틸리티

### field_clear()
모든 tc.name 삭제
from win32.insert_field import field_clear
field_clear("파일.hwp")

### 보안 팝업 방지
from hwp_file_manager import open_hwp
open_hwp(hwp, filepath)
