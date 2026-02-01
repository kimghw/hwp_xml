# excel

HWPX 파일의 테이블을 Excel로 변환하는 모듈

## 파일 목록

| 파일 | 설명 |
|------|------|
| `hwpx_to_excel.py` | HWPX 테이블 → Excel 변환 핵심 모듈 |
| `cell_info_sheet.py` | 셀 상세 정보(위치, 크기, 글꼴 등)를 별도 시트로 저장 |
| `styles.py` | Excel 셀 스타일 적용 (테두리, 배경색, 폰트) |
| `table_placement.py` | Excel 테이블 배치 (열 너비, 행 높이, 페이지 설정) |
| `nested_table.py` | 중첩 테이블 처리 (계층 구조, 셀 위치 매핑) |
| `bookmark.py` | 북마크 기반 테이블 추출 |

## 사용법

```python
from excel import HwpxToExcel, convert_hwpx_to_excel

# 방법 1: 클래스 사용
converter = HwpxToExcel()
converter.convert("input.hwpx", "output.xlsx")

# 방법 2: 함수 사용
convert_hwpx_to_excel("input.hwpx", "output.xlsx")

# 북마크 기반 추출
bookmarks = converter.get_bookmarks("input.hwpx")
converter.convert_by_bookmark("input.hwpx", "북마크명")
```
