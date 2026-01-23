# excel

HWPX 파일의 테이블을 Excel로 변환하는 모듈

## 주요 파일

| 파일 | 설명 |
|------|------|
| `hwpx_to_excel.py` | HWPX 테이블 → Excel 변환 핵심 모듈 (HWPUNIT 단위 변환) |
| `cell_info_sheet.py` | 셀 상세 정보(위치, 크기, 글꼴, 배경색)를 별도 시트로 저장 |

## 사용법

```python
from excel import HwpxToExcel, convert_hwpx_to_excel

# 방법 1: 클래스 사용
converter = HwpxToExcel("input.hwpx")
converter.convert("output.xlsx")

# 방법 2: 함수 사용
convert_hwpx_to_excel("input.hwpx", "output.xlsx")
```
