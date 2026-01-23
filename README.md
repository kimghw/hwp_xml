# hwp_xml

HWPX 파일 파싱 및 Excel 변환 도구

## 폴더 구조

### excel/
HWPX → Excel 변환 모듈
- `hwpx_to_excel.py`: 테이블 변환 (HWPUNIT 단위 변환)
- `cell_info_sheet.py`: 셀 상세 정보 시트 생성

### hwpxml/
HWPX 파일 파싱 모듈
- `get_table_property.py`: 테이블/셀 속성 추출
- `get_cell_detail.py`: 셀 스타일 정보
- `extract_cell_index.py`: `[index:##숫자]` 패턴 매핑
- `get_page_property.py`: 페이지 속성, 단위 변환

### win32/
Windows 한글 COM API 연동 (WSL에서 cmd.exe로 실행)
- `insert_ctrl_id.py`: 테이블에 ctrl_id 메타데이터 삽입
- `get_table_property.py`: COM API로 테이블 속성 추출
- `get_table_info.py`: 테이블 위치/구조 조회

## WSL에서 Windows Python 실행

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_ctrl_id.py" 2>&1
```
