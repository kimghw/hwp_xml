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
- `get_cell_detail.py`: 셀 스타일 정보, 중첩 테이블 분리 파싱
- `extract_cell_index.py`: `[index:##숫자]` 패턴 매핑
- `get_page_property.py`: 페이지 속성, 단위 변환
- `set_field_by_header.py`: 필드명을 헤더 기준으로 설정

### core/
공통 유틸리티 모듈
- `unit.py`: HWPUNIT 단위 변환
- `file_dialog.py`: Windows 파일 탐색기 대화상자 (WSL↔Windows 경로 변환)

### win32/
Windows 한글 COM API 연동 (WSL에서 cmd.exe로 실행)
- `hwp_file_manager.py`: 공통 유틸리티 (인스턴스 연결, 파일 대화상자, 저장)
- `convert_hwp.py`: HWP↔HWPX 변환
- `insert_ctrl_id.py`: 테이블에 ctrl_id 메타데이터 삽입
- `insert_table_field.py`: 테이블 셀에 필드명 설정
- `insert_field_within_redcells.py`: 빨간색 배경 빈 셀에 필드명 자동 설정
- `get_para_style.py`: 문단 스타일 추출 → YAML
- `get_table_property.py`: COM API로 테이블 속성 추출
- `extract_field.py`: 기존 필드명 추출/제거

### workflow/
통합 워크플로우
- `workflow5_integrated.py`: 북마크 기반 HWP→Excel 변환 파이프라인

## WSL에서 Windows Python 실행

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_ctrl_id.py" 2>&1
```

## 한글 COM API 유틸리티

```python
from hwp_file_manager import get_hwp_instance, create_hwp_instance, save_hwp

hwp = get_hwp_instance()      # 열린 한글 연결
hwp = create_hwp_instance()   # 새 한글 생성 (SecurityModule 포함)
save_hwp(hwp, filepath, "HWP")
```
