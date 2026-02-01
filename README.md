# hwp_xml

HWPX 파일 파싱, Excel 변환, 데이터 병합 도구

## 프로젝트 개요

- HWPX(한글 XML) 파일 파싱 및 속성 추출
- HWPX → Excel 변환 (테이블, 북마크 기반)
- Excel/JSON 데이터를 HWPX 템플릿에 병합
- Windows 한글 COM API 연동 (WSL 지원)

## 폴더 구조

```
hwp_xml/
├── core/           # 공통 유틸리티 (단위 변환, 파일 대화상자)
├── hwpxml/         # HWPX 파싱 (테이블/셀/페이지 속성 추출)
├── excel/          # HWPX → Excel 변환
├── merge/          # 데이터 병합 (HWPX + Excel/JSON)
│   ├── field/      # 필드명 자동 삽입/관리
│   ├── formatters/ # 콘텐츠 포맷터 (캡션, 불릿 등)
│   └── table/      # 테이블 병합 로직
├── win32/          # Windows 한글 COM API
├── workflow/       # 통합 워크플로우
├── agent/          # AI 에이전트 포맷터
├── reference/      # 참고 문서
└── data/           # 테스트 데이터
```

### core/
- `unit.py`: HWPUNIT 단위 변환
- `file_dialog.py`: Windows 파일 탐색기 대화상자

### hwpxml/
- `get_table_property.py`: 테이블/셀 속성 추출
- `get_cell_detail.py`: 셀 스타일, 중첩 테이블 파싱
- `set_field_by_header.py`: 헤더 기준 필드명 설정

### excel/
- `hwpx_to_excel.py`: HWPX → Excel 변환
- `bookmark.py`: 북마크 기반 변환
- `nested_table.py`: 중첩 테이블 처리

### merge/
- `merge_hwpx.py`: HWPX 병합 메인 로직
- `run_merge.py`: 병합 실행 스크립트
- `field/`: 필드명 자동 삽입, 빈 필드 채우기
- `formatters/`: 캡션/불릿 포맷터
- `table/`: 테이블 행 병합, 셀 분할

### win32/
- `hwp_file_manager.py`: COM API 유틸리티
- `convert_hwp.py`: HWP↔HWPX 변환
- `insert_table_field.py`: 테이블 셀 필드명 설정
- `extract_field.py`: 필드명 추출/제거

### workflow/
- `workflow5_integrated.py`: 북마크 기반 HWP→Excel 파이프라인

### agent/
- `bullet_formatter.py`: AI 기반 불릿 포맷터
- `caption_formatter.py`: AI 기반 캡션 포맷터

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
