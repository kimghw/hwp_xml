# win32

Windows 아래아한글 COM API 연동 모듈

## 주요 파일

| 파일 | 설명 |
|------|------|
| `hwp_file_manager.py` | 공통 유틸리티 (인스턴스 연결, 보안 모듈, 파일 열기/저장) |
| `convert_hwp.py` | HWP ↔ HWPX 변환 |
| `get_table_property.py` | COM API로 테이블 속성 추출 |
| `get_table_info.py` | 모든 테이블 위치/구조 정보 조회 |
| `get_para_style.py` | 문단/글자 스타일 추출 → YAML |
| `insert_table_field.py` | 테이블 셀에 필드 메타데이터(tc.name) 삽입 |
| `insert_field.py` | 색상 기반 셀 필드 자동 설정 |
| `insert_listid_on_hwp.py` | 각 셀에 list_id 텍스트 삽입 |
| `extract_field.py` | 기존 필드 추출 및 삭제 → YAML |
| `extract_cell_meta.py` | 테이블 셀 메타데이터(list_id, field_name) 추출 |
| `check_field.py` | HWPX 파일의 tc.name 필드 확인 |

## 실행 방법

Windows 전용 모듈이므로 WSL에서는 cmd.exe로 실행:

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python <파일명>.py" 2>&1
```

## 사용법

```python
from hwp_file_manager import get_hwp_instance, create_hwp_instance

# 열린 한글 연결
hwp = get_hwp_instance()

# 새 한글 인스턴스 생성 (보안 모듈 자동 등록)
hwp = create_hwp_instance()
```
