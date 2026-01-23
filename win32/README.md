# win32

Windows 아래아한글 COM API 연동 모듈

## 주요 파일

| 파일 | 설명 |
|------|------|
| `insert_ctrl_id.py` | 테이블에 ctrl_id(list_id, para_id 등) 메타데이터 삽입 |
| `get_table_property.py` | COM API로 테이블 속성 추출 → DataFrame 변환 |
| `get_table_info.py` | 모든 테이블 위치/구조 정보 조회 |

## 실행 방법

Windows 전용 모듈이므로 WSL에서는 cmd.exe로 실행:

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_ctrl_id.py" 2>&1
```

## 사용법

```python
from win32 import TableProperty, insert_ctrl_id

# 테이블 속성 추출
tp = TableProperty("document.hwp")
tables = tp.get_all_tables()

# ctrl_id 삽입
insert_ctrl_id("document.hwp")
```
