# 주의: 사용자가 입력을 요청한 경우에만 수정하세요.
# 간략히 요점만 전체적인 맥락만 참고할 수 있도록 작성해 주세요.

## WSL에서 Windows Python 실행

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_ctrl_id.py" 2>&1
```

- `cmd.exe /c` : Windows cmd 실행
- `cd /d C:\경로` : Windows 경로로 이동
- `2>&1` : stderr도 출력


## check_, test_ 접미사는 삭제해도 되며, 가능한 테스트 파일이나 체크 파일에 메인 로직을 넣지 말고 기존 로직을 쓰고, 기존 로직을 수정하고 핵라

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

## 중첩 테이블 처리

HWPX에서 테이블 셀(`tc`) 안에 또 다른 테이블(`tbl`)이 포함될 수 있음.

### 문제점
`root.iter('tc')`로 모든 셀을 가져오면 부모/자식 테이블 셀이 문서 순서로 섞여서 테이블별 셀 개수로 나눌 때 데이터가 어긋남.

### 해결 방법 (get_cell_detail.py)
- `from_hwpx_by_table()`: 테이블별로 그룹화된 셀 정보 반환
- `_find_tables_recursive()`: 재귀적으로 테이블을 문서 순서대로 탐색
- `_parse_table_direct_cells()`: 각 테이블의 직접 셀만 파싱 (중첩 테이블 내부 셀 제외)

```python
# 사용 예시
parser = GetCellDetail()
table_cell_details = parser.from_hwpx_by_table(hwpx_path)
# table_cell_details[0]: 첫 번째 테이블의 셀들
# table_cell_details[1]: 두 번째 테이블의 셀들 (중첩 테이블 포함)
```

## 폴더 구조 (계속)

### core/
공통 유틸리티 모듈
- `unit.py`: HWPUNIT 단위 변환
- `file_dialog.py`: Windows 파일 탐색기 대화상자 (WSL에서 PowerShell 사용)
  - `open_hwp_dialog()`: HWP/HWPX 파일 선택
  - `windows_to_wsl_path()`: Windows↔WSL 경로 변환

### win32/
Windows 한글 COM API 연동 (WSL에서 cmd.exe로 실행)
- `hwp_file_manager.py`: 공통 유틸리티 (인스턴스 연결, 파일 대화상자, 저장)
- `convert_hwp.py`: HWP↔HWPX 변환
- `insert_ctrl_id.py`: 테이블에 ctrl_id 메타데이터 삽입
- `insert_table_field.py`: 테이블 셀에 필드 이름(메타데이터) 설정
- `get_para_style.py`: 문단 스타일 추출 → `파일_para.yaml`
- `get_table_property.py`: COM API로 테이블 속성 추출
- `get_table_info.py`: 테이블 위치/구조 조회
- `insert_field_within_redcells.py`: 빨간색 배경 빈 셀에 필드명 자동 설정 (`[왼쪽][위쪽]` 형식)

## 한글 COM API 공통 유틸리티 (hwp_file_manager.py)

```python
from hwp_file_manager import get_hwp_instance, create_hwp_instance, get_active_filepath, open_file_dialog, save_hwp

hwp = get_hwp_instance()  # 열린 한글 연결 (없으면 None)
hwp = create_hwp_instance()  # 새 한글 생성 (SecurityModule 포함)
filepath = get_active_filepath(hwp)  # 열린 문서 경로
filepath = open_file_dialog()  # 파일 선택 대화상자 (Windows 탐색기 열림)
save_hwp(hwp, filepath, "HWP")  # 편집 가능하게 저장
```

## 보안 모듈 (RegisterModule)

`create_hwp_instance()`에서 `SecurityModule` 등록으로 모든 보안 경고 자동 허용:
SecurityModule 만 사용할 것

```python
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
```

| 모듈명 | 설명 |
|--------|------|
| `FilePathCheckerModule` | 파일 경로 접근만 허용 |
| `SecurityModule` | 모든 보안 경고 자동 허용 (스크립트, 매크로, 개인정보 등) |

### SecurityModule 허용 항목
- 스크립트 실행 경고
- 매크로 실행 경고
- 개인정보 포함 경고
- 외부 파일 접근 경고

### 테스트 방법
```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python security_module_test.py" 2>&1
```

## caption_list_id 규칙

- **parent 테이블**: 항상 caption_list_id 존재 (첫 셀 list_id - 1)
- **nested 테이블**: caption이 있으면 caption_list_id, 없으면 null

## 북마크 조회

### HWPX XML 태그
HWPX는 ZIP 파일로 `Contents/section*.xml`에 북마크 저장

| 태그 | 설명 |
|------|------|
| `{namespace}bookmark` | 북마크 단일 태그 |
| `{namespace}bookmarkStart` | 북마크 시작 태그 |
| `{namespace}bookmarkEnd` | 북마크 끝 태그 |

### COM API vs HWPX
- COM API (HeadCtrl 순회): Name이 None으로 나옴 → 개수만 확인 가능
- HWPX XML 파싱: 북마크 이름 정상 조회 가능

**권장**: COM API로 개수 확인 → HWPX 변환 후 XML에서 이름 파싱