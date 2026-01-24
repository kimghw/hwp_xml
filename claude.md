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
- `get_cell_detail.py`: 셀 스타일 정보
- `extract_cell_index.py`: `[index:##숫자]` 패턴 매핑
- `get_page_property.py`: 페이지 속성, 단위 변환

### win32/
Windows 한글 COM API 연동 (WSL에서 cmd.exe로 실행)
- `insert_ctrl_id.py`: 테이블에 ctrl_id 메타데이터 삽입
- `insert_table_field.py`: 테이블 셀에 필드 이름(메타데이터) 설정
  - tc 태그 name 속성에 JSON 삽입: `{"tblIdx":N,"rowAddr":R,"colAddr":C,"rowSpan":RS,"colSpan":CS}`
  - 캡션에 `{caption:tbl_N|}` 삽입
  - 값 채우기용 참조 메타데이터 (실제 값 삽입 로직 없음)
- `get_table_property.py`: COM API로 테이블 속성 추출
- `get_table_info.py`: 테이블 위치/구조 조회

## 한글 COM API 주의사항

### 한글 인스턴스 연결
- `get_hwp_instance()`: ROT에서 실행 중인 한글에 연결 (없으면 None)
- 한글이 없으면 `win32.gencache.EnsureDispatch("hwpframe.hwpobject")`로 새로 실행
- 보안 모듈 등록 필수: `hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")`

### 파일 저장 (편집 가능하게)
`hwp.SaveAs()`는 읽기 전용으로 저장될 수 있음. 아래 방식 사용:
```python
hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = filepath
hwp.HParameterSet.HFileOpenSave.Format = "HWP"  # 또는 "HWPX"
hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
```

### 저장 후 문서 닫기
저장 후 파일 잠금 해제를 위해 문서 닫기 필요:
```python
hwp.Clear(1)  # 1: 저장 안 함 (이미 저장했으므로)
```