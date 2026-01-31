# Workflow 4, 5, 6 종합 분석

## 멀티에이전트 관점 개요

```
┌─────────────────────────────────────────────────────────────────┐
│                    HWP 문서 처리 파이프라인                       │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  [HWP 원본] ─→ [HWPX 변환] ─→ [메타데이터 추출] ─→ [Excel 생성]  │
│      │              │               │                  │        │
│      │              │               ▼                  ▼        │
│      │              │         _meta.yaml          .xlsx         │
│      │              │         _para.yaml      _by_bookmark.xlsx │
│      │              │         _field.yaml                       │
│      │              │                                           │
│      └──────────────┴──────────────────────────────────────────┘
│                                                                 │
│  Workflow 4: 전체 테이블 → 단일 Excel                            │
│  Workflow 5: 북마크별 테이블 → 시트 분리 Excel                    │
│  Workflow 6: 색상 기반 필드 자동 설정                             │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 파일 생애주기 (Lifecycle)

```
Phase 1: 입력                Phase 2: 변환              Phase 3: 추출              Phase 4: 출력

┌──────────┐     COM API     ┌──────────┐    HWPXML     ┌──────────┐   openpyxl   ┌──────────┐
│   HWP    │ ─────────────→  │   HWPX   │ ───────────→  │ Metadata │ ──────────→  │  Excel   │
│ (원본)   │   FileSaveAs    │  (임시)   │    Parser     │  (YAML)  │              │ (최종)   │
└──────────┘                 └──────────┘               └──────────┘              └──────────┘
     │                            │                          │                         │
     ▼                            ▼                          ▼                         ▼
  사용자가               section*.xml에서            _meta.yaml              테이블 데이터 +
  색상 작업              tc, tbl 태그 파싱           _para.yaml              meta/para 시트
  (Workflow 6)           필드명/셀정보 추출          _field.yaml             북마크별 시트
```

---

## 메타데이터 생성 흐름

### 1. _meta.yaml 생성 경로

```
COM API (한글)                    HWPX XML                      출력

InsertTableField                 section*.xml
    │                                │
    ├─ collect_table_list_ids()      │
    │  (테이블별 list_id 수집)        │
    │                                │
    ├─ insert_field_to_xml() ───────→ tc.name에 JSON 삽입
    │  (테이블 idx, row, col 정보)    │  {"tblIdx": 0, "row": 0, "col": 0}
    │                                │
    └─ insert_caption_text()          │
       (캡션 텍스트 삽입)              │
                                      │
ExtractCellMeta                       │
    │                                │
    ├─ _extract_cell_positions()     │
    │  (COM API로 셀별 list_id)       │
    │                                │
    └─ _extract_field_names_from_hwpx() ─→ tc.name에서 JSON 파싱
                                      │
                                      ▼
                              _merge_to_yaml() → _meta.yaml
```

### 2. _para.yaml 생성 경로

```
COM API (한글)                               출력

GetParaStyle
    │
    ├─ get_all_para_styles()
    │     ├─ 문단별 순회 (MoveNextParaBegin)
    │     ├─ ParaShape 속성 추출 (align, line_spacing, margin)
    │     ├─ CharShape 속성 추출 (font, size, bold, color)
    │     └─ 위치 정보 (list_id, para_id, line_count)
    │
    └─ to_yaml() ─────────────────────→ _para.yaml
```

### 3. _field.yaml 생성 경로

```
HWPX XML                                          출력

GetCellDetail.from_hwpx_by_table()
    │
    ├─ 테이블별 셀 정보 파싱
    │   - row, col, row_span, col_span
    │   - border.bg_color (배경색)
    │   - paragraphs[].text (텍스트)
    │
    │        색상 판별
    │            │
    │     ┌──────┴──────┐
    │     ▼             ▼
    │  노란색         빨간색 (빈 셀)
    │   │               │
    │   │               ├─ 왼쪽 3개 텍스트 탐색
    │   │               └─ 위쪽 3개 텍스트 탐색
    │   ▼                       ▼
    │  셀 텍스트            [L:왼쪽][T:위쪽] 형식
    │   (20자)                     │
    │                              │
    └───────────┬──────────────────┘
                ▼
        tc.name에 설정 → _field.yaml
```

---

## 모듈 의존 관계

```
                           ┌───────────────────────┐
                           │   workflow/           │
                           │  workflow4_integrated │
                           │  workflow5_integrated │
                           └───────────┬───────────┘
                                       │
                    ┌──────────────────┼──────────────────┐
                    │                  │                  │
                    ▼                  ▼                  ▼
           ┌────────────────┐ ┌───────────────┐ ┌────────────────┐
           │  win32/        │ │  excel/       │ │  hwpxml/       │
           │ hwp_file_manager│ │ hwpx_to_excel │ │ get_cell_detail│
           │ insert_table_field│              │ │ get_table_property│
           │ extract_cell_meta│               │ │                │
           │ get_para_style │ │ cell_info_sheet│ │                │
           │ insert_field   │ │               │ │                │
           │ extract_field  │ │               │ │                │
           └────────┬───────┘ └───────┬───────┘ └────────┬───────┘
                    │                 │                   │
                    ▼                 ▼                   ▼
              ┌──────────┐      ┌──────────┐       ┌──────────┐
              │ pywin32  │      │ openpyxl │       │ xml.etree│
              │ (COM API)│      │          │       │ zipfile  │
              └──────────┘      └──────────┘       └──────────┘
```

---

## 참조 관계 (Key 연결)

```
테이블 식별자 체계

┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│    table_id     │     │    list_id      │     │    tbl_idx      │
│   (HWPX 고유)    │     │  (COM API 순서)  │     │   (순차 인덱스)  │
└────────┬────────┘     └────────┬────────┘     └────────┬────────┘
         │                       │                       │
         └───────────────────────┼───────────────────────┘
                                 ▼
                    ┌────────────────────────┐
                    │    _meta.yaml          │
                    │ - tbl_idx: 0           │
                    │   table_id: "20476..." │
                    │   cells:               │
                    │     - [row, col, ...,  │
                    │        list_id]        │
                    └────────────────────────┘

셀 위치 식별자: (table_id, row, col) ─────→ 고유 셀 식별
_meta.yaml + _field.yaml 연결: table_id + row + col 로 매칭
```

---

# Workflow 4: 통합 Excel 생성

workflow4_integrated.py - workflow1 + workflow2 + workflow3 조합

## 실행 방법
```bash
python workflow4_integrated.py [파일경로]
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow4_integrated.py" 2>&1
```

## 프로세스
1. HWP 열기 - 기존 인스턴스 연결 또는 새로 생성
2. HWPX 임시 변환
3. workflow1 실행 → {파일}_meta.yaml
4. workflow2 실행 → {파일}_para.yaml
5. workflow3 실행 → {파일}.xlsx
6. 임시 파일 삭제

## 출력 파일 (data/파일명/)
- {파일명}_meta.yaml : 테이블/셀 메타데이터
- {파일명}_para.yaml : 문단 스타일 정보
- {파일명}.xlsx : Excel 변환 결과

---

# Workflow 5: 북마크별 Excel 생성

workflow5_integrated.py - workflow1 + 북마크별 시트 분리

## 실행 방법
```bash
python workflow5_integrated.py [파일경로]
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow5_integrated.py" 2>&1
```

## 프로세스
1. HWP 열기
2. 북마크 목록 확인 (HeadCtrl 순회)
3. HWPX 임시 변환
4. 기존 필드 추출/백업
5. workflow1 실행 → {파일}_meta.yaml
6. 북마크별 시트 분리 → {파일}_by_bookmark.xlsx
7. meta 시트에 field 매칭

## 출력 파일 (data/파일명/)
- {파일명}_meta.yaml : 메타데이터
- {파일명}_field.yaml : 필드 정보
- {파일명}_by_bookmark.xlsx : 북마크별 시트

---

# Workflow 6: 색상 기반 셀 필드 자동 설정

insert_field.py - 색상 기반 필드명 자동 설정

## 실행 방법
```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_field.py" 2>&1
python insert_field.py <입력.hwp>
python insert_field.py <입력_insert_field.hwpx>
```

## HWP 파일 선택 시
1. HWP → _insert_field.hwpx 변환
2. 한글에서 파일 열림
3. 사용자 색상 블럭 작업 후 저장
4. 문서 닫으면 자동 감지 (XHwpDocuments.Count 폴링)
5. tc.name에 필드명 설정
6. YAML 출력

## 색상별 처리 규칙
| 색상 | 조건 | 필드명 형식 |
|------|------|-------------|
| 노란색 | 텍스트 있는 셀 | 셀 텍스트 (20자 제한) |
| 빨간색 | 빈 셀만 | [L:왼쪽][T:위쪽] |

## 빨간색 셀 탐색 규칙
- 왼쪽/위쪽으로 최대 3개 텍스트 수집
- 빨간색 범위 내에서만 탐색
- 병합 셀은 시작 위치로 점프

---

# 공통: 출력 폴더 구조

```
입력파일.hwp
입력파일_insert_field.hwpx    ← Workflow 6
data/
  └── 입력파일/
      ├── 입력파일_meta.yaml         ← Workflow 4, 5
      ├── 입력파일_para.yaml         ← Workflow 4
      ├── 입력파일_field.yaml        ← Workflow 5, 6
      ├── 입력파일.xlsx              ← Workflow 4
      └── 입력파일_by_bookmark.xlsx  ← Workflow 5
```

## Workflow별 산출물 요약
| Workflow | 산출 파일 |
|----------|-----------|
| **4** | _meta.yaml, _para.yaml, .xlsx |
| **5** | _meta.yaml, _field.yaml, _by_bookmark.xlsx |
| **6** | _insert_field.hwpx, _field.yaml |

---

# 멀티에이전트 분석

## 에이전트 역할 분담

```
┌─────────────────────────────────────────────────────────────────────┐
│                     Multi-Agent Architecture                        │
├─────────────────────────────────────────────────────────────────────┤
│  ┌─────────────┐   ┌─────────────┐   ┌─────────────┐               │
│  │   Agent 1   │   │   Agent 2   │   │   Agent 3   │               │
│  │  COM API    │   │ XML Parser  │   │   Excel     │               │
│  │  Handler    │   │  Handler    │   │  Generator  │               │
│  ├─────────────┤   ├─────────────┤   ├─────────────┤               │
│  │ • HWP 열기   │   │ • HWPX 압축 │   │ • 시트 생성  │               │
│  │ • 문서 탐색  │   │   해제      │   │ • 셀 스타일 │               │
│  │ • 속성 추출  │   │ • XML 파싱  │   │ • 데이터 병합│               │
│  │ • 필드 삽입  │   │ • 태그 수정  │   │ • 파일 저장  │               │
│  │ • HWPX 변환  │   │ • 압축 복원 │   │              │               │
│  └──────┬──────┘   └──────┬──────┘   └──────┬──────┘               │
│         └────────────────┬┴─────────────────┘                       │
│                          ▼                                          │
│              ┌───────────────────────┐                              │
│              │    Orchestrator       │                              │
│              │    (Workflow 4/5/6)   │                              │
│              │  • 단계별 실행 조율    │                              │
│              │  • 오류 처리          │                              │
│              │  • 리소스 정리        │                              │
│              └───────────────────────┘                              │
└─────────────────────────────────────────────────────────────────────┘
```

## 데이터 흐름 상세

### Workflow 4 실행 순서
```
Step 1 → Step 2 → Step 3 → Step 4
get_hwp_instance() → save_as_hwpx() → run_workflow1/2() → run_workflow3()
[HWP 연결/생성]    [임시 HWPX 생성]   [메타/문단 추출]    [Excel + meta시트]
```

### Workflow 5 실행 순서
```
Step 1-3 → Step 4 → Step 5 → Step 6
(WF4 동일) → extract_existing_fields() → run_workflow1() → run_bookmark_excel()
           [기존 필드 백업]             [메타+필드복원]    [북마크별 시트분리]
                                                              ↓
                                                    add_meta_to_sheets()
                                                    [meta시트에 field 매칭]
```

### Workflow 6 실행 순서
```
HWP 선택: convert_hwp_to_hwpx() → 사용자 색상작업 → 문서닫힘 감지
                      ↓                                   ↓
HWPX 선택: ─────────────────────────────────────→ process_hwpx_field()
                                                  [색상 기반 필드 설정]
                                                          ↓
                                                  _field.yaml 저장
```

## 오류 처리 전략

| 단계 | 가능한 오류 | 처리 방법 |
|------|------------|----------|
| HWP 연결 | 한글 미실행 | create_hwp_instance() 새 인스턴스 |
| HWPX 변환 | 파일 잠금 | hwp.Clear(1) 문서 닫기 |
| XML 파싱 | 인코딩 오류 | utf-8 명시적 지정 |
| 필드 삽입 | 테이블 없음 | count == 0 체크 후 건너뜀 |
| 정리 | 임시파일 삭제 실패 | try-except 무시 |

---

# API 참조

## COM API 주요 메서드
| 메서드 | 설명 | 사용처 |
|--------|------|--------|
| hwp.RegisterModule() | 보안 모듈 등록 | 인스턴스 생성 시 |
| hwp.HAction.Execute() | 액션 실행 | 파일 열기/저장 |
| hwp.HeadCtrl | 컨트롤 순회 시작 | 북마크 탐색 |
| hwp.MoveToBookmark() | 북마크로 이동 | 마커 삽입 |
| hwp.XHwpDocuments.Count | 열린 문서 수 | 문서 닫힘 감지 |

## HWPX XML 주요 태그
| 태그 | 속성 | 설명 |
|------|------|------|
| hp:tbl | id | 테이블 고유 ID |
| hp:tr | - | 테이블 행 |
| hp:tc | name | 테이블 셀 (필드명 저장) |
| hp:cellAddr | rowAddr, colAddr | 셀 주소 |
| hp:bookmark | name | 북마크 이름 |

## YAML 스키마

### _meta.yaml
```yaml
tables:
  - tbl_idx: 0
    table_id: "2047609131"
    type: parent | nested
    size: "5x3"
    list_range: "1234-1256"
    caption_list_id: 1233
    caption: "표 1. 연구 현황"
    cells:
      - [row, col, row_span, col_span, list_id]
```

### _field.yaml
```yaml
- table_idx: 1
  list_id: 1234
  table_id: "2047609131"
  row: 26
  col: 3
  field_name: "[L:연구시설][T:구입기관]"
  type: red | yellow
```

### _para.yaml
```yaml
- list_id: 1234
  para_id: 0
  text: "문단 내용..."
  style_name: "본문"
  align: "center"
  line_spacing: 160
  font_name: "맑은 고딕"
  font_size: 10
  bold: false
```

---

# 워크플로우 조합 시나리오

## 시나리오 1: 전체 문서 분석
```bash
# Workflow 4 실행 → 전체 테이블 Excel + 메타데이터
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow4_integrated.py" 2>&1
# 결과: _meta.yaml, _para.yaml, .xlsx
```

## 시나리오 2: 북마크 기반 분석 + 필드 설정
```bash
# 1. Workflow 6 먼저 실행 → 필드 설정
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_field.py" 2>&1

# 2. Workflow 5 실행 → 북마크별 시트 + 필드 포함
cmd.exe /c "cd /d C:\hwp_xml\workflow && python workflow5_integrated.py" 2>&1
# 결과: _field.yaml, _meta.yaml, _by_bookmark.xlsx (field_name 포함)
```

## 시나리오 3: 필드만 추출
```bash
# Workflow 6만 실행 (이미 색상 작업된 HWPX)
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_field.py 파일_insert_field.hwpx" 2>&1
# 결과: _field.yaml만 생성
```

---

# 상세 클래스 구조

## Workflow4 클래스 상세

```python
class Workflow4:
    """통합 워크플로우 실행기"""

    # 인스턴스 변수
    hwp: object          # 한글 COM 객체
    hwp_created: bool    # 새로 생성 여부 (True면 종료 시 Quit)
    filepath: str        # 원본 HWP 파일 경로
    temp_hwpx: str       # 임시 HWPX 파일 경로
    cell_positions: dict # COM API에서 추출한 셀 위치 정보
    field_names: list    # HWPX에서 추출한 필드명 정보
    para_styles: list    # 문단 스타일 정보

    # 주요 메서드
    def _get_hwp() -> object:
        """기존 한글 연결 또는 새 인스턴스 생성"""

    def _open_file(filepath: str) -> str:
        """HWP 파일 열기 (인자/열린문서/대화상자)"""

    def _save_as_hwpx() -> str:
        """HWP → 임시 HWPX 변환 (temp폴더/workflow4_temp.hwpx)"""

    def _run_workflow1(base_path: str) -> str:
        """메타데이터 추출 → _meta.yaml"""

    def _run_workflow2(base_path: str) -> str:
        """문단 스타일 추출 → _para.yaml"""

    def _run_workflow3(base_path: str, split_by_para: bool) -> str:
        """Excel 변환 + 메타 시트 추가"""

    def _add_meta_sheet(wb: Workbook):
        """meta 시트: 테이블/셀 정보"""

    def _add_para_sheet(wb: Workbook):
        """para 시트: 문단 스타일 정보"""

    def _cleanup():
        """임시 파일 삭제"""

    def run(filepath: str, split_by_para: bool) -> dict:
        """전체 워크플로우 실행"""
```

## Workflow5 클래스 상세

```python
class Workflow5:
    """북마크별 Excel 생성 워크플로우"""

    # 인스턴스 변수
    hwp: object
    hwp_created: bool
    filepath: str
    temp_hwpx: str
    output_dir: str         # 결과 저장 폴더 (data/파일명/)
    bookmarks: list         # 북마크 목록
    markers_inserted: bool  # 마커 삽입 여부
    cell_positions: dict
    field_names: list
    existing_fields: list   # 기존 필드 목록 (복원용)
    field_extractor: object # ExtractField 인스턴스

    # 주요 메서드
    def _create_output_dir() -> str:
        """data/파일명/ 폴더 생성"""

    def _get_bookmarks() -> int:
        """HeadCtrl 순회로 북마크 개수 확인"""

    def _insert_bookmark_markers():
        """북마크 위치에 {{BOOKMARK:이름}} 마커 삽입"""

    def _remove_bookmark_markers():
        """원본에서 마커 삭제"""

    def _extract_existing_fields(base_path: str) -> str:
        """기존 tc.name 추출 → _field.yaml (삭제 없이 백업)"""

    def _run_workflow1(base_path: str) -> str:
        """메타데이터 추출 + 기존 필드 복원"""

    def _restore_original_field_names(hwpx_path: str):
        """JSON 필드 삭제 후 기존 필드명 복원"""

    def _run_bookmark_excel(base_path: str, split_by_para: bool) -> str:
        """북마크별 시트 분리 Excel 생성"""

    def _add_meta_to_sheets(excel_path: str, meta_yaml: str, field_yaml: str):
        """meta 시트에 _meta + _field 데이터 병합"""

    def run(filepath: str, split_by_para: bool) -> dict:
        """전체 워크플로우 실행"""
```

---

# 데이터 구조 상세

## COM API 셀 위치 정보 (cell_positions)

```python
# ExtractCellMeta._extract_cell_positions() 반환값
cell_positions = [
    {   # 테이블 0
        'table_id': '2047609131',
        'row_count': 5,
        'col_count': 3,
        'cells': {
            (0, 0): [1234, 1, 1],  # (row, col): [list_id, row_span, col_span]
            (0, 1): [1235, 1, 1],
            (1, 0): [1236, 2, 1],  # row_span=2 (병합 셀)
            ...
        }
    },
    {   # 테이블 1
        ...
    }
]
```

## HWPX 필드명 정보 (field_names)

```python
# ExtractCellMeta._extract_field_names_from_hwpx() 반환값
field_names = [
    {   # 테이블 0
        'table_id': '2047609131',
        'row_count': 5,
        'col_count': 3,
        'caption': '표 1. 연구 현황',
        'cells': [
            {
                'row': 0,
                'col': 0,
                'row_span': 1,
                'col_span': 1,
                'field_name': '{"tblIdx":0,"row":0,"col":0,"type":"parent"}'
            },
            ...
        ]
    }
]
```

## 문단 스타일 정보 (para_styles)

```python
# GetParaStyle.get_all_para_styles() 반환값
@dataclass
class ParaStyle:
    list_id: int         # 문단 소속 컨트롤의 list_id
    para_id: int         # 문단 번호 (0부터 시작)
    text: str            # 문단 텍스트
    line_count: int      # 줄 수
    start_char_id: int   # 시작 글자 ID
    next_line_char_id: int  # 다음줄 시작 글자 ID
    end_char_id: int     # 끝 글자 ID
    style_name: str      # 스타일 이름 ("본문", "제목1" 등)
    align: str           # 정렬 ("left", "center", "right", "justify")
    indent: float        # 들여쓰기 (HWPUNIT → pt 변환)
    margin_left: float   # 왼쪽 여백
    margin_right: float  # 오른쪽 여백
    line_spacing: int    # 줄간격 (%)
    line_spacing_type: str  # 줄간격 타입 ("percent", "fixed")
    space_before: float  # 문단 앞 간격
    space_after: float   # 문단 뒤 간격
    char_style: CharStyle  # 글자 스타일 정보

@dataclass
class CharStyle:
    font_name: str       # 폰트 이름
    font_size: float     # 폰트 크기 (pt)
    bold: bool           # 굵게
    italic: bool         # 기울임
    underline: bool      # 밑줄
    strikeout: bool      # 취소선
    text_color: str      # 글자색 (HEX)
    highlight_color: str # 형광펜 색상 (HEX)
```

---

# HWPX XML 구조 상세

## section*.xml 테이블 구조

```xml
<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">

  <!-- 테이블 시작 -->
  <hp:tbl id="2047609131" rowCnt="5" colCnt="3">

    <!-- 테이블 캡션 (있을 경우) -->
    <hp:caption>
      <hp:subList>
        <hp:p>표 1. 연구 현황</hp:p>
      </hp:subList>
    </hp:caption>

    <!-- 테이블 행 -->
    <hp:tr>
      <!-- 테이블 셀 -->
      <hp:tc name="필드명저장위치" rowSpan="1" colSpan="1">
        <hp:cellAddr rowAddr="0" colAddr="0"/>
        <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
        <hp:cellBorder borderFill="1"/>
        <hp:subList>
          <hp:p>
            <hp:run>
              <hp:t>셀 텍스트</hp:t>
            </hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>

      <!-- 병합 셀 예시 -->
      <hp:tc rowSpan="2" colSpan="1">
        <hp:cellAddr rowAddr="1" colAddr="0"/>
        ...
      </hp:tc>

    </hp:tr>

    <!-- 중첩 테이블 (셀 내부) -->
    <hp:tr>
      <hp:tc>
        <hp:subList>
          <hp:tbl id="다른ID">  <!-- 중첩 테이블 -->
            ...
          </hp:tbl>
        </hp:subList>
      </hp:tc>
    </hp:tr>

  </hp:tbl>

</hp:sec>
```

## 북마크 XML 구조

```xml
<!-- 단일 위치 북마크 -->
<hp:bookmark name="연구개요"/>

<!-- 범위 북마크 -->
<hp:bookmarkStart name="연구목표"/>
  ... 북마크 범위 내 콘텐츠 ...
<hp:bookmarkEnd name="연구목표"/>
```

---

# 색상 판별 로직 상세

## insert_field.py 색상 함수

```python
def is_red_color(color: str) -> bool:
    """빨간색 계열 판별

    판별 조건:
    1. 정확한 색상값: #ff0000, #cf2741 등
    2. RGB 근사값: R > 180, G < 80, B < 80
    """
    red_colors = ['ff0000', 'cf2741', 'ff0000ff', 'cf2741ff']

    # 정확한 매칭
    if color_lower in red_colors:
        return True

    # RGB 근사 판별
    r = int(color[0:2], 16)
    g = int(color[2:4], 16)
    b = int(color[4:6], 16)
    return r > 180 and g < 80 and b < 80

def is_yellow_color(color: str) -> bool:
    """노란색 계열 판별

    판별 조건:
    1. 정확한 색상값: #ffff00 등
    2. RGB 근사값: R > 200, G > 200, B < 100
    """
    yellow_colors = ['ffff00', 'ffff00ff', 'fff000', 'fff000ff']

    # RGB 근사 판별
    return r > 200 and g > 200 and b < 100
```

## 빨간색 셀 필드명 생성 알고리즘

```
입력: 빨간색 배경의 빈 셀 (row, col)

1. 왼쪽 탐색 (최대 3개)
   c = col - 1
   while c >= 0 and len(left_texts) < 3:
       cell = find_cell_at(row, c)
       if not is_red_color(cell.bg_color):
           break  # 빨간색 범위 벗어남
       if cell.text:
           left_texts.append(cell.text)
       c = cell.start_col - 1  # 병합 셀 점프

2. 위쪽 탐색 (최대 3개)
   r = row - 1
   while r >= 0 and len(top_texts) < 3:
       cell = find_cell_at(r, col)
       if not is_red_color(cell.bg_color):
           break  # 빨간색 범위 벗어남
       if cell.text:
           top_texts.append(cell.text)
       r = cell.start_row - 1  # 병합 셀 점프

3. 필드명 조합
   parts = []
   for t in left_texts:
       parts.append('[L:' + t + ']')
   for t in top_texts:
       parts.append('[T:' + t + ']')

   field_name = ''.join(parts)
   # 예: "[L:연구시설][L:장비][T:구입기관][T:수량]"
```

---

# Excel 출력 시트 상세

## meta 시트 컬럼

| 컬럼 | 타입 | 설명 | 예시 |
|------|------|------|------|
| tbl_idx | int | 테이블 순번 | 0, 1, 2 |
| table_id | str | HWPX 테이블 ID | "2047609131" |
| type | str | 테이블 타입 | "parent", "nested" |
| size | str | 테이블 크기 | "5x3" |
| list_range | str | list_id 범위 | "1234-1256" |
| caption_list_id | int | 캡션 list_id | 1233 |
| caption | str | 캡션 텍스트 | "표 1. 연구 현황" |
| row | int | 셀 행 위치 | 0 |
| col | int | 셀 열 위치 | 2 |
| row_span | int | 행 병합 | 1 |
| col_span | int | 열 병합 | 2 |
| list_id | int | 셀 list_id | 1245 |

## para 시트 컬럼

| 컬럼 | 타입 | 설명 |
|------|------|------|
| list_id | int | 소속 컨트롤 list_id |
| para_id | int | 문단 번호 |
| text | str | 문단 텍스트 (1000자 제한) |
| line_count | int | 줄 수 |
| start_char | int | 시작 글자 ID |
| next_line_char | int | 다음줄 시작 글자 ID |
| end_char | int | 끝 글자 ID |
| style_name | str | 스타일 이름 |
| align | str | 정렬 |
| indent | float | 들여쓰기 |
| margin_left | float | 왼쪽 여백 |
| margin_right | float | 오른쪽 여백 |
| line_spacing | int | 줄간격 (%) |
| line_spacing_type | str | 줄간격 타입 |
| space_before | float | 문단 앞 간격 |
| space_after | float | 문단 뒤 간격 |
| font_name | str | 폰트 이름 |
| font_size | float | 폰트 크기 |
| bold | bool | 굵게 |
| italic | bool | 기울임 |
| underline | bool | 밑줄 |
| strikeout | bool | 취소선 |
| text_color | str | 글자색 |
| highlight_color | str | 형광펜 색상 |

---

# 단위 변환 참조

## HWPUNIT 변환

```python
# core/unit.py

# HWPUNIT: 한글 내부 단위 (1/7200 인치)
HWPUNIT_PER_INCH = 7200
HWPUNIT_PER_PT = 100  # 1pt = 100 HWPUNIT
HWPUNIT_PER_MM = 283.465  # 1mm ≈ 283.465 HWPUNIT

def hwpunit_to_pt(hwpunit: int) -> float:
    """HWPUNIT → 포인트 변환"""
    return hwpunit / HWPUNIT_PER_PT

def hwpunit_to_mm(hwpunit: int) -> float:
    """HWPUNIT → 밀리미터 변환"""
    return hwpunit / HWPUNIT_PER_MM

def pt_to_hwpunit(pt: float) -> int:
    """포인트 → HWPUNIT 변환"""
    return int(pt * HWPUNIT_PER_PT)
```

## 색상 변환

```python
# HWPX XML 색상 형식: RRGGBB 또는 RRGGBBAA (알파 포함)
# Excel 색상 형식: AARRGGBB (openpyxl)

def hwpx_to_excel_color(hwpx_color: str) -> str:
    """HWPX 색상 → Excel 색상 변환"""
    color = hwpx_color.lstrip('#')
    if len(color) == 6:
        return 'FF' + color.upper()  # 알파 추가
    elif len(color) == 8:
        return color[6:8] + color[0:6]  # 알파 위치 변경
    return 'FF000000'  # 기본값 (검정)
```

---

# 보안 모듈 상세

## SecurityModule 동작

```python
# win32/hwp_file_manager.py

def create_hwp_instance(visible: bool = True) -> object:
    """보안 모듈이 적용된 한글 인스턴스 생성"""

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = visible

    # 보안 모듈 등록 - 모든 보안 경고 자동 허용
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

    return hwp

# SecurityModule이 허용하는 항목:
# - 스크립트 실행 경고
# - 매크로 실행 경고
# - 개인정보 포함 경고
# - 외부 파일 접근 경고
# - ActiveX 컨트롤 경고
```

## open_hwp 함수 (팝업 방지)

```python
def open_hwp(hwp, filepath: str) -> bool:
    """팝업 없이 HWP 파일 열기

    hwp.Open()은 팝업이 발생할 수 있으므로
    HAction.Execute("FileOpen")을 사용
    """
    hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = filepath
    hwp.HParameterSet.HFileOpenSave.Format = ""  # 자동 감지
    return hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
```
