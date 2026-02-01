# 테이블 병합 규칙

## 개요

Base 파일(원본)의 테이블에 Add 파일(추가 데이터)의 내용을 병합합니다.
병합은 `input_` 필드가 있는 영역에만 데이터를 추가하며, 기존 데이터(`data_`)는 유지됩니다.

## 파일 구분

| 구분 | 설명 |
|------|------|
| Base 파일 | 원본 HWPX 파일. 테이블 구조와 필드명(nc_name)이 설정된 템플릿 |
| Add 파일 | 추가할 데이터가 있는 파일. Base와 동일한 필드명 구조 |

## 접두사별 병합 동작

| 접두사 | 병합 시 동작 | 설명 |
|--------|-------------|------|
| `header_` | 유지 | 테이블 헤더. 변경 없음 |
| `data_` | 유지 | 기존 데이터. 변경 없음 |
| `add_` | 내용 추가 | 기존 셀 텍스트 뒤에 새 내용 추가 (행 추가 없음). 포맷터 적용 가능 |
| `stub_` | 새 행 생성 | 행 헤더. 데이터 추가 시 새 행 생성 |
| `gstub_` | rowspan 확장 | 그룹 헤더. 같은 값이면 rowspan 확장, 다른 값이면 새 셀 생성 |
| `input_` | 데이터 입력 | 빈 셀에 데이터 입력. 빈 셀 없으면 새 행 추가 |

## 병합 로직 상세

### 1. 기본 흐름

```
1. Base 파일에서 테이블 구조 파싱
2. Add 파일에서 데이터 추출 (빈 input 행 제외)
3. 필드명 매칭으로 데이터 병합
4. 필요시 행 추가 (gstub_ rowspan 확장 포함)
```

### 2. 데이터 추출 필터링

Add 파일에서 데이터 추출 시 다음 행은 제외됩니다:
- `data_` 필드만 있는 행 (기존 데이터 행)
- `input_` 값이 모두 비어있는 행 (빈 행)
- 헤더 행 (row 0)

### 3. input_ 필드 처리

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │  ← 유지
├──────────┼──────────┼──────────┤
│ gstub_X  │ input_1  │ input_2  │  ← 빈 셀이면 데이터 입력
│ (병합)   │          │          │
├──────────┼──────────┼──────────┤
│          │ input_1  │ input_2  │  ← gstub 영역 내 추가 행
├──────────┼──────────┼──────────┤
│ stub_Y   │ input_1  │ input_2  │  ← stub는 새 행 생성
└──────────┴──────────┴──────────┘
```

### 4. gstub_ 처리 규칙

**같은 gstub 값인 경우:**
- 기존 gstub 셀의 rowspan 확장
- 새 행에는 gstub 셀 없음 (rowspan으로 커버)

**다른 gstub 값인 경우:**
- 새 gstub 셀 생성
- rowspan=1로 시작, 이후 같은 값이면 확장

```
Before:                          After (같은 값 추가):
┌────────┬────────┐              ┌────────┬────────┐
│ gstub  │ input  │              │ gstub  │ data1  │
│ "A"    │        │              │ "A"    ├────────┤
│        │        │              │(rs=2)  │ data2  │  ← 새 행, gstub rowspan 확장
└────────┴────────┘              └────────┴────────┘

After (다른 값 추가):
┌────────┬────────┐
│ gstub  │ data1  │
│ "A"    │        │
├────────┼────────┤
│ gstub  │ data2  │  ← 새 gstub 셀 생성
│ "B"    │        │
└────────┴────────┘
```

**추가 파일에 gstub가 없는 경우:**
- 템플릿에 `gstub_`가 있지만 추가 파일에 없으면 → `stub_`로 대체
- `stub_`도 없으면 → `input_`처럼 개별 행으로 처리

```
템플릿:                           추가 데이터 (gstub 없음):
┌────────┬────────┐              data = [
│ gstub  │ input  │                {"input_1": "A"},
│ "X"    │        │                {"input_1": "B"},
└────────┴────────┘              ]

결과 (개별 행으로 처리):
┌────────┬────────┐
│ gstub  │ A      │  ← 빈 gstub 셀
│        │        │
├────────┼────────┤
│        │ B      │  ← 빈 gstub 셀 (개별 행)
└────────┴────────┘
```

### 5. stub_ 처리 규칙

stub_는 항상 새 행을 생성합니다.

```
Before:                          After:
┌────────┬────────┐              ┌────────┬────────┐
│ stub_X │ input  │              │ stub_X │ data1  │
└────────┴────────┘              ├────────┼────────┤
                                 │ stub_X │ data2  │  ← 새 행 + stub 복사
                                 └────────┴────────┘
```

### 6. 중첩 stub/gstub 처리

여러 stub/gstub가 연속으로 있는 경우:

```
┌──────────┬──────────┬──────────┬──────────┐
│ gstub_A  │ stub_B   │ input_1  │ input_2  │
│ (2행병합)│          │          │          │
├──────────┼──────────┼──────────┼──────────┤
│          │ stub_C   │ input_1  │ input_2  │
└──────────┴──────────┴──────────┴──────────┘
```

데이터 추가 시:
- `gstub_A`가 같으면 rowspan 확장
- `stub_B`, `stub_C`는 각각 새 행 생성

### 7. input_ rowspan 병합 처리

템플릿의 `input_` 셀이 rowspan으로 병합되어 있어도, 새 행 추가 시 **개별 셀(rowspan=1)**로 생성됩니다.

```
템플릿:                           추가 데이터:
┌──────────┬──────────┐          data = [
│ stub_A   │ input_1  │            {"stub_X": "X", "input_1": "A"},
├──────────┤ (2행병합)│            {"stub_X": "Y", "input_1": "B"},
│ stub_B   │          │          ]
└──────────┴──────────┘

결과 (각각 개별 셀로 복사):
┌──────────┬──────────┐
│ stub_A   │ input_1  │  ← 기존 (rowspan=2)
├──────────┤          │
│ stub_B   │          │
├──────────┼──────────┤
│ stub_X   │ A        │  ← 새 행 (rowspan=1)
├──────────┼──────────┤
│ stub_Y   │ B        │  ← 새 행 (rowspan=1)
└──────────┴──────────┘
```

## 데이터 매칭 규칙

### 필드명 기준 매칭

Add 파일의 input_ 데이터는 Base 파일의 동일한 필드명 위치에 매칭됩니다.

```python
# Add 파일 데이터 예시
add_data = {
    "input_abc123": "새 데이터1",
    "input_def456": "새 데이터2",
    "gstub_xyz789": "그룹A",  # gstub 값도 포함
}
```

### 그룹화된 input_ 처리

동일한 필드명을 가진 input_ 셀들은 같은 열에 데이터가 들어갑니다.

```
Base 테이블:
┌──────────┬──────────┐
│ stub_X   │ input_1  │  ← input_1 (row 1)
├──────────┼──────────┤
│ stub_Y   │ input_1  │  ← input_1 (row 2), 같은 필드명
└──────────┴──────────┘

Add 데이터: [{"input_1": "A"}, {"input_1": "B"}]

결과:
┌──────────┬──────────┐
│ stub_X   │ A        │
├──────────┼──────────┤
│ stub_Y   │ B        │
└──────────┴──────────┘
```

## 사용법

### CLI 사용법

```bash
# 기본 병합 (형식 검토 포함)
python -m merge.run_merge -o output.hwpx template.hwpx addition.hwpx

# 단순 병합 (형식 검토 없음)
python -m merge.run_merge -o output.hwpx --simple template.hwpx addition.hwpx

# 개요 구조만 출력
python -m merge.run_merge --list-outlines template.hwpx addition.hwpx

# SDK 비활성화 (정규식만 사용)
python -m merge.run_merge -o output.hwpx --no-sdk template.hwpx addition.hwpx
```

### Python API

```python
from merge.table import TableMerger

# 1. Base 파일 로드
merger = TableMerger()
merger.load_base_table("base.hwpx", table_index=0)

# 2. Add 데이터 준비
add_data = [
    {"gstub_abc": "그룹A", "input_123": "값1", "input_456": "값2"},
    {"gstub_abc": "그룹A", "input_123": "값3", "input_456": "값4"},  # gstub 확장
    {"gstub_abc": "그룹B", "input_123": "값5", "input_456": "값6"},  # 새 gstub
]

# 3. 병합 실행
merger.merge_with_stub(add_data)

# 4. 저장
merger.save("output.hwpx")
```

### 포맷터 사용

```python
from merge.table import TableMerger

# 포맷터 활성화 (기본값)
merger = TableMerger(use_formatter=True)

# 포맷터 비활성화
merger = TableMerger(use_formatter=False)

# 커스텀 설정 파일
merger = TableMerger(formatter_config_path="my_config.yaml")
```

## 병합 모드

| 모드 | 설명 |
|------|------|
| `fill_empty` | 빈 input_ 셀만 채움, 행 추가 안 함 |
| `append_row` | 항상 새 행 추가 |
| `smart` | 빈 셀 먼저 채우고, 부족하면 행 추가 (기본값) |

## add_ 필드 포맷터

`add_` 필드에 글머리 기호 등 포맷을 자동 적용할 수 있습니다.

### 설정 파일

`merge/formatters/table_formatter_config.yaml`:

```yaml
default:
  formatter: none
  separator: " "

fields:
  - pattern: "^add_.*"
    formatter: bullet
    options:
      style: default
      auto_detect: true

bullet:
  style: default
  styles:
    default:
      0: { symbol: "□ ", indent: " " }
      1: { symbol: "○", indent: "   " }
      2: { symbol: "- ", indent: "    " }
```

### 텍스트 구분자

```
기존 텍스트: "첫 번째 내용"
추가 텍스트: "두 번째 내용"

결과 (같은 문단): "첫 번째 내용 두 번째 내용"  ← 빈칸 1개
```

## 제약 사항

1. **data_ 셀은 변경 불가**: 기존 데이터 보존
2. **header_ 셀은 변경 불가**: 테이블 구조 유지
3. **colspan은 유지**: 병합된 열 구조 유지
4. **필드명 일치 필요**: Base와 Add의 필드명이 일치해야 매칭
5. **빈 input 행 무시**: Add 파일에서 input_ 값이 모두 비어있는 행은 병합 대상 제외
