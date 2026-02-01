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
| `add_` | 내용 추가 | 기존 셀 텍스트 뒤에 새 내용 추가 (행 추가 없음) |
| `stub_` | 새 행 생성 | 행 헤더. 데이터 추가 시 새 행 생성 |
| `gstub_` | rowspan 확장 | 그룹 헤더. 같은 값이면 rowspan 확장, 다른 값이면 새 셀 생성 |
| `input_` | 데이터 입력 | 빈 셀에 데이터 입력. 빈 셀 없으면 새 행 추가 |

## 병합 로직 상세

### 1. 기본 흐름

```
1. Base 파일에서 테이블 구조 파싱
2. Add 파일에서 input_ 필드 데이터 추출
3. 필드명 매칭으로 데이터 병합
4. 필요시 행 추가 (gstub_ rowspan 확장 포함)
```

### 2. input_ 필드 처리

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

### 3. gstub_ 처리 규칙

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

### 4. stub_ 처리 규칙

stub_는 항상 새 행을 생성합니다.

```
Before:                          After:
┌────────┬────────┐              ┌────────┬────────┐
│ stub_X │ input  │              │ stub_X │ data1  │
└────────┴────────┘              ├────────┼────────┤
                                 │ stub_X │ data2  │  ← 새 행 + stub 복사
                                 └────────┴────────┘
```

### 5. 중첩 stub/gstub 처리

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

### 6. input_ rowspan 병합 처리

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

```python
from merge.table_merger import TableMerger

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

## 병합 모드

| 모드 | 설명 |
|------|------|
| `fill_empty` | 빈 input_ 셀만 채움, 행 추가 안 함 |
| `append_row` | 항상 새 행 추가 |
| `smart` | 빈 셀 먼저 채우고, 부족하면 행 추가 (기본값) |

## 스타일 규칙

### 테이블 셀 스타일

| 상황 | 스타일 처리 |
|------|------------|
| `input_` 데이터 입력 | Base 테이블 셀 스타일 유지 |
| 새 행 추가 | Base 테이블 마지막 행 스타일 복사 |
| `add_` 내용 추가 | 기존 셀 스타일 유지 |
| 같은 문단 내 추가 | 빈칸 1개로 구분 |

### 텍스트 구분자

```
기존 텍스트: "첫 번째 내용"
추가 텍스트: "두 번째 내용"

결과 (같은 문단): "첫 번째 내용 두 번째 내용"  ← 빈칸 1개
결과 (새 문단):   "첫 번째 내용\n두 번째 내용"  ← 줄바꿈
```

## Claude Code SDK 검증

병합 전 데이터 형식을 Claude Code SDK로 검증하여 양식에 맞게 조정합니다.

### 검증 대상

| 대상 | 검증 내용 |
|------|----------|
| 개요 본문 | 문단 형식, 헤딩 레벨, 목록 스타일 검증 후 병합 |
| `add_` 필드 | 기존 셀 형식에 맞게 검증 후 추가 |
| `input_` 필드 | 데이터 형식 검증 (선택적) |

### 검증 흐름

```
1. 개요 본문 병합
   ├─ Base 문서에서 개요 구조 파싱
   ├─ Add 데이터 SDK 검증
   ├─ 형식 맞춤 (헤딩, 목록 등)
   └─ 병합

2. add_ 필드 추가
   ├─ 기존 셀 스타일 분석
   ├─ Add 데이터 SDK 검증
   ├─ 양식에 맞게 조정
   └─ 기존 텍스트 뒤에 추가

3. input_ 필드 입력
   ├─ 데이터 형식 검증 (선택적)
   └─ Base 셀 스타일로 입력
```

### SDK 검증 예시

```python
from merge.format_validator import AddFieldValidator, create_sdk_validator

# 기본 규칙 기반 검증
validator = AddFieldValidator()

# add_ 필드 검증
add_text = "추가할 내용..."
result = validator.validate_add_content(
    add_text,
    base_cell_style="bullet_list"  # 기존 셀이 bullet list면 맞춤
)
print(result.validated_text)  # 검증/조정된 텍스트
print(result.changes_made)    # 변경 내역

# 개요 본문 검증
outline_text = "## 섹션 제목\n내용..."
result = validator.validate_outline(
    outline_text,
    target_level=2  # 목표 헤딩 레벨
)

# input_ 필드 검증 (선택적)
input_text = "2024-01-15"
result = validator.validate_input_content(
    input_text,
    expected_format="date"  # date, number, text
)

# 일괄 검증
data_list = [
    {"add_memo": "메모1", "input_value": "100"},
    {"add_memo": "메모2", "input_value": "200"},
]
results = validator.validate_batch(
    data_list,
    field_styles={"add_memo": "plain"}
)
```

### Claude Code SDK 연동

```python
from merge.format_validator import AddFieldValidator, create_sdk_validator

# SDK 클라이언트가 있는 경우
# sdk_client = ... (실제 Claude Code SDK 클라이언트)
# sdk_validator = create_sdk_validator(sdk_client)
# validator = AddFieldValidator(sdk_validator=sdk_validator)

# SDK 없이 기본 규칙만 사용
validator = AddFieldValidator()
```

### TableMerger에서 검증 사용

```python
from merge.table_merger import TableMerger

# 검증 활성화
merger = TableMerger(validate_format=True)
merger.load_base_table("base.hwpx", table_index=0)

# add_ 필드 스타일 지정
field_styles = {
    "add_memo": "bullet_list",  # 글머리 기호 목록
    "add_note": "plain",        # 일반 텍스트
}

add_data = [
    {"add_memo": "첫 번째 메모", "input_value": "100"},
    {"add_memo": "두 번째 메모", "input_value": "200"},
]

# 병합 실행 (검증 포함)
merger.merge_with_stub(
    add_data,
    field_styles=field_styles,
    add_separator=" "  # 같은 문단 구분자 (빈칸 1개)
)

merger.save("output.hwpx")
```

## 제약 사항

1. **data_ 셀은 변경 불가**: 기존 데이터 보존
2. **header_ 셀은 변경 불가**: 테이블 구조 유지
3. **colspan은 유지**: 병합된 열 구조 유지
4. **필드명 일치 필요**: Base와 Add의 필드명이 일치해야 매칭
5. **SDK 검증 실패 시**: 원본 데이터 그대로 사용 (경고 로그)
