# nc_name 자동 생성 규칙

## 접두사 분류

| 접두사 | 조건 | 설명 |
|--------|------|------|
| `header_` | 행 전체가 배경색 있음 + 연속된 행도 모두 배경색 | 테이블 헤더 셀 |
| `add_` | (최상단 데이터 행 + 텍스트 30자 이상 + 배경색 없음) 또는 (1x1 단일 셀 + 배경색 없음) | 병합 시 내용만 추가 |
| `stub_` | 텍스트 있음 + 같은 행의 오른쪽 어딘가에 빈 셀 존재 + rowspan=1 | 행 헤더 (병합 시 새 행 생성) |
| `gstub_` | 텍스트 있음 + 같은 행의 오른쪽 어딘가에 빈 셀 존재 + rowspan>1 | 그룹 행 헤더 (병합 시 셀 병합) |
| `input_` | 빈 셀 | 입력 필드 (병합 시 새 행 생성) |
| `data_` | 위 조건에 해당 안 되는 텍스트 셀 | 일반 데이터 |

## 적용 우선순위

1. `header_` - 행 전체가 배경색 있고, 연속된 행도 모두 배경색인 상단 행
2. `add_` - 긴 텍스트 (30자 이상)
3. `stub_` / `gstub_` - 같은 행의 오른쪽 어딘가에 빈 셀이 있는 텍스트 셀
4. `input_` - 빈 셀
5. `data_` - 나머지 텍스트 셀

## 배경색 판단 규칙

**배경색 있음**으로 인정되는 조건:
- 색상 값이 존재하고
- `none`, 빈 문자열이 아니며
- **흰색 계열이 아닌 경우**

**흰색 계열 (배경색 없음으로 처리)**:
- `#FFFFFF` (순수 흰색)
- RGB 값이 모두 220 이상인 색상 (예: `#FFFFE5`, `#F0F0F0`, `#FFFFF0`, `#E6E6E6` 등)

## 필드명 형식

모든 필드명은 8자리 랜덤 UUID 사용:
```
header_f92ba704
stub_3a2b1c4d
input_83b65a64
```

## stub_ / gstub_ 상세 규칙

**조건**: 텍스트가 있고, 같은 행의 오른쪽 어딘가에 빈 셀이 존재

- `stub_`: rowspan = 1 (단일 행)
- `gstub_`: rowspan > 1 (여러 행 병합)

### stub가 연속으로 존재하는 경우

stub 옆에 또 다른 stub가 있을 수 있음. 각 stub의 오른쪽 끝에 빈 셀이 있으면 모두 stub로 인식.

```
┌──────────┬──────────┬──────────┬──────────┬──────────┐
│ header_  │ header_  │ header_  │ header_  │ header_  │
├──────────┼──────────┼──────────┼──────────┼──────────┤
│ stub_A   │ stub_B   │ stub_C   │ input_1  │ input_2  │  ← stub 3개 연속
├──────────┼──────────┼──────────┼──────────┼──────────┤
│ gstub_D  │ stub_E   │ stub_F   │ input_1  │ input_2  │  ← gstub + stub 2개
│ (2행병합)│          │          │          │          │
├──────────┼──────────┼──────────┼──────────┼──────────┤
│          │ stub_G   │ stub_H   │ input_1  │ input_2  │  ← gstub 영역 + stub 2개
└──────────┴──────────┴──────────┴──────────┴──────────┘
```

- `stub_A`, `stub_B`, `stub_C`: 모두 오른쪽에 빈 셀(input_1, input_2)이 있으므로 stub
- `gstub_D`: rowspan=2이고 오른쪽에 빈 셀 있음 → gstub
- `stub_E`, `stub_F`, `stub_G`, `stub_H`: rowspan=1이고 오른쪽에 빈 셀 있음 → stub

## input_ 그룹화 규칙 (동일 필드명 조건)

빈 셀(`input_`)은 다음 조건을 **모두** 만족하면 동일 필드명 사용:

1. **위로 이동** 시 첫 번째 텍스트 셀(헤더)의 값이 동일
   - 단, `stub_`/`gstub_`는 헤더로 사용 불가 (건너뜀)
2. **좌측 조건** 중 하나 만족:
   - 좌측에 텍스트 셀이 없음
   - 좌측에 `gstub_`만 있음 (stub_ 없이)
   - 좌측에 `stub_`/`gstub_`만 있음 (일반 텍스트 셀 없음)
3. **셀 너비(colspan)가 동일**
   - 같은 열에서 위아래 빈 셀의 colspan이 같아야 함

### 그룹화 예시 1: 기본

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │  ← 배경색 있음
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │  ← col1, col2 각각 다른 이름
├──────────┼──────────┼──────────┤
│ gstub_Y  │ input_1  │ input_2  │  ← 위 헤더(B,C) 동일 + 좌측에 gstub만 → 같은 이름
│ (병합)   │          │          │
├──────────┼──────────┼──────────┤
│ stub_Z   │ input_1  │ input_2  │  ← 위 헤더 동일 + 좌측에 stub만 → 같은 이름
└──────────┴──────────┴──────────┘
```

- `(1,1)`, `(2,1)`, `(3,1)` → 모두 `input_1` (헤더 "B" 동일, 좌측에 stub/gstub만 있음)
- `(1,2)`, `(2,2)`, `(3,2)` → 모두 `input_2` (헤더 "C" 동일, 좌측에 stub/gstub만 있음)

### 그룹화 예시 2: colspan이 다른 경우

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │
├──────────┼──────────┴──────────┤
│ stub_X   │ input_1 (colspan=2) │  ← colspan=2
├──────────┼──────────┬──────────┤
│ stub_Y   │ input_2  │ input_3  │  ← colspan=1, 1
└──────────┴──────────┴──────────┘
```

- `(1,1)` → `input_1` (colspan=2)
- `(2,1)` → `input_2` (colspan=1, 헤더 "B" 동일하지만 colspan 다름 → 다른 이름)
- `(2,2)` → `input_3` (colspan=1)

### 그룹화 예시 3: 좌측에 일반 텍스트 셀이 있는 경우

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │
├──────────┼──────────┼──────────┤
│ stub_Y   │ data_Z   │ input_3  │  ← (2,1)에 텍스트 있음
└──────────┴──────────┴──────────┘
```

- `(1,1)` → `input_1`
- `(1,2)`, `(2,2)` → 다른 이름 (`(2,2)` 좌측에 일반 텍스트 셀 `data_Z`가 있음)

## 사용법

```bash
# 기본 사용 (기존 필드명 유지)
python -m merge.insert_auto_field input.hwpx output.hwpx

# 기존 필드명 무시하고 새로 생성
python -m merge.insert_auto_field input.hwpx output.hwpx --regenerate
```

## Python API

```python
from merge.insert_auto_field import insert_auto_fields

# 기본 사용
tables = insert_auto_fields("input.hwpx", "output.hwpx")

# 기존 필드명 무시하고 새로 생성
tables = insert_auto_fields("input.hwpx", "output.hwpx", regenerate=True)
```
