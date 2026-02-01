# nc_name 자동 생성 규칙

## 접두사 분류

| 접두사 | 조건 | 설명 |
|--------|------|------|
| `header_` | 행 전체가 배경색 있음 + 연속된 행도 모두 배경색 | 테이블 헤더 셀 |
| `add_` | (1x1 단일 셀 + 배경색 없음) 또는 (첫 데이터 행 + 30자 이상 + 배경색 없음) | 병합 시 내용만 추가 |
| `stub_` | 텍스트 있음 + 같은 행의 오른쪽 어딘가에 빈 셀 존재 + rowspan=1 | 행 헤더 (병합 시 새 행 생성) |
| `gstub_` | 텍스트 있음 + 같은 행의 오른쪽 어딘가에 빈 셀 존재 + rowspan>1 | 그룹 행 헤더 (병합 시 셀 병합) |
| `input_` | 빈 셀 | 입력 필드 (병합 시 새 행 생성) |
| `data_` | 위 조건에 해당 안 되는 텍스트 셀 | 일반 데이터 |

## 적용 우선순위

1. `header_` - 행 전체가 배경색 있고, 상단부터 연속된 행
2. `add_` - 1x1 단일 셀 또는 첫 데이터 행의 30자 이상 텍스트
3. `stub_` / `gstub_` - 오른쪽에 빈 셀이 있는 텍스트 셀
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
│ (2행병합) │          │          │          │          │
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
2. **좌측에 일반 텍스트 셀이 없음**
   - `stub_`/`gstub_`만 있거나 아무것도 없어야 함
3. **셀 너비(colspan)가 동일**
   - 같은 열에서 위아래 빈 셀의 colspan이 같아야 함
4. **좌측 stub/gstub 패턴이 동일**
   - 좌측의 stub/gstub 조합이 완전히 같아야 함
   - stub 하나라도 다르면 다른 그룹

### 그룹화 예시 1: stub 패턴이 같은 경우

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │  ← 배경색 있음
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │  ← stub_X 동일 → 같은 이름
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │  ← stub_X 동일 → 같은 이름
└──────────┴──────────┴──────────┘
```

- `(1,1)`, `(2,1)`, `(3,1)` → 모두 `input_1` (헤더 "B" 동일, stub_X 동일)
- `(1,2)`, `(2,2)`, `(3,2)` → 모두 `input_2` (헤더 "C" 동일, stub_X 동일)

### 그룹화 예시 2: stub 패턴이 다른 경우 (다른 그룹)

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │  ← 배경색 있음
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │
├──────────┼──────────┼──────────┤
│ stub_Y   │ input_3  │ input_4  │  ← stub_Y ≠ stub_X → 다른 이름!
├──────────┼──────────┼──────────┤
│ stub_Z   │ input_5  │ input_6  │  ← stub_Z ≠ stub_X, stub_Y → 다른 이름!
└──────────┴──────────┴──────────┘
```

- `(1,1)` → `input_1`, `(2,1)` → `input_3`, `(3,1)` → `input_5` (stub 다름 → 각각 다른 이름)
- `(1,2)` → `input_2`, `(2,2)` → `input_4`, `(3,2)` → `input_6` (stub 다름 → 각각 다른 이름)

### 그룹화 예시 3: gstub 병합 영역

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │
├──────────┼──────────┼──────────┤
│ gstub_D  │ input_1  │ input_2  │  ← gstub_D 시작
│ (2행병합)│          │          │
├──────────┼──────────┼──────────┤
│          │ input_1  │ input_2  │  ← gstub_D 영역 (같은 gstub) → 같은 이름
└──────────┴──────────┴──────────┘
```

- gstub 병합 영역 내의 input 셀들은 같은 gstub를 공유하므로 동일 필드명

### 그룹화 예시 4: colspan이 다른 경우

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │
├──────────┼──────────┴──────────┤
│ stub_X   │ input_1 (colspan=2) │  ← colspan=2
├──────────┼──────────┬──────────┤
│ stub_X   │ input_2  │ input_3  │  ← colspan=1, 1 (stub 같지만 colspan 다름)
└──────────┴──────────┴──────────┘
```

- `(1,1)` → `input_1` (colspan=2)
- `(2,1)` → `input_2` (colspan=1, stub_X 동일하지만 colspan 다름 → 다른 이름)
- `(2,2)` → `input_3` (colspan=1)

### 그룹화 예시 5: 좌측에 일반 텍스트 셀이 있는 경우

```
┌──────────┬──────────┬──────────┐
│ header_A │ header_B │ header_C │
├──────────┼──────────┼──────────┤
│ stub_X   │ input_1  │ input_2  │
├──────────┼──────────┼──────────┤
│ stub_X   │ data_Z   │ input_3  │  ← (2,1)에 일반 텍스트 있음
└──────────┴──────────┴──────────┘
```

- `(1,1)` → `input_1`
- `(1,2)` → `input_2`
- `(2,2)` → `input_3` (좌측에 일반 텍스트 셀 `data_Z`가 있음 → 그룹화 불가)

## 병합 동작

| 접두사 | 병합 시 동작 |
|--------|-------------|
| `header_` | 유지 (병합 안 함) |
| `add_` | 기존 셀에 내용 추가 |
| `stub_` | 새 행 생성 |
| `gstub_` | 셀 병합 (rowspan 확장) |
| `input_` | 빈 셀 채우기 또는 새 행 생성 |
| `data_` | 유지 |

## 사용법

```bash
# 기본 사용 (기존 필드명 유지)
python -m merge.field.insert_auto_field input.hwpx output.hwpx

# 기존 필드명 무시하고 새로 생성
python -m merge.field.insert_auto_field input.hwpx output.hwpx --regenerate

# 자동 필드명 + 배경색 적용 (권장)
python -m merge.field.auto_insert_field_template input.hwpx output.hwpx
```

## Python API

```python
from merge.field.insert_auto_field import insert_auto_fields

# 기본 사용
tables = insert_auto_fields("input.hwpx", "output.hwpx")

# 기존 필드명 무시하고 새로 생성
tables = insert_auto_fields("input.hwpx", "output.hwpx", regenerate=True)
```

### 빈 셀 필드명 채우기

행 추가 후 필드명이 없는 셀에 필드명을 자동으로 설정합니다.

```python
from merge.field.fill_empty import fill_empty_fields

result = fill_empty_fields("input.hwpx", "output.hwpx")
# result: {'tables': 5, 'rows_filled': 3, 'cells_filled': 12}
```
