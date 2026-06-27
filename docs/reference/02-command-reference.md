# 02. 명령어 레퍼런스

> 기준 소스 버전: `ccad021`
> 단일 출처(SSOT): 파라미터 정의는 `modCmdRegistry.bas`, 조건 연산자는 `modActions.bas`(`IsConditionOperator`), 키 이름은 `modActions.bas`(`ParseKeyCode`) 등록값을 그대로 따릅니다.
> 관련 문서: [시작하기](01-getting-started.md) · [작성 보조 도구](03-authoring-tools.md)

총 **5개 카테고리, 33개 명령어**. 각 표의 `P1/P2/P3`은 시트의 `PARAM1(C) / PARAM2(D) / PARAM3(E)` 열에 대응합니다.

---

## 표기 규칙

- **필수**: 반드시 입력해야 하는 파라미터. (옵션)은 생략 가능.
- **타입**: `num`(숫자) / `str`(문자).
- **RAW / RESOLVE**: 파라미터의 토큰 처리 방식.
  - **RESOLVE** (기본): `{변수}`·`{$토큰}`을 실제 값으로 치환한 뒤 사용.
  - **RAW**: 치환하지 않고 원문 그대로 사용. **딱 두 군데뿐** — `SET/INC/DEC`의 좌변(P1), `IF/WHILE`의 연산자(P2).
- 빈 칸은 해당 열을 비워둔다는 뜻입니다. 예: `KEY`는 P2를 사용하지 않습니다.

---

## A. 흐름 제어 (FLOW) — 16개

분기, 반복, 라벨/점프, 서브루틴을 담당합니다. 블록 명령은 시작–종료가 짝을 이루며 자유롭게 중첩됩니다.

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `IF` | 좌변 (str, 필수) | 연산자 (str, **RAW**, 필수) | 우변 (str, 필수) | 조건이 참이면 다음 줄 실행, 거짓이면 `ELSE`(있으면) 또는 `END_IF`로 점프 |
| `ELSE` | — | — | — | `IF`의 거짓 분기 (선택) |
| `END_IF` | — | — | — | `IF` 블록 종료 |
| `WHILE` | 좌변 (str, 필수) | 연산자 (str, **RAW**, 필수) | 우변 (str, 필수) | 조건이 참인 동안 블록 반복 |
| `END_WHILE` | — | — | — | `WHILE` 블록 종료 (헤더로 복귀) |
| `REPEAT` | 반복횟수 (num, 옵션) | — | — | N회 반복. **0 또는 빈값이면 무한 반복** |
| `END_REPEAT` | — | — | — | `REPEAT` 블록 종료 |
| `SELECT` | 대상값 (str, 필수) | — | — | 다중 분기. 일치하는 `CASE`로 이동 |
| `CASE` | 비교값 (str, 옵션) | — | — | `SELECT` 안의 분기 라벨. 빈값/`ELSE`는 기본값 |
| `END_SELECT` | — | — | — | `SELECT` 블록 종료 |
| `BREAK` | — | — | — | 최내곽 루프(`WHILE`/`REPEAT`) 탈출 |
| `CONTINUE` | — | — | — | 최내곽 루프의 다음 반복으로 |
| `GOTO` | — | — | 라벨/줄번호 (str, 필수) | 지정 라벨(대소문자 무시) 또는 줄번호로 무조건 점프 |
| `CALL` | — | — | 라벨/줄번호 (str, 필수) | 서브루틴 호출. 복귀 주소를 스택에 저장 |
| `RETURN` | — | — | — | `CALL` 호출부로 복귀 (스택이 비면 오류) |
| `LABEL` | — | — | 라벨명 (str, 필수) | 점프 대상 라벨 정의. 실행 전 사전 스캔됨 |

> `GOTO`·`CALL`·`LABEL`의 인자는 **P3(E열)**에 들어갑니다 (P1·P2는 비움).

**블록 짝**: `IF ↔ END_IF`, `WHILE ↔ END_WHILE`, `REPEAT ↔ END_REPEAT`, `SELECT ↔ END_SELECT`. 모든 블록이 짝을 이뤄야 하며 중첩 가능합니다.

**루프 토큰**: 루프 안에서 `{$LOOP}`은 현재 반복 횟수(1부터)를 반환합니다.

### 예제

```
# IF / ELSE
IF | {결과} | EQ | 성공
  CLICK | 100 | 100
ELSE
  CLICK | 200 | 200
END_IF

# WHILE 루프
SET | {i} | 0
WHILE | {i} | < | 5
  CLICK | 100 | {i}
  INC | {i}
END_WHILE

# REPEAT (5회) + 루프 카운터
REPEAT | 5
  NOTIFY |  |  | 반복 {$LOOP}회차|1
END_REPEAT

# SELECT / CASE
SELECT | {모드}
  CASE | A
    CLICK | 100 | 100
  CASE | ELSE
    CLICK | 300 | 300
END_SELECT

# 서브루틴
CALL | 인사하기
EXIT
LABEL | 인사하기
  TYPE | 0 |  | 안녕하세요
RETURN
```

---

## B. 동작 (ACTION) — 6개

마우스·키보드 입력. 좌표는 **화면 픽셀**(0,0 = 좌상단)입니다.

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `MOVE` | x (num, 필수) | y (num, 필수) | 사후대기ms (num, 옵션) | 커서를 (x, y)로 이동 |
| `CLICK` | x (num, 필수) | y (num, 필수) | 사후대기ms (num, 옵션) | (x, y) 좌클릭 |
| `RCLICK` | x (num, 필수) | y (num, 필수) | 사후대기ms (num, 옵션) | (x, y) 우클릭 |
| `KEY` | 키코드/이름 (str, 필수) | — | 사후대기ms (num, 옵션) | 단일 키 입력 |
| `COMBO` | key1 (str, 필수) | key2 (str, 필수) | 사후대기ms (num, 옵션) | 조합 키 (예: Ctrl+C) |
| `TYPE` | 딜레이ms (num, 옵션) | — | 텍스트 (str, 필수) | 텍스트를 문자당 딜레이로 입력 |

- **사후대기ms**: 동작 후 추가 대기 시간(밀리초). 대기 중 `Esc`로 중단 가능.
- `TYPE`의 딜레이가 0 이하이면 안정성을 위해 **문자당 50ms**가 적용됩니다. 텍스트는 **P3(E열)**에 입력합니다.

### 키 이름 (KEY / COMBO)

숫자 VK 코드(예: `13`) 또는 아래 이름을 사용합니다 (`ParseKeyCode` 기준):

```
ENTER/RETURN(13)  ESC/ESCAPE(27)  TAB(9)   CTRL/CONTROL(17)  ALT/MENU(18)
SHIFT(16)  WIN/WINDOWS(91)  SPACE/SPC(32)  BS/BACKSPACE(8)   DEL/DELETE(46)
INS/INSERT(45)  HOME(36)  END(35)  PGUP/PAGEUP(33)  PGDN/PAGEDOWN(34)
UP(38)  DOWN(40)  LEFT(37)  RIGHT(39)  PRTSC(44)  CAPS(20)  F1~F12
```

### 예제

```
MOVE  | 100 | 200 | 500       # (100,200)로 이동 후 0.5초 대기
CLICK | 150 | 250             # 좌클릭
RCLICK| 160 | 260             # 우클릭
KEY   | ENTER                 # 엔터
KEY   | F5    |   | 300       # F5 누르고 0.3초 대기
COMBO | CTRL  | C             # Ctrl+C (복사)
TYPE  | 50    |   | 안녕하세요  # 50ms 간격으로 입력
```

---

## C. 메모리 (MEMORY) — 3개

사용자 변수에 값을 저장·증감합니다.

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `SET` | 좌변/주소 (str, **RAW**, 필수) | — | 값 (str, 필수) | 변수에 값 대입 |
| `INC` | 좌변/주소 (str, **RAW**, 필수) | — | 증분 (num, 옵션, 기본 1) | 숫자 변수 증가 |
| `DEC` | 좌변/주소 (str, **RAW**, 필수) | — | 감분 (num, 옵션, 기본 1) | 숫자 변수 감소 |

- **좌변(P1)은 RAW**이며 반드시 `{ }`로 감싸야 합니다. 허용 대상:
  - `{변수명}` — 사용자 변수 (대소문자 무시)
  - `{$CLIP}` — 클립보드 (쓰기 대상으로 허용)
  - 읽기 전용 시스템 토큰(`{$LOOP}` `{$PIXEL}` `{$BLANK}` 등)은 좌변으로 쓸 수 없습니다(오류).
- `INC`/`DEC`는 존재하지 않는 변수를 **0**으로 간주하고 계산합니다.

### 예제

```
SET | {카운터} | 0          # 카운터 = 0
SET | {이름}   | 홍길동      # 이름 = 홍길동
SET | {복사본} | {$CLIP}     # 클립보드 내용을 변수로
SET | {$CLIP}  | {이름}      # 변수 값을 클립보드로
INC | {카운터}               # +1 (기본)
INC | {카운터} | 5           # +5
DEC | {카운터} | 2           # -2
```

---

## D. 시스템 (SYSTEM) — 4개

창 제어, 대기, 알림, 종료.

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `FOCUS` | — | — | 창제목 (str, 필수) | 제목으로 창을 전면화. 미발견 시 오류, 성공 시 0.5초 대기 |
| `SLEEP` | — | — | 대기ms (num, 필수) | 지정 시간 대기 (`Esc`로 중단 가능) |
| `NOTIFY` | — | — | 메시지[\|초] (str, 필수) | 토스트 알림 표시 |
| `EXIT` | — | — | — | 스크립트를 정상 종료 |

- `FOCUS`·`SLEEP`·`NOTIFY`의 인자는 **P3(E열)**에 입력합니다.
- `NOTIFY` 형식: `메시지` 또는 `메시지|초`. 초를 생략하거나 0 이하면 **기본 2초** 표시됩니다.

### 예제

```
FOCUS  |  |  | 메모장          # "메모장" 창을 앞으로
SLEEP  |  |  | 1000           # 1초 대기
NOTIFY |  |  | 작업 완료|3      # 3초간 알림
NOTIFY |  |  | 저장됨          # 2초(기본) 알림
EXIT                          # 종료
```

---

## E. 비전 (VISION) — 4개

픽셀 색상을 기준으로 화면과 동기화하거나 결과를 검증합니다. 색상은 16진수 `#RRGGBB` 형식입니다.

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `WAIT_APPEAR` | x (num, 필수) | y (num, 필수) | hexColor (str, 필수) | (x,y) 픽셀이 지정 색이 될 때까지 대기 |
| `WAIT_VANISH` | x (num, 필수) | y (num, 필수) | hexColor (str, 필수) | (x,y) 픽셀이 지정 색에서 벗어날 때까지 대기 |
| `CLICK_SYNC` | x (num, 필수) | y (num, 필수) | — | (x,y) 좌클릭 후 색 변화(반응)를 확인하고 진행 |
| `ASSERT` | x (num, 필수) | y (num, 필수) | 조건 (str, 필수) | (x,y) 픽셀 색을 검증 |

- `WAIT_APPEAR`·`WAIT_VANISH`·`CLICK_SYNC`는 **타임아웃 30초**입니다. 시간 내 조건 미충족 시 오류로 중단됩니다 (`Esc`로도 중단 가능).
- `ASSERT`의 조건(P3)은 `!#RRGGBB` 형식으로 **"이 색이 아니어야 함"**을 검사합니다. 실패 시 오류.

### 예제

```
WAIT_APPEAR | 100 | 200 | #FF0000   # (100,200)이 빨강이 될 때까지 대기
WAIT_VANISH | 150 | 250 | #00FF00   # (150,250)이 초록에서 바뀔 때까지 대기
CLICK_SYNC  | 200 | 300             # 클릭하고 반응 확인
ASSERT      | 50  | 50  | !#FFFFFF  # (50,50)이 흰색이면 실패
```

---

## 조건 연산자 (IF / WHILE 의 P2)

`IF`/`WHILE`의 연산자(P2)는 **RAW**로 처리되며, 아래 목록만 허용됩니다 (`IsConditionOperator` 기준).

| 연산자 | 의미 |
|--------|------|
| `EQ` / `==` | 정확히 같음 (`*` `?`는 와일드카드가 아닌 리터럴로 취급) |
| `!=` | 다름 |
| `>` `<` `>=` `<=` | 수치 비교 |
| `LIKE` | 와일드카드 패턴 — `*`(임의 길이) `?`(한 글자) `#`(한 숫자) |
| `STARTS` | 좌변이 우변으로 시작 |
| `ENDS` | 좌변이 우변으로 끝 |
| `CONTAINS` | 좌변이 우변을 포함 |

> ⚠️ `NE`, `LT`, `LE`, `GT`, `GE`, `IN`, `NOTCONTAIN` 등은 **존재하지 않는 연산자**입니다. 사용하면 검증 오류로 표시됩니다.

```
IF | {횟수}   | >        | 10
IF | {파일명} | ENDS     | .xlsx
IF | {상태}   | CONTAINS | 완료
IF | {코드}   | LIKE     | A##         # A 뒤 숫자 2자리
```

---

## 토큰·변수 시스템

RESOLVE 슬롯의 파라미터는 실행 시 `{ }` 토큰이 실제 값으로 치환됩니다.

### 고정 토큰 (정적 문자)

| 토큰 | 값 |
|------|-----|
| `{$BLANK}` | 빈 문자열 |
| `{$TAB}` | 탭 문자 |
| `{$CR}` | 캐리지 리턴 |
| `{$LF}` | 라인 피드 |
| `{$SPACE}` | 공백 한 칸 |

### 동적 토큰 (런타임 값)

| 토큰 | 값 |
|------|-----|
| `{$CLIP}` | 현재 클립보드 텍스트 |
| `{$LOOP}` | 최내곽 루프의 현재 반복(1부터), 루프 밖이면 빈값 |
| `{$PIXEL:x,y}` | (x,y) 픽셀 색상 `#RRGGBB` |

### 사용자 변수

`{변수명}` — `SET`/`INC`/`DEC`로 만든 변수. 대소문자를 구분하지 않으며, 미정의 변수는 빈 문자열로 읽힙니다.

### 중첩 해석

토큰 안에 토큰을 넣을 수 있습니다(재귀 치환):

```
SET | {x} | 100
SET | {y} | 200
IF  | {$PIXEL:{x},{y}} | EQ | #FFFFFF     # (100,200) 색이 흰색이면
  NOTIFY |  |  | 흰색 감지
END_IF
```

---

다음: 포맷터·실시간 검증·단축키 등 작성을 돕는 도구는 **[03. 작성 보조 도구](03-authoring-tools.md)** 를 참고하세요.

---

## FOREACH_ROW data iteration

`FOREACH_ROW` reads a worksheet or range from an open workbook or read-only external workbook into a memory snapshot, then loops over each data row. The first row is treated as headers, and iteration starts at the next row.

| Command | P1 (C) | P2 (D) | P3 (E) | Description |
|--------|--------|--------|--------|-------------|
| `FOREACH_ROW` | workbook name/path (str, RESOLVE, required) | sheet or `sheet!address` (str, RESOLVE, required) | mode: `ALL_ROWS`/`VISIBLE_ROWS` (optional) | Start table-row iteration |
| `END_FOREACH` |  |  |  | Advance to the next row or finish the loop |

- If P1 contains `\`, `/`, or `:`, it is treated as a path and opened read-only. If the same file name is already open, the open workbook is reused and not closed.
- If P1 is not a path, open workbook names are searched by case-insensitive partial match. Zero or multiple matches stop with an error.
- P2 supports `Sheet1`, `Sheet1!A1`, and `Sheet1!A1:C10`.
- P3 blank or `ALL_ROWS` loops over all data rows.
- P3 `VISIBLE_ROWS` loops over the header plus currently visible data rows only. Rows hidden by filters or manual row hiding are skipped; hidden columns are still loaded.
- Large explicit ranges are clamped to the worksheet's actual used range.
- Cell values are read from `.Text` to preserve visible strings. Narrow-column `####` cells are reconstructed from `Value` and `NumberFormat`.

### FOREACH_ROW tokens

| Token | Value |
|------|-------|
| `{$ROW.column}` | Current row value for the matching header |
| `{$RIDX}` | Current `FOREACH_ROW` iteration number, 1-based |

`{$LOOP}` keeps its existing meaning: the innermost active loop iteration. Inside a `WHILE` or `REPEAT` nested within `FOREACH_ROW`, use `{$RIDX}` for the row index.
