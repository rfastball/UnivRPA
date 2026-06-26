# UnivRPA Script Program Model 설계 스펙 (v3.2)

> **개정 이력**
>
> **v3.2**
>
> v3.1 리뷰 반영판.
>
> ① SELECT 구조는 1차 구현부터 **재스캔이 아니라 단일 인덱스 조회**로 고정한다.
> ② `LABEL`은 **정적 심볼**로 확정한다. `LABEL` 선언부에는 토큰 문법을 허용하지 않는다.
> ③ `FindBlockPair`의 시작 줄·방향별 반환 계약을 예시로 명문화한다.
> ④ `LABEL`의 장기 ParamSpec 방향을 `RESOLVE`에서 `RAW`로 전환 검토 대상으로 명시한다.
> ⑤ `LINE_NOT_FOUND = CMD_ERROR(-1)` 계열 sentinel은 값은 공유하되 의미 이름을 분리한다.
>
> **v3.1**
>
> SELECT/CASE 구조를 **단일 인덱스**로 확정하고 문서를 lean화.
>
> ① SELECT 디스패치(`ResolveSelectTarget`)와 CASE 폴스루(`FindBlockPair`)의 두 전방 스캐너를 하나의 구조 인덱스로 통합.
> ② 폴스루의 `FindBlockPair` 의존 제거. `EndSelectOfCase(caseLine)` 신설.
> ③ `FindBlockPair`의 MIDBLOCK 즉시반환을 `BlockGroup=IF`로 한정.
> ④ sentinel 명시. `EndSelectLine = -1`, 미존재 라벨 `ResolveJumpTarget = -1`은 `CMD_ERROR(-1)`와 의도적으로 동치.
> ⑤ SELECT 폴백 우선순위 `first-match / last-fallback` 보존 명문화.

---

# 1. 스코프

`modEngine`의 `m_arrMain`은 단순 캐시가 아니다.

현재 구조에서 `m_arrMain`은 Excel `CurrentRegion`에서 헤더를 제외한 2차원 배열이며, 다음 기능의 공통 기반이다.

```text
명령어 조회
파라미터 조회
라벨 검색
점프 대상 해석
블록 탐색
SELECT / CASE 탐색
런타임 줄 번호 ↔ 시트 행 번호 매핑
```

따라서 `m_arrMain`은 로드된 스크립트 프로그램 이미지, 즉 **ProgramImage**로 취급한다.

```text
Before:
  Excel Sheet
    → Variant Array
      → modEngine 실행

After:
  Excel Sheet
    → clsScriptProgram 로드/분석
      → modEngine 실행
```

본 스펙은 이 ProgramImage를 `clsScriptProgram`으로 승격한다.

단, 새 언어 분석기를 별도로 만들지 않는다. 이미 1급 시민화된 `modCmdRegistry`의 `CmdDef`, `BlockRole`, `BlockGroup`, `ParamSpec`을 구조 의미의 진실원천으로 소비한다.

## 1.1 목표

```text
배열 접근 캡슐화
숫자 열 인덱스 제거
구조 분석을 Registry 기반으로 단일화
FLOW 핸들러의 modEngine 역질의 결합 완화
정규화 단일 출처 마련
정적 검증·실행·편집 보조의 공통 모델 기반 마련
```

## 1.2 비목표

```text
명령어 문법 변경 없음
시트 구조 변경 없음: COMMAND / PARAM1 / PARAM2 / PARAM3 유지
Registry 구조 변경 최소화
토큰 문법 변경 없음
WinAPI / Action 계층 변경 없음
AST / Compiler 전면 도입 아님
ICommand 전면 변경은 1차 제외
```

이번 작업은 “언어를 새로 만든다”가 아니다.

이미 존재하는 Excel-hosted DSL 런타임에서 **로드된 프로그램 이미지의 소유권과 책임을 분리**하는 작업이다.

---

# 2. 책임 분리

```text
clsScriptProgram
  무엇이 적혀 있는가
  구조의 위치가 어디인가
  줄 번호와 시트 행 번호는 어떻게 대응되는가

modCmdRegistry
  명령어가 무엇을 뜻하는가
  Category / BlockRole / BlockGroup / ParamSpec
  구조 의미의 진실원천

modEngine
  어떻게 실행하는가
  다음 IP는 어디인가
  런타임 값은 어떻게 해석하는가
  SELECT 값 비교와 폴백은 어떻게 처리하는가

modContext
  런타임 상태
  Memory / CallStack / LoopStack

Command Handlers
  각 Category별 실행 의미

modActions
  현실 세계에 대한 효과

modWinAPI
  OS와의 대화
```

핵심 책임선은 다음과 같다.

```text
구조 의미 = Registry
구조 위치 = Program
런타임 값·폴백 = Engine
실행 상태 = Context
현실 효과 = Actions / WinAPI
```

Program은 구조 위치를 계산한다.
Registry는 명령어의 의미를 정의한다.
Engine은 실행 시점의 값을 해석한다.
Context는 실행 중 상태를 보관한다.

어느 하나가 다른 계층의 책임을 흡수하지 않는다.

---

# 3. 핵심 설계 결정

## 3.1 `m_arrMain`은 ProgramImage다

시트의 DSL은 **Script**이고, 실행 시점에 메모리에 올라온 이미지는 **Program**이다.

권장 클래스명은 다음과 같다.

```text
clsScriptProgram
```

이 클래스는 기존 `m_arrMain`을 내부 구현으로 숨기고, 의미 기반 접근자와 구조 질의 API를 제공한다.

---

## 3.2 배열은 내부 구현으로 숨긴다

외부 모듈은 더 이상 아래와 같은 접근을 하지 않는다.

```vba
m_arrMain(R, 2)
m_arrMain(R, 3)
m_arrMain(R, 4)
m_arrMain(R, 5)
```

대신 다음 API를 사용한다.

```text
program.CmdAt(R)
program.P1At(R)
program.P2At(R)
program.P3At(R)
program.RawCmdAt(R)
```

열 상수는 `clsScriptProgram` 내부에 둔다.

```vba
Private Const COL_ROW As Long = 1
Private Const COL_CMD As Long = 2
Private Const COL_P1  As Long = 3
Private Const COL_P2  As Long = 4
Private Const COL_P3  As Long = 5
```

외부 모듈은 “몇 번째 열인가”가 아니라 “무슨 의미인가”만 본다.

---

## 3.3 구조 분석은 Program 책임, 구조 의미는 Registry 책임

Program은 다음 질문에 답한다.

```text
이 줄의 명령어는 무엇인가?
이 줄은 빈 줄인가?
이 줄은 주석인가?
이 IF의 ELSE 또는 END_IF는 어디인가?
이 WHILE의 END_WHILE은 어디인가?
이 CASE가 속한 END_SELECT는 어디인가?
이 라벨은 몇 번째 런타임 줄인가?
이 런타임 줄은 실제 시트 몇 행인가?
```

그러나 Program은 명령어 의미를 하드코딩하지 않는다.

블록 판정은 Registry에서 가져온다.

```text
CmdDef.BlockRole
CmdDef.BlockGroup
```

즉, Program은 `IF`, `END_IF`, `CASE` 같은 문자열에 직접 의존하지 않는다.
그 대신 `GetCmdDef(cmd)`를 통해 `STARTER`, `MIDBLOCK`, `ENDER`, `BlockGroup`을 읽는다.

---

## 3.4 토큰 해석은 Program 책임이 아니다

다음 토큰들은 실행 시점 상태에 의존한다.

```text
{name}
{$CLIP}
{$LOOP}
{$PIXEL:x,y}
{$BLANK}
{$TAB}
```

따라서 Program은 토큰을 해석하지 않는다.

Program은 원문 파라미터만 제공한다.

```text
program.P1At(R)
program.P2At(R)
program.P3At(R)
```

토큰 해석은 `modEngine.ResolveParam`, `ResolveToken`, `ResolvePixelToken`에 남긴다.

장기적으로 필요하면 `modTokenResolver`로 분리할 수 있으나, 1차 작업 범위에는 포함하지 않는다.

---

## 3.5 LABEL은 정적 심볼이다

라벨은 런타임 값이 아니라 프로그램의 정적 심볼이다.

따라서 `LABEL`은 Program 로드/분석 단계에서 인덱싱한다.

```text
LABEL | | | LOGIN
GOTO  | | | LOGIN
CALL  | | | DO_WORK
```

Program은 다음 인덱스를 만든다.

```text
LabelName → RuntimeLine
```

중복 라벨은 안전한 검증 실패로 처리한다.

기존처럼 `modContext.AddLabel`에서 `End`로 VBA 런타임을 강제 종료하지 않는다.

## 3.5.1 LABEL 선언부의 토큰 금지

`LABEL` 선언 자체에는 토큰 문법을 허용하지 않는다.

금지 예:

```text
LABEL | | | {target}
LABEL | | | {$CLIP}
LABEL | | | STEP_{n}
```

이유는 단순하다.

```text
라벨 인덱스는 실행 전에 만들어야 한다.
토큰은 실행 중 상태를 해석해야 알 수 있다.
```

따라서 둘은 같은 위치에 둘 수 없다.

`LABEL`은 반드시 정적 문자열이어야 한다.

## 3.5.2 GOTO/CALL의 동적 대상은 Engine 정책으로 유지 가능

`LABEL` 선언은 정적이어야 하지만, `GOTO`/`CALL`의 대상은 Engine에서 토큰 해석 후 조회할 수 있다.

예:

```text
SET   | {next} | | LOGIN
GOTO  |        | | {next}
LABEL |        | | LOGIN
```

실행 흐름:

```text
1. Engine이 GOTO P3의 토큰을 해석한다.
2. 해석 결과가 숫자면 줄 번호로 사용한다.
3. 해석 결과가 문자열이면 Program.LabelLine(name)에서 조회한다.
4. 없으면 CMD_ERROR(-1)를 반환한다.
```

즉:

```text
LABEL 선언 = 정적
GOTO/CALL 대상 = 런타임 해석 가능
최종 조회 = Program의 정적 라벨 인덱스
```

## 3.5.3 장기 Registry 조정 방향

현재 `LABEL`의 P3가 `RESOLVE`라면, 의미상 장기적으로 `RAW`가 더 적합하다.

권장 방향:

```vba
Reg "LABEL", FLOW, NOBLOCK, "", clrFLOW, FULLROW, _
    sp3:=P("라벨명", DT_STR, RAW, True)
```

단, 1차 구현에서는 기존 Registry 변경 범위를 최소화할 수 있다.
이 경우 Program의 LabelIndex 구축 단계에서 `LABEL P3`를 강제로 RAW 정적 심볼로 취급한다.

---

## 3.6 FLOW 핸들러는 Engine이 아니라 ActiveProgram에 구조를 묻는다

기존 FLOW 핸들러는 다음과 같이 `modEngine`에 구조를 역질의한다.

```text
modEngine.FindBlockPair
modEngine.CmdAt
modEngine.ResolveSelectTarget
```

1차에서는 ICommand 시그니처를 바꾸지 않는다.

대신 `modEngine.ActiveProgram()`을 둔다.

```vba
Public Function ActiveProgram() As clsScriptProgram
    Set ActiveProgram = m_program
End Function
```

FLOW 핸들러는 다음처럼 Program을 통해 구조를 조회한다.

```vba
Dim program As clsScriptProgram
Set program = modEngine.ActiveProgram()

targetLine = program.FindBlockPair(R, "IF", 1)
```

장기적으로는 Flow Handler에 Program을 주입하는 방식으로 전환할 수 있다.

---

## 3.7 블록 구조의 진실원천은 Registry다

블록 시작, 끝, 중간 역할은 Registry의 `CmdDef`가 가진다.

```text
BlockRole:
  STARTER
  ENDER
  MIDBLOCK
  NOBLOCK

BlockGroup:
  IF
  WHILE
  REPEAT
  SELECT
  ""
```

Program의 구조 탐색기는 문자열을 직접 비교하지 않는다.

나쁜 방향:

```vba
If cmd = "IF" Then depth = depth + 1
If cmd = "END_IF" Then depth = depth - 1
```

좋은 방향:

```vba
Set def = modCmdRegistry.GetCmdDef(cmd)

If def.BlockRole = STARTER And def.BlockGroup = group Then
    depth = depth + 1
End If
```

이렇게 하면 새 블록 명령이 추가될 때 Registry 한 곳만 갱신하면 된다.

---

# 4. SELECT 구조 설계

## 4.1 SELECT는 단일 구조 인덱스를 사용한다

SELECT는 `FindBlockPair`의 변종으로 다루지 않는다.

SELECT는 다음 두 기능이 필요하다.

```text
1. SELECT 줄에서 실행할 CASE를 찾는다.
2. CASE 줄에 폴스루로 도달했을 때 END_SELECT로 건너뛴다.
```

기존 구조에서는 이 둘이 서로 다른 스캐너를 탔다.

```text
SELECT 실행:
  ResolveSelectTarget이 전방 스캔

CASE 폴스루:
  FindBlockPair(R, "SELECT", 1)이 전방 스캔
```

v3.2에서는 둘을 하나의 구조 인덱스로 통합한다.

Program 로드/분석 단계에서 다음 정보를 만든다.

```text
SelectLine → SelectBlock
CaseLine   → EndSelectLine
```

`SelectBlock`은 다음만 가진다.

```text
SelectLine
CaseLines()
EndSelectLine
```

값 비교, fallback, 토큰 해석은 담지 않는다.

---

## 4.2 SELECT 인덱스 구축 규칙

SELECT 구조 인덱스는 Registry의 구조 역할만 본다.

```text
STARTER + BlockGroup=SELECT
  → 중첩 SELECT depth 증가

MIDBLOCK + BlockGroup=SELECT + depth=0
  → 현재 SELECT 블록의 CASE 줄로 기록

ENDER + BlockGroup=SELECT
  → depth=0이면 현재 SELECT 블록의 END_SELECT
  → depth>0이면 중첩 SELECT 종료로 보고 depth 감소

그 외 IF / WHILE / REPEAT
  → SELECT depth에 영향 없음
```

핵심은 `BlockGroup=SELECT` 게이트다.

CASE 본문 안에 중첩 IF, WHILE, REPEAT가 있어도 SELECT depth를 건드리지 않는다.
중첩 SELECT의 CASE만 depth로 배제한다.

---

## 4.3 SELECT는 1차 구현부터 인덱스 조회로 고정한다

중요한 결정:

```text
FindBlockPair(IF/WHILE/REPEAT)는 1차에서 호출 시 재스캔 가능하다.
그러나 SELECT는 1차부터 단일 인덱스를 사용한다.
```

즉, 다음 API는 재스캔하지 않는다.

```text
SelectBlock(selectLine)
EndSelectOfCase(caseLine)
```

두 API는 Program 로드/분석 단계에서 만든 SELECT 인덱스를 조회한다.

허용:

```text
IF/WHILE/REPEAT FindBlockPair = 1차 재스캔 가능
```

불허:

```text
SELECT SelectBlock = 호출 시 재스캔
CASE EndSelectOfCase = 호출 시 재스캔
```

SELECT는 이번 설계의 주요 회귀 방지 대상이므로, 캐싱을 P7로 미루지 않는다.

---

## 4.4 SELECT 실행 책임 분리

Program은 구조만 반환한다.

```text
CASE 줄 목록
END_SELECT 줄
```

Engine은 실행 의미를 처리한다.

```text
SELECT 대상값 해석
CASE 값 토큰 해석
값 비교
fallback 선택
최종 점프 줄 결정
```

## 4.5 SELECT 동작 보존 규칙

기존 동작은 다음처럼 보존한다.

```text
first-match 우선
  일치 CASE를 만나면 즉시 그 다음 줄로 점프한다.

last-fallback 우선
  CASE ELSE 또는 빈 CASE를 만날 때마다 fallback 후보를 갱신한다.
  마지막 fallback이 이긴다.

무일치·무폴백
  END_SELECT 줄 자체로 점프한다.
  END_SELECT 실행이 다음 줄로 진행한다.

END_SELECT 누락
  EndSelectLine = -1
  Engine은 CMD_ERROR(-1)로 처리한다.
```

예:

```text
SELECT | status
CASE   | A
TYPE   | alpha
CASE   | ELSE
TYPE   | fallback
END_SELECT
```

실행 흐름:

```text
status = A
  → CASE A 다음 줄로 점프

status = Z
  → CASE ELSE 다음 줄로 점프

CASE 본문 실행 후 다음 CASE에 도달
  → EndSelectOfCase(caseLine) 조회
  → END_SELECT로 점프
```

---

# 5. FindBlockPair 설계

## 5.1 FindBlockPair는 SELECT를 다루지 않는다

`FindBlockPair`는 다음 그룹만 허용한다.

```text
IF
WHILE
REPEAT
```

다음 호출은 거부한다.

```text
FindBlockPair(..., "SELECT", ...)
```

반환:

```text
-1
```

SELECT는 전용 API를 사용한다.

```text
SelectBlock(selectLine)
EndSelectOfCase(caseLine)
```

---

## 5.2 MIDBLOCK 즉시반환은 IF 전용이다

`BlockRole=MIDBLOCK`을 무조건 즉시반환하면 위험하다.

현재는 `ELSE`와 `CASE`가 MIDBLOCK이다.

하지만 미래에 WHILE/REPEAT 계열 MIDBLOCK이 추가될 수 있다.

따라서 다음 규칙을 둔다.

```text
depth=0에서 MIDBLOCK 즉시반환은 BlockGroup=IF일 때만 허용한다.
```

즉:

```text
IF의 ELSE
  → 즉시반환 가능

SELECT의 CASE
  → FindBlockPair 대상 아님

WHILE/REPEAT의 미래 MIDBLOCK
  → 즉시반환 금지
```

---

## 5.3 FindBlockPair 시작 줄·방향 계약

기존 동작 보존을 위해 `FindBlockPair`는 현재 줄을 다시 보지 않는다.

탐색 시작 위치:

```text
Direction = 1
  startLine + 1부터 아래로 탐색

Direction = -1
  startLine - 1부터 위로 탐색
```

## 5.4 FindBlockPair 계약 예시

다음 예시는 구현 계약이다.

### IF 시작 줄에서 찾기

```text
1 IF
2   TYPE
3 ELSE
4   TYPE
5 END_IF
```

```text
FindBlockPair(1, "IF", 1) = 3
```

같은 depth의 `ELSE`를 우선 반환한다.
`ELSE`가 없으면 `END_IF`를 반환한다.

### ELSE 줄에서 찾기

```text
1 IF
2   TYPE
3 ELSE
4   TYPE
5 END_IF
```

```text
FindBlockPair(3, "IF", 1) = 5
```

`ELSE` 이후에는 같은 IF 블록의 `END_IF`를 반환한다.

### IF 안에 중첩 IF가 있는 경우

```text
1 IF
2   IF
3   ELSE
4   END_IF
5 ELSE
6 END_IF
```

```text
FindBlockPair(1, "IF", 1) = 5
FindBlockPair(2, "IF", 1) = 3
FindBlockPair(5, "IF", 1) = 6
```

중첩 IF의 ELSE는 바깥 IF의 ELSE로 오인하지 않는다.

### WHILE 정방향

```text
1 WHILE
2   TYPE
3 END_WHILE
```

```text
FindBlockPair(1, "WHILE", 1) = 3
```

### WHILE 역방향

```text
1 WHILE
2   TYPE
3 END_WHILE
```

```text
FindBlockPair(3, "WHILE", -1) = 1
```

### REPEAT 정방향

```text
1 REPEAT
2   TYPE
3 END_REPEAT
```

```text
FindBlockPair(1, "REPEAT", 1) = 3
```

### REPEAT 역방향

```text
1 REPEAT
2   TYPE
3 END_REPEAT
```

```text
FindBlockPair(3, "REPEAT", -1) = 1
```

### SELECT 거부

```text
FindBlockPair(1, "SELECT", 1) = -1
FindBlockPair(8, "SELECT", -1) = -1
```

SELECT는 `FindBlockPair`의 책임이 아니다.

---

# 6. 줄 접근과 정규화

## 6.1 줄 범위 검사는 Program 책임이다

기존 배열 접근은 범위 밖 오류를 빈 문자열처럼 삼킬 수 있다.

이는 다음 문제를 만든다.

```text
범위 밖 줄 번호
  → 빈 줄로 오판
  → 조용히 skip
  → 오프바이원 은폐
```

따라서 Program은 줄 번호를 검사한다.

```text
1 <= runtimeLine <= TotalLines
```

범위 밖 접근은 검출 가능한 실패로 처리한다.

---

## 6.2 정규화 단일 규칙

`CmdAt` 정규화 규칙:

```text
Trim
">" 제거
Trim
UCase
```

예:

```text
">>  click " → "CLICK"
```

단, 빈 줄과 주석 판정은 기존 동작 보존을 위해 `>` 제거 전의 Trim 원문 기준으로 고정한다.

```text
Trim(raw) = ""
  → blank

Left(Trim(raw), 1) = "#"
  → comment
```

Program은 다음 API를 제공한다.

```text
IsBlankLine(line)
IsCommentLine(line)
IsSkippable(line)
CmdAt(line)
RawCmdAt(line)
```

장기적으로 `modDevTools`, `modScriptAssist`의 정규화 복사본도 Program 또는 공용 Normalizer로 수렴한다.

---

# 7. Sentinel 정책

실행 반환 코드에는 기존 상수를 유지한다.

```vba
Public Const CMD_NEXT_LINE As Long = 0
Public Const CMD_ERROR     As Long = -1
Public Const CMD_HALT      As Long = -2
```

구조 조회 실패도 현재 실행 계약상 `CMD_ERROR(-1)`와 같은 값을 사용한다.

단, 의미 이름은 분리한다.

```vba
Private Const LINE_NOT_FOUND As Long = CMD_ERROR
Private Const END_SELECT_NOT_FOUND As Long = CMD_ERROR
Private Const LABEL_NOT_FOUND As Long = CMD_ERROR
```

값은 같아도 의미는 다르다.

```text
-1 = 실행 실패 코드
-1 = 구조 탐색 실패
-1 = 라벨 미존재
```

1차에서는 값 동치를 유지한다.
장기적으로 trace/dry-run/diagnostics가 강화되면 실패 원인을 별도 타입으로 분리할 수 있다.

---

# 8. clsScriptProgram API 표면

## 8.1 Load

```text
LoadFromRange(range)
LoadFromArray(array, firstSheetRow = 1)
```

`firstSheetRow`는 시트 행 앵커를 보존하기 위한 값이다.

예:

```text
런타임 line 1 = 실제 Sheet row 2
```

## 8.2 Line Access

```text
TotalLines
RawCmdAt(line)
CmdAt(line)
P1At(line)
P2At(line)
P3At(line)
```

## 8.3 Line Classification

```text
IsBlankLine(line)
IsCommentLine(line)
IsSkippable(line)
```

## 8.4 Mapping

```text
SheetRowFromRuntimeLine(line)
RuntimeLineFromSheetRow(sheetRow)
```

## 8.5 Structure

```text
FindBlockPair(startLine, group, direction)
  group ∈ {IF, WHILE, REPEAT}
  SELECT는 거부하고 -1 반환

SelectBlock(selectLine)
  → clsSelectBlock

EndSelectOfCase(caseLine)
  → Long
```

## 8.6 Labels

```text
BuildLabelIndex
HasLabel(name)
LabelLine(name)
ResolveJumpTarget(target)
HasDuplicateLabels
DuplicateLabels
```

## 8.7 Validation

```text
ValidateStructure
ValidateLabels
ValidateLineRange
```

1차에서는 최소 검증만 구현하고, 상세 진단 객체는 2차로 미룰 수 있다.

---

# 9. clsSelectBlock 권장 구조

SELECT 결과 타입은 `clsSelectBlock` 클래스를 권장한다.

이유:

```text
Dictionary 값으로 넣기 쉽다.
CaseLines 배열 복사 문제가 줄어든다.
CaseCount / CaseLineAt 같은 접근자를 제공하기 쉽다.
깨진 SELECT 상태를 표현하기 쉽다.
```

권장 API:

```vba
' clsSelectBlock

Private m_selectLine As Long
Private m_endSelectLine As Long
Private m_caseLines As Collection

Public Property Get SelectLine() As Long
End Property

Public Property Let SelectLine(ByVal value As Long)
End Property

Public Property Get EndSelectLine() As Long
End Property

Public Property Let EndSelectLine(ByVal value As Long)
End Property

Public Property Get CaseCount() As Long
End Property

Public Function CaseLineAt(ByVal index As Long) As Long
End Function

Public Sub AddCaseLine(ByVal lineNo As Long)
End Sub
```

---

# 10. modEngine 방향

## 10.1 m_arrMain 제거

`modEngine`은 더 이상 배열을 직접 보유하지 않는다.

기존:

```vba
Private m_arrMain As Variant
```

변경:

```vba
Private m_program As clsScriptProgram
```

## 10.2 ActiveProgram 접근자

ICommand 시그니처를 유지하기 위해 1차에서는 접근자를 둔다.

```vba
Public Function ActiveProgram() As clsScriptProgram
    Set ActiveProgram = m_program
End Function
```

테스트 하니스에서도 동일 경로를 사용한다.

## 10.3 RunRPA 흐름

```text
1. ActiveSheet 스냅샷
2. CurrentRegion 확보
3. 헤더 제외
4. clsScriptProgram.LoadFromRange
5. Program.BuildStructureIndex
6. Program.BuildLabelIndex
7. 중복 라벨 또는 구조 오류가 있으면 안전 중단
8. 실행 루프 시작
9. ExecuteSmartAction
10. NextPointer
11. CleanUp
```

## 10.4 ExecuteSmartAction 방향

기존:

```vba
ExecuteSmartAction(arr, R)
```

1차 최소 변경:

```vba
ExecuteSmartAction(R)
```

또는 호환성을 위해 일시적으로 기존 시그니처를 유지하되 내부에서 `m_program`을 사용한다.

권장 최종 방향:

```vba
Public Function ExecuteSmartAction(ByVal R As Long) As Long
```

명령어와 파라미터는 Program에서 읽는다.

```vba
cmd = program.CmdAt(R)
rawP1 = program.P1At(R)
rawP2 = program.P2At(R)
rawP3 = program.P3At(R)
```

토큰 해석은 Engine이 수행한다.

---

# 11. SELECT Engine 헬퍼

`ResolveSelectTarget`은 구조 스캔을 하지 않는다.

변경 전:

```text
ResolveSelectTarget
  → m_arrMain 직접 전방 스캔
  → CASE 비교
  → fallback 처리
```

변경 후:

```text
ResolveSelectTarget
  → program.SelectBlock(selectLine) 조회
  → CASE 목록 순회
  → CASE 값 토큰 해석
  → first-match / last-fallback 처리
```

권장 형태:

```vba
Private Function ResolveSelectTarget(ByVal selectLine As Long, ByVal selVal As String) As Long
    Dim block As clsSelectBlock
    Set block = m_program.SelectBlock(selectLine)

    If block Is Nothing Then
        ResolveSelectTarget = CMD_ERROR
        Exit Function
    End If

    If block.EndSelectLine <= 0 Then
        ResolveSelectTarget = CMD_ERROR
        Exit Function
    End If

    Dim fallbackLine As Long
    fallbackLine = 0

    Dim i As Long
    For i = 1 To block.CaseCount
        Dim caseLine As Long
        caseLine = block.CaseLineAt(i)

        Dim caseVal As String
        caseVal = ResolveParam(m_program.P1At(caseLine))

        If UCase$(Trim$(caseVal)) = "ELSE" Or Len(caseVal) = 0 Then
            fallbackLine = caseLine
        ElseIf modActions.Action_CheckCondition(selVal, "EQ", caseVal) Then
            ResolveSelectTarget = caseLine + 1
            Exit Function
        End If
    Next i

    If fallbackLine > 0 Then
        ResolveSelectTarget = fallbackLine + 1
    Else
        ResolveSelectTarget = block.EndSelectLine
    End If
End Function
```

CASE 폴스루:

```vba
Case "CASE"
    Dim endSelLine As Long
    endSelLine = modEngine.ActiveProgram().EndSelectOfCase(R)
    If endSelLine > 0 Then
        ICommand_Execute = endSelLine
    Else
        ICommand_Execute = CMD_ERROR
    End If
```

---

# 12. 구현 단계

## P1. 열 인덱스 상수화

`clsScriptProgram` 내부에 열 상수를 둔다.

```vba
Private Const COL_ROW As Long = 1
Private Const COL_CMD As Long = 2
Private Const COL_P1  As Long = 3
Private Const COL_P2  As Long = 4
Private Const COL_P3  As Long = 5
```

외부 모듈은 열 번호를 알지 못한다.

---

## P2. clsScriptProgram 도입

구현 항목:

```text
LoadFromArray
LoadFromRange
TotalLines
RawCmdAt
CmdAt
P1At / P2At / P3At
IsBlankLine
IsCommentLine
IsSkippable
SheetRowFromRuntimeLine
RuntimeLineFromSheetRow
범위 검사
정규화 일원화
```

---

## P3. modEngine이 Program 사용

변경 항목:

```text
m_arrMain 제거
m_program 도입
ActiveProgram 추가
RunRPA에서 LoadFromRange 사용
상태 표시에서 program API 사용
ExecuteSmartAction에서 program API 사용
```

컴파일 갭 방지를 위해 이 Phase에서 배열 접근을 최대한 Program 경유로 전환한다.

---

## P4. FindBlockPair 이동

`modEngine.FindBlockPair`를 Program으로 이동한다.

구현 규칙:

```text
Registry 기반 탐색
group ∈ {IF, WHILE, REPEAT}만 허용
SELECT 전달 시 -1 반환
MIDBLOCK 즉시반환은 BlockGroup=IF 한정
startLine + direction부터 탐색
```

---

## P5. SELECT 단일 인덱스 구현

1차부터 반드시 인덱스 기반으로 구현한다.

구현 항목:

```text
clsSelectBlock 도입
Program.BuildSelectIndex
SelectLine → clsSelectBlock
CaseLine → EndSelectLine
SelectBlock(selectLine)
EndSelectOfCase(caseLine)
```

`ResolveSelectTarget`은 더 이상 전방 스캔하지 않는다.

---

## P6. LabelMap 이관

구현 항목:

```text
Program.BuildLabelIndex
HasLabel
LabelLine
ResolveJumpTarget
HasDuplicateLabels
DuplicateLabels
```

정책:

```text
LABEL 선언은 정적 심볼
LABEL P3에 토큰 문법 포함 시 검증 오류
중복 라벨은 안전 중단
End 사용 금지
```

`modContext.AddLabel` / `GetLabel`은 제거하거나 호환 래퍼로 축소한다.

---

## P7. FLOW 결합 정리

1차:

```text
clsFlowCommands는 modEngine.ActiveProgram() 경유
ICommand 시그니처 유지
```

장기:

```text
Flow Handler에 Program 주입
Engine 역질의 최소화
```

---

## P8. DevTools / ScriptAssist 수렴

1차에서는 런타임 안정화 우선.

2차에서 다음 복사본을 정리한다.

```text
modDevTools.PureCommandAt
modDevTools.IsBlockStarter
modDevTools.IsBlockEnder
modScriptAssist.NormalizeCmd
```

가능하면 Program 또는 공용 Normalizer로 수렴한다.

---

## P9. Instruction 모델

2차 이후 작업.

```text
clsInstruction
CmdDef 캐싱
Block 구조 캐싱
dry-run
trace
breakpoint
diagnostics
```

---

# 13. 검증 핵심

## 13.1 매핑

```text
A1 비시작 영역에서 SheetRow ↔ RuntimeLine 왕복 일치
Runtime Line / Sheet Row / COMMAND 동시 표기
범위 밖 RuntimeLine 접근 검출
```

## 13.2 정규화

```text
"> IF" → "IF"
">> TYPE" → "TYPE"
빈 줄은 IsBlankLine=True
"# comment"는 IsCommentLine=True
들여쓰기 주석의 기존 동작 보존
```

## 13.3 IF 블록

```text
IF → ELSE
IF → END_IF
ELSE → END_IF
중첩 IF depth 보존
```

## 13.4 WHILE 블록

```text
WHILE → END_WHILE
END_WHILE → WHILE
중첩 WHILE depth 보존
WHILE 조건 참인데 END_WHILE이 없으면 LoopPush 금지
```

## 13.5 REPEAT 블록

```text
REPEAT → END_REPEAT
END_REPEAT → REPEAT
중첩 REPEAT depth 보존
END_REPEAT 누락 시 CMD_ERROR
```

## 13.6 FindBlockPair 거부

```text
FindBlockPair(..., "SELECT", 1) = -1
FindBlockPair(..., "SELECT", -1) = -1
```

## 13.7 MIDBLOCK 안전성

```text
IF의 ELSE는 depth=0에서 즉시반환
SELECT의 CASE는 FindBlockPair 대상 아님
WHILE/REPEAT에 가상 MIDBLOCK 등록해도 즉시반환 안 됨
```

## 13.8 SELECT

```text
SelectBlock = CASE 목록 + END_SELECT
EndSelectOfCase = CASE → END_SELECT 역인덱스
중첩 SELECT 내부 CASE 미수집
CASE 본문 안 IF/WHILE/REPEAT는 SELECT depth에 영향 없음
일치 CASE 점프 = first-match
CASE ELSE 폴백 = last-fallback
빈 CASE 폴백 = last-fallback
무일치·무폴백 → END_SELECT 줄
END_SELECT 누락 → -1 → CMD_ERROR
CASE 폴스루는 FindBlockPair 미사용
```

## 13.9 LABEL

```text
LABEL 인덱스 생성
중복 LABEL → HasDuplicateLabels=True
중복 LABEL 발생 시 End 사용 금지
LABEL P3 빈값 검출
LABEL P3 토큰 문법 검출
GOTO 정적 라벨 조회
CALL 정적 라벨 조회
GOTO/CALL 동적 대상 해석 후 라벨 조회
미존재 라벨 → ResolveJumpTarget=-1
```

## 13.10 점프 범위

```text
GOTO 0 → CMD_ERROR
GOTO 음수 → CMD_ERROR
GOTO TotalLines 초과 → CMD_ERROR
GOTO 숫자 정상 범위 → 해당 줄
```

---

# 14. 최종 방향

```text
Excel Sheet
  = 소스

ScriptProgram
  = 로드된 프로그램 이미지
  = 구조 위치
  = 줄/행 매핑
  = 라벨 인덱스
  = SELECT 구조 인덱스

Registry
  = 언어 정의
  = 구조 의미의 진실원천

Engine
  = 인터프리터
  = 런타임 값 해석
  = SELECT 값 비교와 fallback
  = 다음 IP 결정

Context
  = 실행 상태
  = Memory / CallStack / LoopStack

Actions / WinAPI
  = 현실 세계 효과
```

이번 설계의 본질은 배열을 클래스로 감싸는 것이 아니다.

UnivRPA의 Excel 시트는 이미 작은 DSL이고, `m_arrMain`은 이미 그 DSL의 로드된 프로그램 이미지다.

따라서 `clsScriptProgram`은 장식 계층이 아니라 런타임 모델이다.

핵심 원칙은 다음 네 가지다.

```text
1. Program은 구조 위치를 안다.
2. Registry는 구조 의미를 안다.
3. Engine은 런타임 값을 안다.
4. Context는 실행 상태를 안다.
```

특히 SELECT는 `FindBlockPair`의 하위 기능이 아니다.

SELECT는 작은 디스패치 테이블이다.

따라서 1차 구현부터 다음 원칙을 지킨다.

```text
SELECT 구조는 단일 인덱스
SELECT 실행은 Engine
CASE 폴스루는 EndSelectOfCase
FindBlockPair는 SELECT 거부
```

이로써 UnivRPA는 “엑셀 매크로 묶음”이 아니라, Excel을 호스트로 삼는 작은 DSL 런타임으로 정렬된다.
