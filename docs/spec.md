# spec.md — UnivRPA 단위테스터 확장 (동적 흐름·메모리 검증)

> 대상: 코드 에이전트(Claude Code). 본 문서만으로 구현 가능하도록 자립형으로 작성됨.
> 코드베이스: `UnivRPA.xlsb` (VBA, 28 모듈). 테스트 모듈 `modTester` 확장이 핵심.

---

## 목표

`clsScriptProgram` 1급 객체 위에 **헤드리스 완주 드라이버**를 얹어, 현재 테스터가 검증 못 하는 **동적 흐름제어·메모리 상호작용·네거티브 경로**를 단위검증한다. 기존 단일라인 하니스는 일반 드라이버의 1스텝 케이스로 흡수한다.

---

## 범위

수정/추가 대상은 **`modTester` 단일 모듈에 한정**한다(필요 시 보조 표준 모듈 1개 신설 가능, 단 테스트 전용).

- `modTester` — 신규 드라이버·단언 동사·시나리오 항목·진입점 추가, 로직 항목 이관
- (선택) `modTestKit` — 드라이버/동사를 분리하고 싶을 경우의 신설 표준 모듈명. 분리 불필요하면 `modTester` 내부에 둔다.

---

## Negative Scope (건드리지 않을 것 — 확정)

다음 **프로덕션 모듈은 일절 수정 금지**. 트레이스는 드라이버 루프가 소유하며, 엔진에 훅을 심지 않는다.

- `modEngine`, `modContext`, `clsScriptProgram`, `clsInstruction`, `clsSelectBlock`
- `modCmdRegistry`, `CmdDef`, `ParamSpec`, `ICommand`
- `clsFlowCommands`, `clsMemoryCommands`, `clsActionCommands`, `clsVisionCommands`, `clsSystemCommands`
- `modActions`, `modWinAPI`, `modHelper`, `modDevTools`, `modScriptAssist`, UI 폼류

`Run_System_SelfDiagnosis`의 **로직 항목 이관**은 삭제가 아니라 헤드리스 스위트로 **이동**이다. 순 커버리지는 보존 또는 증가해야 한다(절대 감소 금지).

---

## 인터페이스 / 제약

### 실행 경로 결합 (필독 — 설계 전제)

FLOW 핸들러는 실행 중 `modEngine.ActiveProgram()`을 도로 호출한다(`FindBlockPair`, `EndSelectOfCase`, `ResolveSelectTarget`). 따라서 **실행은 반드시 엔진의 `m_program`을 경유**해야 한다. 테스터가 자체 보유한 `clsScriptProgram`으로 직접 실행하면 엔진이 그 객체를 보지 못한다.

- **실행 항목** → `modEngine.LoadProgramFromArray` → `modEngine.ExecuteSmartAction` 경유 (절대 직접 실행 조립 금지)
- **정적 구조 항목** → 테스터가 `New clsScriptProgram`으로 직접 조회 (엔진 불필요)

### 의존 API (프로덕션, 호출만)

엔진:
- `modEngine.LoadProgramFromArray(arr, firstRow)` — 적재 + 구조/라벨 인덱스 빌드
- `modEngine.ExecuteSmartAction(R) As Long` — 한 줄 실행, 반환코드 리턴
- `modEngine.ActiveProgram() As clsScriptProgram`

반환코드 상수 (`modCmdRegistry`, Public):
- `CMD_NEXT_LINE = 0`, `CMD_ERROR = -1`, `CMD_HALT = -2`

컨텍스트 상태:
- `modContext.InitContext` / `ResetContext`
- `modContext.GetVar(key) As String` / `SetVar(key, val)`
- `modContext.GetStackDepth() As Long` (콜스택)
- `modContext.LoopDepth() As Long`, `GetActiveLoopIteration() As Long` (루프스택)

정적 조회 (`clsScriptProgram`, 직접 New):
- `LoadFromArray(arr, firstRow)` → `BuildInstructionCache`만 수행. **SELECT/라벨 조회 전 `ValidateStructure()`/`ValidateLabels()`(또는 `BuildStructureIndex`/`BuildLabelIndex`) 선행 필수.**
- `totalLines`, `CmdAt/P1At/P2At/P3At`, `IsBlankLine/IsCommentLine/IsSkippable`
- `FindBlockPair(start, group, dir)`(인덱스 불요), `SelectBlock(n)`, `EndSelectOfCase(n)`
- `HasLabel/LabelLine/ResolveJumpTarget/HasDuplicateLabels/DuplicateLabels`
- `InstructionAt(n) As clsInstruction` → `.cmd/.p1/.p2/.p3/.def/.RuntimeLine/.sheetRow/.rawCmd`

### 기존 modTester 자산 (재사용)

- `MakeCmd(arr, R, cmd, p1, p2, p3)` — 배열 1행 채움(2~5열). **입력 DSL 그대로 유지.**
- `ExpectLong/ExpectText/ExpectBool(name, exp, act, ByRef failCount)`
- `AssertCond`, `LogPass/LogFail`
- `Test_RegistryDispatchSync()`(라우팅 전수), `Test_ScriptProgramModel()`(정적 구조) — 유지·확장

### 신규 드라이버 계약

```vba
Public Enum eRunResult
    rrCompleted = 0   ' PC가 범위 밖으로 정상 이탈
    rrHalted    = 1   ' EXIT → CMD_HALT
    rrErrored   = 2   ' CMD_ERROR 또는 미지의 반환코드
    rrExhausted = 3   ' 스텝예산 초과 (폭주 루프 포획)
End Enum

Public Function RunProgram(ByRef arr As Variant, ByVal firstRow As Long, ByVal maxSteps As Long, _
                           ByRef outTrace() As Long, ByRef outLastCode As Long, _
                           ByRef outLastLine As Long) As eRunResult
```

동작(= `RunRPA` 구동 루프에서 UI/ESC만 제거):
1. `modContext.InitContext` (시나리오 격리: 메모리·콜스택·루프스택 리셋)
2. `modEngine.LoadProgramFromArray arr, firstRow` **1회만**
3. `curLine = 1`부터 루프. 매 스텝 **진입 직전 `curLine`을 `outTrace`에 기록** → `code = ExecuteSmartAction(curLine)` → `outLastCode = code` → 분기:
   - `CMD_NEXT_LINE` → `curLine = curLine + 1`
   - `code > 0` → `curLine = code` (점프)
   - `CMD_HALT` → `outLastLine = curLine`; **rrHalted** 반환
   - `CMD_ERROR` 또는 그 외 음수 → `outLastLine = curLine`; **rrErrored** 반환
4. 루프 가드: `curLine < 1 Or curLine > totalLines` → **rrCompleted** 반환
5. 스텝 카운트가 `maxSteps` 도달(터미널 미발생) → `outLastLine = curLine`; **rrExhausted** 반환

#### 마이크로 기본값 (확정됨)
1. **트레이스 입도**: PC가 닿는 *모든* 줄 기록(빈줄/주석 포함, 실제 경로 그대로).
2. **종료 4분류 유지**: `rrCompleted/rrHalted/rrErrored/rrExhausted` 전부 구분.
3. **스텝예산 기본 10,000**, 호출별 `maxSteps` override.
4. **반환형**: 종료사유는 함수값(`eRunResult`), `outTrace/outLastCode/outLastLine`은 ByRef out. 단언 동사는 out-param + `modContext`를 읽어 판정.

#### 단일라인 sugar
```vba
Public Function Step1(ByRef arr As Variant, ByVal R As Long) As Long  ' = RunProgram(arr, R 기준, maxSteps:=1) 후 outLastCode 반환
```
- 1스텝 실행. `ERROR`/`HALT`면 그 코드가 `outLastCode`에 잡혀 즉시 종료. 점프(예: 50)면 `outLastCode=50` 잡고 `rrExhausted`로 종료 → 테스트는 `outLastCode`만 단언.
- 별도 메커니즘 아님(래퍼). 기존 `ExecuteArrayLine`의 reload-per-call은 폐기하고 이 sugar로 대체.

### 신규 단언 동사 (전부 기존 `Expect*` 패턴 복제, `ByRef failCount`)

| 동사 | 관찰축 | 판정 |
|---|---|---|
| `ExpectVar(name, exp, failCount)` | M | `modContext.GetVar(name) = exp` |
| `ExpectStackDepth(exp, failCount)` | S | `modContext.GetStackDepth() = exp` |
| `ExpectLoopDepth(exp, failCount)` | S | `modContext.LoopDepth() = exp` |
| `ExpectTerminal(exp As eRunResult, actual, failCount)` | R | 종료사유 일치 |
| `ExpectVisits(trace, line, n, failCount)` | T | trace 내 `line` 등장 횟수 = n |
| `ExpectTrace(trace, expArr, failCount)` | T | 소규모 경로 정확순서 일치 |

`ExpectVisits`가 동적 흐름 주력: "본문 줄 정확히 N회", "ELSE 본문 0회", "미매칭 CASE 0회"가 전부 이걸로 환원.

### 시나리오 어휘 제약

헤드리스 시나리오는 **FLOW + MEMORY + EXIT + 주석(#)만** 사용. ACTION/VISION/`FOCUS`/`SLEEP`/`NOTIFY` 등 실제 I/O 명령 **금지**(결정론·무인 실행 보장). 부수효과 관찰은 `SET`/`INC`로 메모리에 흔적을 남겨 단언.

### 진입점 분리

- `Run_System_SelfDiagnosis` (기존) → **I/O 반수동** 전용. 13개 I/O 명령 + 시트/창. 로직 항목은 헤드리스로 이관해 덜어냄.
- `Run_Flow_Diagnosis` (신규) → **헤드리스**. FLOW/MEMORY/EXIT/resolver/static. 시트·창 0, `InitContext`만.
- 공유 리포트 코어(`failCount` 누적, `LogPass/LogFail`, 최종 요약). 선택적 `Run_All`이 둘을 호출.

---

## 검증 항목 (승인된 매트릭스)

C=현재커버 / G=신규(공백) / X=로직범위밖. 관찰축 R(반환코드) M(메모리) S(스택) T(경로) D(decode).

### FLOW — 반환코드·상태 계약

각 명령의 계약은 코드에서 도출된 사실이며, 시나리오는 이를 단언해야 한다(시나리오 프로그램 구체는 에이전트 재량).

| 구문 | 계약 (assert 대상) | 축 | 상태 |
|---|---|---|---|
| LABEL | no-op → NEXT_LINE | R | G |
| GOTO | 라벨/숫자 점프 ; 미해석 → ERROR | R | C(정적)/G(실행·에러) |
| CALL | 타깃 점프 + `PushStack(R+1)` ; 미해석 → ERROR | R,S | G(본문실행·에러) |
| RETURN | `PopStack` 복귀 ; 빈스택 → ERROR | R,S | G(에러) |
| IF 참 | NEXT_LINE, **본문 실행 후 ELSE 도달 시 END_IF로 점프(ELSE본문 스킵)** | R,T | G |
| IF 거짓 | ELSE면 ELSE+1 / END_IF면 그 줄 ; 짝없음 → ERROR | R,T | G |
| WHILE | 참=push 또는 회차+1; **제로트립**(즉시거짓 → END_WHILE+1, push 없음); 거짓 시 활성프레임이면 POP | R,S,T | G |
| END_WHILE | WHILE 헤더 백점프 | R,T | G(실행) |
| REPEAT | maxCount(=P1; 0/빈/≤0=무한) push, NEXT_LINE ; 짝없음 → ERROR | R,S | G |
| END_REPEAT | 무한 또는 iter<max → 회차+1, head+1 재실행 ; else POP, NEXT_LINE ; 프레임불일치 → ERROR | R,S,T | G |
| BREAK | `LoopTopEnd+1` 탈출 + 프레임 POP ; 루프밖 → ERROR | R,S | G |
| CONTINUE | `LoopTopEnd`(ENDER)로 점프 → 재평가 ; 루프밖 → ERROR | R,T | G |
| SELECT | 매칭 CASE+1 / 폴백(ELSE·빈값)+1 / END_SELECT ; no-block → ERROR | R,T | C(단발)/G(다분기·폴백·에러) |
| CASE | 실행 도달=폴스루 → END_SELECT 점프 ; 미발견 → ERROR | R,T | C(단발)/G(연쇄) |
| END_IF/END_SELECT | NEXT_LINE | R | C(간접) |
| **중첩** | loop-in-loop / if-in-loop / call-in-loop. 최내곽 `LoopDepth` 피크, 정리 후 0, 외부 회차 보존 | S,T | G (최고가치) |

**REPEAT 회차 스펙 고정**: `REPEAT N`에서 본문은 **정확히 N회**(REPEAT 줄 1회 진입 + END_REPEAT가 iter<max 동안 재실행). max=2 → 본문 2회. 이 값을 기대값으로 못박는다.

### MEMORY

| 항목 | 축 | 상태 |
|---|---|---|
| SET 쓰기 / 덮어쓰기 / `{$BLANK}` 클리어 | M | C |
| SET 읽기전용 좌변 거부(`{$LOOP}`·`{$PIXEL}`) → ERROR | R | G (INC만 C) |
| SET 형식오류 좌변 거부 → ERROR | R | G |
| INC/DEC 정수 RMW(미정의→0) | M | C |
| INC/DEC 증분 P3(변수·음수) | M | G(변수·음수) |
| **루프 내 누적**(예: 1..N 합) — 흐름×메모리 | M,T | G |

### EXIT / 순수 서브로직

| 항목 | 축 | 상태 |
|---|---|---|
| EXIT → HALT (`rrHalted`) | R | G |
| WAIT_VANISH 즉시통과 분기 | R | C |
| 조건 연산자 EQ/LIKE/STARTS/ENDS/CONTAINS/금지`=` (`Action_CheckCondition` 직접) | 직접 | C |
| I/O 13종 실제 동작 (MOVE/CLICK/.../ASSERT/FOCUS/SLEEP/NOTIFY) | — | **X** (반수동) |

### RESOLVER / LVALUE (decode 축)

| 항목 | 축 | 상태 |
|---|---|---|
| 보간 `{a}_world` / 리터럴 `{$TAB}` | D | C |
| `{$CR}` / `{$LF}` / `{$SPACE}` | D | G |
| 중첩 `{$PIXEL:{x},{y}}` | D | G |
| 짝없는 `{` graceful fallback | D | G |
| 미지의 `{$XXX}` 리터럴 통과 | D | G |
| `{$LOOP}` 루프밖=0 / **루프안 라이브 회차** | D,T | C(밖)/G(안) |

> 정적 구조(FindBlockPair/SelectBlock/라벨/ResolveJumpTarget/ResolveSelectTarget)는 `Test_ScriptProgramModel`가 이미 C. 신규 항목은 그 위에 얹는다.

---

## 수락 기준

- [ ] `RunProgram`이 3줄(SET/SET/EXIT) 프로그램에서 `rrHalted` 반환, 두 변수값 정확.
- [ ] WHILE 카운팅: 본문 줄 방문 = 의도한 N회, 종료 후 `LoopDepth=0`, `{$LOOP}` 캡처값이 1..N 오름차순.
- [ ] WHILE 제로트립: 본문 방문 0회, `LoopDepth` 시종 0.
- [ ] `REPEAT 3`: 본문 방문 정확히 3회.
- [ ] 중첩 루프: 최내곽 시점 `LoopDepth=2`, 완주 후 0, 외부 회차 보존.
- [ ] BREAK: `LoopTopEnd+1`로 탈출, `LoopDepth`가 바깥 레벨로 복귀.
- [ ] CONTINUE: 본문 잔여 스킵 후 다음 회차 진입(방문 패턴으로 확인).
- [ ] IF 참: 참 본문 방문 ≥1, **ELSE 본문 방문 0**, 정상 종료.
- [ ] IF 거짓: ELSE 본문 방문 ≥1, **IF 본문 방문 0**.
- [ ] SELECT: 매칭 CASE 본문만 방문, 다른 CASE 본문 0; 폴백 케이스 동작 확인.
- [ ] CALL/RETURN: 서브 본문 방문, CALL 다음 줄로 복귀, 균형 후 `GetStackDepth=0`; 중첩 CALL LIFO 복귀.
- [ ] 네거티브: GOTO 미해석 → `rrErrored`; RETURN 빈스택 → `rrErrored`; BREAK/CONTINUE 루프밖 → `rrErrored`; SELECT no-block → `rrErrored`.
- [ ] 폭주 가드: 항상-참 WHILE → 예산 내 `rrExhausted`(엑셀 무응답 없이 종료).
- [ ] SET `{$LOOP}` 좌변 → `rrErrored`.
- [ ] 메모리: 루프 내 누적 합 정확(흐름×메모리).
- [ ] resolver: `{$CR}/{$LF}/{$SPACE}` 치환, 중첩 `{$PIXEL}` 해석, 짝없는 `{` 리터럴 보존, 미지 `{$XXX}` 원형 통과.
- [ ] `Run_Flow_Diagnosis` 실행 시 **시트 생성·창 활성화·실제 입력 0건**.
- [ ] 기존 `Test_ScriptProgramModel` / `Test_RegistryDispatchSync` 여전히 통과.
- [ ] **diff가 `modTester`(및 선택적 `modTestKit`)에 한정** — Negative Scope 모듈 무변경.

---

## 구현 순서

1. **스켈레톤**: `eRunResult` enum, `RunProgram` 드라이버 + `Step1` sugar, 신규 `Expect*` 동사. 3줄(SET/SET/EXIT) 스모크로 드라이버·트레이스·종료분류 검증.
2. **단일 구문 해피패스**: IF/ELSE, WHILE, REPEAT, SELECT, GOTO, CALL/RETURN 각 1개 시나리오(방문수·종료사유·스택/루프 깊이).
3. **중첩·상호작용**: loop-in-loop, if-in-loop, call-in-loop, BREAK/CONTINUE, 중첩 CALL, 루프 누적·카운터 매트릭스.
4. **네거티브 경로**: 모든 ERROR 계약 + `rrExhausted` 폭주 가드.
5. **메모리·resolver 헤드리스 항목**: SET 좌변 거부, INC/DEC 변수·음수, `{$CR/$LF/$SPACE}`, 중첩 pixel, 짝없는 brace, 미지 `$`, `{$LOOP}` 라이브.
6. **배선**: `Run_Flow_Diagnosis`가 2~5 집계, `Run_System_SelfDiagnosis`에서 로직 항목 이관(커버리지 보존), 공유 리포트 코어, 선택적 `Run_All`.

---

## 미확정 항목

없음. (드라이버 계약·마이크로 기본값 4종·검증 매트릭스·진입점 분리 전부 확정됨.)
