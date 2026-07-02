# spec: `ELSE_IF` 명령어 도입

## 목표
`IF` 블록에 `ELSE_IF` 분기를 추가하여 N-way 분기를 한 레벨 들여쓰기로 평탄화한다. SELECT/CASE에 패턴 매칭을 도입하지 않고도 패턴 분기 케이스를 흡수한다.

## 범위
- `modCmdRegistry` — `ELSE_IF` 명령어 등록 (1개 추가)
- `clsCommandsFlow` — `IF` / `ELSE` / `END_IF` 핸들러 수정, `ELSE_IF` 핸들러 신설
- `clsScriptProgramIndex` (또는 동등 모듈) — `EndIfOf(R) As Long` 인덱스 신설
- `clsScriptProgramAnalysis` — `IF` 블록 구조 정적 검증 규칙 추가
- `modContext` — `IfBranchStack` 신설 + 점프 reconcile 헬퍼
- `02-command-reference.md` — `ELSE_IF` 추가, FLOW 명령 개수 18→19, IF 블록 짝/문법 보강
- `README.md` — FLOW 명령 목록·개수 갱신

## Negative Scope
- **SELECT/CASE 매칭 의미론은 변경하지 않는다.** 패턴 매칭/모드 플래그/와일드카드 자동 승격 모두 도입 안 함.
- **조건 연산자 목록은 변경하지 않는다.** `LIKE`/`STARTS`/`ENDS`/`CONTAINS` 등 기존 그대로.
- **`ELSE` 의미론은 변경하지 않는다.** 토큰 처리, fallthrough 시 END_IF로 점프하는 현재 동작 유지.
- **사용자 변수/토큰 시스템은 건드리지 않는다.**
- **시트 포맷터 색상 그룹은 기존 FLOW 그대로** — 새 색상 그룹 추가 없음.

## 인터페이스 / 제약

### 명령어 사양

| 명령어 | P1 (C) | P2 (D) | P3 (E) | 설명 |
|--------|--------|--------|--------|------|
| `ELSE_IF` | 좌변 (str, 필수, RESOLVE) | 연산자 (str, RAW, 필수) | 우변 (str, 필수, RESOLVE) | `IF`/`ELSE_IF` 거짓 시 평가되는 추가 분기. 조건 참이면 본문 실행, 거짓이면 다음 `ELSE_IF`/`ELSE`/`END_IF`로 |

- 파라미터 사양은 `IF`와 동일 (P2만 RAW, P1/P3 RESOLVE).
- 카테고리: FLOW. BlockGroup: `IF`. BlockRole: `MIDBLOCK`.

### 블록 문법

```
IF | ... | ... | ...
  body1
ELSE_IF | ... | ... | ...
  body2
ELSE_IF | ... | ... | ...
  body3
ELSE
  bodyN
END_IF
```

- `ELSE_IF`는 0개 이상, 임의 개수 허용.
- `ELSE`는 0~1개. 있으면 마지막 `ELSE_IF` 뒤에 위치해야 함.
- **`ELSE` 뒤에 `ELSE_IF` 금지** (정적 검증에서 오류).
- 중첩 가능. 현재 `IF` 블록과 동일한 짝 매칭 규칙.

### 실행 의미론 (runtime)

분기 채택을 추적하기 위해 `modContext`에 `IfBranchStack`을 추가한다. 각 프레임은 채택된 분기의 `END_IF` 줄 번호를 담는다.

```
IF / ELSE_IF at R:
    if cmd == "ELSE_IF" and IfBranchPeek() == EndIfOf(R):
        # fallthrough — 이전 분기가 채택되어 body가 실행을 마치고 도달
        go to EndIfOf(R)
        return

    if condition true:
        IfBranchPush(EndIfOf(R))
        next line

    else:  # condition false
        target = FindBlockPair(R, "IF", 1)
        case target.cmd:
            "ELSE_IF": go to target          # 평가 위임
            "ELSE":    go to target + 1      # else body 실행 (push 불필요)
            "END_IF":  go to target

ELSE at R:
    # 현재와 동일 — fallthrough로만 도달
    go to FindBlockPair(R, "IF", 1)   # = EndIfOf(R)

END_IF at R:
    if IfBranchPeek() == R:
        IfBranchPop()
    next line
```

**핵심 불변식**: `IF`/`ELSE_IF`가 참으로 평가되어 body로 진입한 경우에만 `IfBranchPush`. `END_IF`가 자기 줄과 일치하는 top을 만나면 `IfBranchPop`. `ELSE` 진입은 push 안 함 (END_IF에서 pop할 게 없으므로 자연스럽게 균형).

### 점프 reconcile

`BREAK` / `CONTINUE` / `GOTO` / `RETURN`이 IF 블록 경계를 가로질러 점프할 때 `IfBranchStack`에 stale 프레임이 남지 않도록 reconcile한다.

- 규칙: 점프 대상 줄(target)이 프레임의 `[headerLine, endIfLine]` 범위 밖이면 해당 프레임 pop.
- 기존 `modContext.ReconcileLoopStackForJump`와 동일 패턴으로 `ReconcileIfBranchStackForJump(targetLine)` 신설.
- `modEngine.ResolveGotoTarget` 등에서 LoopStack reconcile 직후 호출.

### 인덱스: `EndIfOf(R)`

- 입력: `IF` 블록 내 임의의 줄 (IF, ELSE_IF, ELSE, END_IF 중 하나)
- 출력: 해당 블록의 `END_IF` 줄 번호
- 입력이 IF 블록 소속이 아니면 0 반환
- 구현: 파싱 단계에서 `clsScriptProgramIndex` 빌드 시 `IF` 블록을 스캔하여 멤버 줄 → END_IF 매핑 테이블 구축. SELECT 블록 인덱스(`clsSelectBlock`)와 동일한 접근.

### 정적 검증

`clsScriptProgramAnalysis`에 다음 규칙 추가:

1. `ELSE_IF`는 반드시 `IF` 블록 내부에 있어야 한다 (`IF` 없이 단독 등장 시 오류).
2. `IF` 블록 내부에서 `ELSE` 다음에 `ELSE_IF`가 오면 오류.
3. `ELSE_IF`의 P1/P2/P3 누락 시 오류 (`IF`와 동일).
4. `ELSE_IF`의 P2는 `IsConditionOperator` 통과해야 한다 (`IF`와 동일).

### 시트 표시

- 들여쓰기 깊이: `IF` 헤더와 같은 레벨 (= `ELSE`와 동일).
- 색상 그룹: 기존 FLOW (`clrFLOW`) 그대로.
- 페인트 범위: `IF`/`ELSE`와 동일 (`FULLROW` 또는 현 규약 따름).
- 포맷터(`Ctrl+Shift+F`)는 `ELSE_IF`를 `IF` 사슬의 일부로 인식해 동일하게 정렬.

## 수락 기준

- [ ] `ELSE_IF` 명령어가 레지스트리에 등록되고 시트 검증·힌트에 노출된다.
- [ ] `IF → ELSE_IF → END_IF` (ELSE 없음): 모든 분기 false면 어떤 body도 실행 안 됨, 첫 true 분기 body만 실행됨.
- [ ] `IF → ELSE_IF → ELSE → END_IF`: 모든 IF/ELSE_IF false → ELSE body 실행. 어느 하나 true → 해당 body만 실행, ELSE 안 함.
- [ ] `IF → ELSE_IF → ELSE_IF → ELSE → END_IF`: 첫 true 분기 body만 실행됨 (이후 ELSE_IF 평가 안 됨, ELSE 안 함).
- [ ] 중첩: 외부 IF body 안의 내부 IF/ELSE_IF/END_IF가 외부 분기 추적을 깨지 않는다.
- [ ] `ELSE_IF`의 P2 연산자가 `LIKE`/`STARTS`/`ENDS`/`CONTAINS` 등 모든 기존 조건 연산자를 지원한다.
- [ ] `ELSE` 뒤 `ELSE_IF`는 정적 검증에서 오류로 표시된다.
- [ ] `ELSE_IF` 단독 (대응 `IF` 없음)은 정적 검증에서 오류로 표시된다.
- [ ] `IF` body 안에서 `BREAK` / `GOTO` / `RETURN`으로 블록 밖으로 점프해도 `IfBranchStack`이 누수되지 않는다.
- [ ] 기존 `IF → ELSE → END_IF` 스크립트가 그대로 동작한다 (회귀 없음).
- [ ] 기존 `IF → END_IF` (ELSE 없음) 스크립트가 그대로 동작한다 (회귀 없음).
- [ ] `02-command-reference.md`에 `ELSE_IF` 항목·예제 추가, FLOW 개수 19로 갱신.
- [ ] `README.md` FLOW 명령 목록과 개수 갱신.

## 구현 순서

1. **레지스트리 + 표시**
   - `modCmdRegistry`에 `ELSE_IF` 등록 (IF와 동일한 파라미터 사양, MIDBLOCK, BlockGroup="IF").
   - 시트 포맷터·힌트·들여쓰기에서 `ELSE_IF`가 IF 사슬로 인식되는지 확인.
   - 이 단계만으로 시트 입력은 가능 (실행은 아직 안 됨).

2. **인덱스: `EndIfOf`**
   - `clsScriptProgramIndex` (또는 동등 모듈)에 IF 블록 스캔 추가.
   - `clsScriptProgram.EndIfOf(R)` public 메서드 노출.
   - 단위 테스트: 단일 IF, 중첩 IF, ELSE_IF 다수 케이스.

3. **`modContext.IfBranchStack` + reconcile 헬퍼**
   - `IfBranchPush`/`IfBranchPop`/`IfBranchPeek` API 신설.
   - `ReconcileIfBranchStackForJump(targetLine)` 신설.
   - `InitContext`에서 초기화.

4. **`clsCommandsFlow` 실행 로직**
   - `IF` 핸들러 수정: false 분기에서 target이 `ELSE_IF`면 target으로(평가), `ELSE`면 target+1로, `END_IF`면 target으로. true 시 `IfBranchPush`.
   - `ELSE_IF` 핸들러 신설: fallthrough 감지 → `EndIfOf`로 점프 / jump 진입 → IF와 동일 로직.
   - `END_IF` 핸들러 수정: peek 일치 시 pop.
   - `ELSE` 핸들러: 기존 그대로 유지 (push 안 함, 단순 `FindBlockPair`).

5. **점프 명령들 reconcile 호출**
   - `modEngine.ResolveGotoTarget`에 `ReconcileIfBranchStackForJump` 호출 추가.
   - `clsCommandsFlow`의 `BREAK` / `CONTINUE` / `RETURN` 분기에서 IfBranchStack reconcile.

6. **정적 검증**
   - `clsScriptProgramAnalysis`에 IF 블록 구조 규칙 추가 (ELSE 뒤 ELSE_IF 금지, ELSE_IF 고아 금지, 파라미터/연산자 검증).

7. **문서**
   - `02-command-reference.md`: `ELSE_IF` 행 추가, IF 블록 문법 절 보강, 예제 추가, FLOW 개수 19로.
   - `README.md`: 카테고리 표의 FLOW 개수와 명령어 나열 갱신.

8. **통합 테스트**
   - 수락 기준의 모든 시나리오 시트 작성 후 실행 검증.
   - 회귀: 기존 IF/ELSE/END_IF 스크립트 재실행.
