# UnivRPA 본체 설계 리뷰 및 개선 태스크 (2026-07-08)

전 소스 정독 기반 설계 리뷰 결과. vbaflow 적용 직전 시점에 작성되었으며,
우선순위는 "vbaflow 온보딩에 필수인 것"부터 배열했다.

## 총평

선언된 아키텍처는 실제로 지켜지고 있다.

- 파사드 규율(clsScriptProgram): 외부 누출 0건. 외부 모듈은 전부 파사드만 호출
- modWinAPI(L4): 무상태 원칙 준수, 상위 역호출 없음
- 레지스트리: GetCmdDef 단일 게이트 관철, m_reg 직접 접근 없음
- 비전 계열: "타임아웃 시 모달 금지, 조용히 실패" 무인 원칙 일관

약점은 레이어 경계가 아니라 레이어를 가로지르는 신호와 지식에 있다:

1. 오류/취소 신호가 상위로 전파되지 않음
2. 레지스트리가 선언한 파라미터 지식을 런타임이 강제하지 않음
3. 명령어 시맨틱이 modDevTools에 두 번째 사본으로 존재
4. 런타임이 팝업 UI에 강결합

---

## A. vbaflow 온보딩 필수: 헤드리스 실행 계약

### A1. Action_Notify의 MsgBox 분기 제거 (유일한 hang 지점)

- 위치: `modActions.bas:493`
- 문제: NOTIFY 지속시간이 0/공백이면 MsgBox로 블로킹.
  vbaflow test에서는 타임아웃 -> Excel kill -> 세션 poison으로 끝난다.
- 조치: 헤드리스 판별(또는 컨텍스트 플래그) 시 타임드 알림 + jsonl 로그로 강등.
- 비고: 런타임 경로에서 헤드리스를 실제로 깨는 곳은 이 한 곳뿐.

### A2. 알림 싱크 단일화

- 문제: 엔진이 12곳 이상에서 `Action_ShowNotification`(초 단위)을 직접 호출하고,
  별도로 `Action_Notify`(ms 단위, 블로킹 가능)가 중첩 존재.
- 조치: 알림 게이트 하나로 통합. "대화형=팝업 / 헤드리스=LogShim만"을
  게이트 안에서 분기하면 엔진 코드 수정 없이 헤드리스 전환 완료.

### A3. 테스트 표면을 vbaflow 계약에 맞추기

- 문제: vbaflow test는 표준 모듈의 인자 없는 Public `Test_*` Sub만 열거하는데,
  modTester의 실질 테스트 12개는 전부 Private Function(실패건수 반환)이라
  하나도 잡히지 않는다. 유일한 Public은 `Test_RegistryDispatchSync`(Long 반환).
- 조치:
  - 각 테스트에 얇은 Public Sub 래퍼 추가
    (내부 Function 호출 후 `LogShim.AssertEq "fail_count", 0, n`)
  - 단언을 LogShim.AssertEq 계열로 표준화하여 vbaflow 판독 계약
    ("에러 없이 반환 AND ok:false 0건")에 편입
  - `Run_Flow_Diagnosis` 래퍼(`Test_FlowDiagnosis`)도 추가 검토

### A4. "헤드리스" 스위트의 오염 제거

- 문제: `Run_Flow_Diagnosis`는 헤드리스를 표방하나 활성 시트 상태에 의존:
  - `Test_ForeachRowExecution`: `modTester.bas:880` Worksheets.Add로 실제 시트 생성/삭제
  - `Test_ScriptAssistFormatterRegression`: `modTester.bas:1444,1466,1478`
    ActiveSheet / .Activate / .Select 의존
- 조치: 시트는 참조로 다루고 Activate/Select 제거 (vbaflow 비가시 세션 안정화).
- 부수: `Run_All`이 `Test_RegistryDispatchSync`, `Test_ScriptProgramModel`을
  두 번 실행하는 중복 정리 (System 자가진단 Test 13/14와 겹침).

---

## B. 신호 전파: 오류와 취소

### B1. CMD_ERROR에 사유 채널 부여

- 문제: CMD_ERROR(-1)는 사유를 나르지 못한다.
  - `m_lastRuntimeError`는 엔진 경로만 설정
  - 핸들러는 자체 팝업 후 -1만 반환
  - `FindBlockPair` 실패 등 다수 경로는 메시지 없이 -1 반환
    (예: `clsCommandsFlow.cls:35-37,197,273`)
  - 결과: RunRPA 최종 알림과 `cmd_error` jsonl 레코드에 "왜"가 빠짐
- 조치: 핸들러 실패 시 사유 기록 단일 함수(예: `modContext.SetLastError`)를
  계약에 추가하고 RunRPA가 알림/로그에 합류. vbaflow test 실패 판독 품질도 상승.

### B2. 취소 신호의 중앙화

- 문제:
  - ESC 폴링이 6곳 이상 분산 복제 (`BetterWait`, `WaitIfPaused`,
    `Action_InputText`, `Action_ClickSync`, `Action_Wait` 등)
  - `BetterWait`는 ESC 감지 시 그냥 Exit Sub -> 엔진은 액션이
    완료됐는지 취소됐는지 구분 불가
  - 일시정지 정책(팝업+폴링)이 L3에 존재 (`modActions.bas:24-44` WaitIfPaused)
- 조치: modContext에 `AbortRequested` 플래그 도입.
  폴링 헬퍼 한 곳만 키를 읽어 플래그를 세우고, 엔진 루프(CheckControlSignals)와
  각 대기 루프는 플래그만 확인. 일시정지 정책은 엔진 쪽으로 회수.
- 부수: 50ms 미만 대기와 `WaitIfPaused`의 Sleep 500/200 구간은 ESC 미감지 사각지대.

---

## C. SSOT 완성: 레지스트리 지식의 런타임 강제

### C1. ParamSpec 런타임 강제

- 문제: `Required`/`DataType`은 modDevTools 포맷터(`modDevTools.bas:249,281,304`)와
  modScriptAssist 메모(`modScriptAssist.bas:600`)만 소비.
  디스패치(ExecuteSmartAction의 DecodeSlot)는 전혀 검증하지 않아
  빈 필수 파라미터/비숫자 좌표가 핸들러까지 내려가 제각각 실패한다.
- 조치: 디코드 게이트에서 Required/DT_NUM 검사 후 균일한 에러
  ("라인 X: P1(x)은 숫자여야 합니다") 반환. 핸들러별 방어 코드 제거 가능.

### C2. modDevTools의 시맨틱 하드코딩 회수

- 문제: `modDevTools.bas:328-338`의 `IsConditionOperatorSlot` /
  `IsMemoryLValueSlot` / `IsJumpTargetSlot`이 IF·WHILE P2, SET·INC·DEC P1,
  GOTO·CALL P3를 직접 하드코딩. 이는 `ParamSpec.IsRaw` 등 레지스트리가
  이미 아는 지식의 두 번째 사본이며, 명령어 추가 시 양쪽 수정을 강제하는 드리프트 구조.
- 조치: CmdDef/ParamSpec 질의로 대체.
- 비고: Test_RegistryDispatchSync는 레지스트리-핸들러만 역검증하며
  레지스트리-DevTools 정합은 아무도 검증하지 않는다.

### C3. 정적 검증 공백 메우기

- 문제 (clsScriptProgramAnalysis):
  - ValidateStructure는 IF 계열과 SELECT 짝만 검증.
    WHILE/REPEAT/FOREACH_ROW 짝 불균형은 통과 -> 런타임에서야 CMD_ERROR
  - CALL이 GOTO 전용 라벨을 가리켜도 통과 (역할 경고만, 오류 아님)
  - 고아 CASE/END_SELECT 미검증
- 조치: IF와 동일한 스택 스캔 한 벌로 루프 짝 검증 추가.
  CALL 대상 적합성과 고아 CASE도 같은 패스에서 처리.

---

## D. 구조 정리: 중복 통합과 죽은 코드

- Levenshtein 3벌: `modEngine.bas:547`, `clsScriptProgramAnalysis.cls:567`,
  `modScriptAssist.bas:622` -> modHelper로 1벌 통합
- 블록 짝 스캔 4벌 (Index.FindBlockPair / BuildIfEndIndex / BuildSelectBlock /
  Analysis.BuildIfStructureAnalysis), 서브루틴 종료행 탐색 2벌
  (`clsScriptProgramIndex.cls:426-451` vs `clsScriptProgramAnalysis.cls:308-333`),
  깊이 모델 2종(m_loopDepths vs m_displayDepths), 점프타깃 검사 2벌
  -> Index를 스캔의 단일 소유자로, Analysis는 소비만
- 검증 시 3중 재빌드: ValidateStructure/ValidateLabels가 로드 직후 만든 인덱스를
  매번 재구축 (`clsScriptProgramAnalysis.cls:116-117,133-135`)
  -> 빌드 완료 플래그. 루프 좌표의 O(N^2) 잠재(라인마다 FindBlockPair 재스캔)도 정리
- 죽은 코드 제거:
  - `modContext.IsRunning`/`CheckRunningState` (호출자 없음)
  - `modActions.Action_WaitForColor` (호출처 없음, 구버전 호환 주석)
  - `Action_CheckCondition`/`IsConditionOperator` shim (modCondition 직접 사용으로 대체됨)
  - clsTableSnapshot 미사용 인덱싱 표면 (GetIndex/SelectRows/GetRecord 등,
    테스트만 사용 -> 유지 시 "테스트 지원" 명시)
  - Index의 CMD_ERROR 별칭 상수 4개 (LINE_NOT_FOUND 등 동일값)
- modEngine 분할: 740줄에 인터프리터 + 토큰 해석기 + 문자열 거리 + 점프 서비스 혼재
  -> ResolveParam/ResolveToken 계열을 modTokenResolver로 추출 (토큰 해석 단독 테스트 가능)
- 기타 소소: `modWinAPI.MoveFormToWorkAreaTopLeftTopMost`(L4에 폼 개념 침투),
  hex 포맷 관용구 중복(modHelper가 GetColorHexAt 미사용), RGB 매직값 상수화,
  `clsScriptProgramSource.ValidateLineRange` 오류 컨텍스트 클래스명 오기

---

## E. 성능 (선택적, 실측 후 판단)

- `clsTableSnapshot.SnapshotVisibleText`(113-148): 셀 단위 .Text/.Value/.NumberFormat
  COM 호출. FOREACH 소스가 수천 행이면 지배 비용.
  -> .Value2 일괄 읽기 + 오버플로("###") 의심 셀만 .Text 보정
- RunRPA 매 줄 폼 업데이트 + 상태바 + DoEvents x2 (`modEngine.bas:116-118,152`):
  SET/IF 위주 데이터 루프에서 병목 -> 시간 기반 스로틀(예: 100ms당 1회)

---

## 실행 순서

1. **A1~A4**: 헤드리스 계약. vbaflow test가 실제로 돌기 위한 전제.
   vbaflow 벤더링(self-contained LogShim 교체)과 같은 커밋 흐름에서 처리
2. **C1**: 디코드 게이트 검증. 파급 대비 효과 최대
3. **B1, B2**: 신호 전파. 이후 모든 디버깅 품질의 토대
4. **C2, C3**: 시맨틱 사본 회수 + 검증 공백
5. **D**: 위 작업에 얹어 점진 정리
6. **E**: 실측 후 판단
