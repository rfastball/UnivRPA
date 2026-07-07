# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# AGENTS.md / CLAUDE.md - UnivRPA

Excel 시트를 IDE처럼 쓰는 저코드 RPA 프레임워크. 스크립트는 6열 명령어 테이블
(A=ROW, B=COMMAND, C~E=PARAM1~3, F=COMMENT)이며, 5개 카테고리 36개 명령어를
줄 단위로 해석 실행한다. 사용자 매뉴얼: `src/docs/reference/` (01~03).

(시그니처 스키마·src 편집 주의는 상위 `../CLAUDE.md` 참조)

대상 워크북: **UnivRPA.xlsm** (xlsb에서 xlsm으로 마이그레이션됨)

테스트: `Public Sub Test_*()`(인자 없음). 단언은 `LogShim.AssertEq`.
`modTester.Run_Flow_Diagnosis`(헤드리스 통합 자가진단)와 `Run_All`(I/O 포함
반수동 진단)은 수동 실행용 스위트다.

## 아키텍처

### 실행 파이프라인

`RunRPA`(modEngine, Ctrl+Shift+W) 기준 한 사이클:

1. ActiveSheet 스냅샷 → `clsScriptProgram.LoadFromRange` 로 ProgramImage 승격
2. `ValidateStructure`/`ValidateLabels` 실행 전 정적 검증 (실패 시 실행 거부)
3. 줄 루프: `ExecuteSmartAction(줄번호)` →
   - `modCmdRegistry.GetCmdDef` 로 명령어 정규화·조회 (미등록이면 CMD_ERROR)
   - decode: `ParamSpec.ResolveMode` 에 따라 슬롯별 RAW(원본 유지) / RESOLVE(토큰 치환)
     - RAW 슬롯은 단 둘: lvalue 주소(SET/INC/DEC P1), 조건 연산자(IF/ELSE_IF/WHILE P2)
   - 2단 디스패치: Category → `ICommand` 핸들러 클래스 → 클래스 내부 명령어 분기
4. 핸들러 반환값이 다음 줄번호: `CMD_NEXT_LINE`(0)=다음 줄, 양수=점프,
   `CMD_ERROR`(-1)=중단, `CMD_HALT`(-2)=의도된 종료

실행 중 제어: Pause=일시정지, Esc=중단 (`modActions.BetterWait` 가 폴링).

### 레이어 규율 (위가 아래만 호출)

| Layer | 모듈 | 역할 |
|---|---|---|
| 0 | `modCmdRegistry` + `CmdDef`/`ParamSpec` | 명령어 지식의 SSOT. 조회는 `GetCmdDef` 단일 게이트만 허용, `m_reg` 직접 접근 금지 |
| 1 | `modEngine` | 인터프리터. API/변수/동작 직접 수행 금지 (아래 레이어에 위임) |
| 2 | `modContext` | 변수(Memory)·콜스택·실행 상태. Get/Set 함수 경유만 |
| 3 | `modActions` | 구체 동작(클릭·입력·대기). 판단 로직 최소화 |
| 4 | `modWinAPI` | Win32 래퍼(64비트 PtrSafe). 무상태, 루프/대기 로직 금지 |

명령어 핸들러 5클래스(`clsCommandsFlow/Action/Memory/System/Vision`)는 모두
`ICommand` 를 구현하며 `modCmdRegistry` 가 카테고리별로 1회 배선한다.
조건 연산자 비교 로직의 SSOT는 `modCondition`.

### ProgramImage (정적 모델)

`clsScriptProgram` 은 파사드. 외부 모듈은 이 클래스에만 질의한다.
내부 분할: `clsScriptProgramSource`(원본/행/명령어 접근, `clsInstruction` 캐시),
`clsScriptProgramIndex`(블록 짝·라벨·좌표 인덱스, `clsSelectBlock`),
`clsScriptProgramAnalysis`(흐름·변수·검증 분석).

FOREACH_ROW 데이터 순회는 `modTableSource` + `clsTableSnapshot`
(시트를 2D 배열로 메모리화한 읽기 전용 컨테이너).

### 작성 보조 (런타임과 분리)

- `modScriptAssist`: 시트 이벤트·상태바·메모 UX. 구조/라벨/흐름 판단을 직접
  계산하지 않고 clsScriptProgram 에 질의
- `modDevTools`: 시트 어댑터·포맷터·색상 서식 (Ctrl+Shift+F 등)
- `frmRPAStatus`: 실행 상태 폼, `frmCapture`: 색상·좌표 캡처 HUD (Ctrl+Shift+E)
- 스크립트 시트 판별 시그니처: A1=`ROW` 그리고 B1=`COMMAND`

## 명령어 추가 절차

1. `modCmdRegistry.BuildRegistry` 에 `Reg` 한 줄 추가 (카테고리·블록역할·색상·ParamSpec)
2. 해당 카테고리의 `clsCommands*` 핸들러 `Execute` 에 분기 구현
3. `modTester.Test_RegistryDispatchSync` 가 레지스트리↔핸들러 라우팅을 전수
   역검증하므로 둘 중 하나만 고치면 잡힌다
4. 블록 명령어(STARTER/MIDBLOCK/ENDER)면 clsScriptProgramIndex 의 짝 인덱싱과
   검증 규칙 영향 확인

## 기타

- `archive/` 는 과거 스냅샷 보관용. 편집 대상은 항상 `src/`
- `spec.md`/`CODEBASE_MAP.md` 는 아직 없음 (상위 규약상 위치는 `src/` 루트)

<!-- vbaflow:managed:start -->
## 개발 루프 (_VbaFlow)

빌드/테스트 루프는 `vbaflow`(C# .NET 8 단일 exe)로 수행한다. PATH 미등록 시 exe 경로:
`..\_VbaFlow\src\VbaFlow\bin\Release\net8.0-windows\win-x64\publish\vbaflow.exe`
온보딩/드리프트 복구: `vbaflow init` (검사만: `vbaflow init --check`).
프로젝트 고유 내용은 이 관리 블록 밖에 작성한다 (init이 블록 내부를 덮어씀).

1. 편집: src\*.bas/.cls/.frm 수정
2. 세션: vbaflow session start (도구 소유 Excel, 멱등)
3. 주입: vbaflow push (src -> 워크북 in-place 주입, --prune 지원)
4. 실행: vbaflow test (Test_* 전체 실행 + LogShim jsonl 판독, --timeout <초>)
5. 수정: fail이면 원인 고치고 1로, green이면 git commit

- 모든 명령 --json 지원: {command, ok, exitCode, hint, errors[], payload}. 성공만 exit 0
- 테스트는 Public Sub Test_*() (인자 없음), 단언은 LogShim.AssertEq (Debug.Assert 금지)
- src\LogShim.bas는 벤더링 사본 - 직접 수정 금지, 갱신은 vbaflow init
- pull은 src의 미푸시 편집을 덮어쓴다 - 워크북 쪽이 진실일 때만
- src 규약: UTF-8 BOM 없음 + CRLF + 끝 개행 1개. CP949 밖 문자(em dash U+2014, 이모지)는 push가 거부
<!-- vbaflow:managed:end -->
