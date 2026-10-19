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
     - RAW 슬롯 4종은 레지스트리가 ParamSpec.RawKind로 선언: 조건 연산자(IF/ELSE_IF/WHILE P2),
       lvalue 주소(SET/INC/DEC P1), 점프대상(GOTO/CALL P3), 일반 텍스트(LABEL P3)
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

<!-- vbaloop:managed:start -->
## 개발 루프 (vbaloop)

빌드/테스트 루프는 `vbaloop`(단일 exe)로 수행한다. PATH 미등록 시 설치 경로:
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (그마저 없으면 `vbaloop setup`으로 재설치).
온보딩/드리프트 복구: `vbaloop init` (검사만: `vbaloop init --check`).
단, init --check가 exit 3 + `exe-older`(vbaloop.toml managed_by > 실행 exe)면 init으로 고치지 말고
사용자에게 `vbaloop update`를 요청한다 (managed_by 줄 수정/삭제 금지).
프로젝트 고유 내용은 이 관리 블록 밖에 작성한다 (init이 블록 내부를 덮어씀).

1. 편집: src\*.bas/.cls/.frm 수정
2. 세션: 자동 확보(push/test가 스스로 띄움, 수동은 vbaloop session start)
3. 주입: vbaloop push (src -> 워크북 in-place 주입, --prune/--force 지원. 주입 직전 _dbg\backups\에 백업)
4. 실행: vbaloop test (Test_* 전체 실행 + LogShim jsonl 판독, --timeout <초>)
5. 수정: fail이면 원인 고치고 1로, green이면 완료(버전관리를 쓰면 이때 커밋)

- 모든 명령 --json 지원: {command, ok, exitCode, hint, errors[], payload}
- exit code: 성공 0 / 환경 오류 3 (세션/Excel/COM/busy/격리 - hint를 따른다) / 그 외 2
- 격리(hint에 recovery clear / --discard): 도구가 Excel을 안전하게 정리할 수 없어 kill 대신 멈춘 상태.
  스스로 Excel을 죽이지 말 것(--discard, doctor --kill-orphans 포함) - 사용자에게 알리고, 사용자가
  그 Excel을 닫으면 vbaloop recovery clear, 버려도 된다고 하면 vbaloop session stop --discard
- busy(exit 3): Excel이 셀 편집/대화상자 중 - 잠시 후 같은 명령 재시도
- 사용자가 정본을 이미 Excel에서 열어둔 경우(read-only 거부 hint): 사용자 확인 후
  vbaloop session attach - 그 Excel에서 push/test가 동작하고 session stop은 분리만 한다
- 테스트는 Public Sub Test_*() (인자 없음), 단언은 LogShim.AssertEq (Debug.Assert 금지)
- src\LogShim.bas는 벤더링 사본 - 직접 수정 금지, 갱신은 vbaloop init
- push 실패 시 hint가 워크북 상태를 말한다(변경 없음 / 전부 원복 / 세션 poison - 저장 안 함).
  디스크 정본은 항상 push 이전 상태. 워크북 파일 자체를 되돌리려면 session stop 후
  vbaloop rollback --latest (목록: rollback --list, 세션이 살아 있으면 exit 3)
- 배포 아티팩트: vbaloop dist --out <경로> (클린 사본 빌드 - 필터/스탬프/무해화/VBE 컴파일 검증/선택 부트스트랩).
  push 선행 필수(마지막 push == src 미동기 시 거부). 절차/실패 플레이북은 SKILL.md
- pull은 동기화 안 된 src 편집이 있으면 차단된다 - 기본은 push로 먼저 반영,
  워크북 쪽이 진실일 때만 pull --force로 덮어쓰기
- 세션 Excel에 직접 COM 접근 금지(GetActiveObject, New-Object -ComObject 등 - 조회/진단 포함).
  워크북 실제 내용 확인은 vbaloop diff --json (읽기 전용, --module <이름>은 단일 모듈 덤프)
- src 규약: UTF-8 BOM 없음 + CRLF + 끝 개행 1개. CP949 밖 문자(em dash U+2014, 이모지)는 push가 거부
- vbaloop 자체 결함 발견 시: vbaloop bugreport "<제목>" 으로 표준 리포트 생성 후 내용 채움
  (절차: .claude\skills\vbaloop\SKILL.md)
- 심화 절차/실패 플레이북: .claude\skills\vbaloop\SKILL.md (vbaloop init이 관리 - 직접 수정 금지)
<!-- vbaloop:managed:end -->
