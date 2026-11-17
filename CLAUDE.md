# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# AGENTS.md / CLAUDE.md - UnivRPA

Excel 시트를 IDE처럼 쓰는 저코드 RPA 프레임워크. 스크립트는 6열 명령어 테이블
(A=ROW, B=COMMAND, C~E=PARAM1~3, F=COMMENT)이며, 5개 카테고리 36개 명령어를
줄 단위로 해석 실행한다. 사용자 매뉴얼: `src/docs/reference/` (01~03).

(시그니처 스키마·src 편집 주의는 상위 `../CLAUDE.md` 참조)

대상 워크북: **UnivRPA.xlsm** (xlsb에서 xlsm으로 마이그레이션됨)

테스트: `Public Sub Test_*()`(인자 없음)를 `vbaloop test`가 열거 실행한다.
단언은 `LogShim.AssertEq`(self-contained, jsonl 직접 기록). 스위트 본체는
`modTester`의 Private `TC_*` 함수이고 `Test_*`는 그 래퍼다.
`modTester.Run_Flow_Diagnosis`(헤드리스 통합 자가진단)와 `Run_All`(I/O 포함
반수동 진단)은 수동 실행용 스위트다. 무인 실행 알림 억제는
`modActions.SetUnattendedMode` (테스트 하니스가 자동 설정).

개발 루프: `vbaloop push` → `vbaloop test` (세션은 자동 확보, 루트 `vbaloop.toml` 참조,
상세는 아래 vbaloop 관리 블록).

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
## Development loop (vbaloop)

The build/test loop runs through `vbaloop` (a single exe). If it is not on PATH, the install path is
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (if even that is missing, reinstall with `vbaloop setup`).
Onboarding/drift repair: `vbaloop init` (check only: `vbaloop init --check`).
If init --check exits 3 with `exe-older` (the managed_by version in vbaloop.toml is newer than the running exe), do not fix it
with init; ask the user to run `vbaloop update` (do not edit or delete the managed_by line).
Write project-specific content outside this managed block (init overwrites the inside of the block).

1. Edit: change src\*.bas/.cls/.frm
   New UserForm: write JSON and plain VBA outside src, then
   `vbaloop form build --spec forms/MyForm.json --code forms/MyForm.vba --json` -> push.
   No template or hand-written FRX is needed. Before form work, read
   `.claude/skills/vbaloop/references/forms.md` for design, code rules and schema.
   Rebuild with --force only to replace the generated pair from those inputs.
2. Session: ephemeral by default - push/pull/test/run/diff start a hidden Excel per command and close it (vbaloop.toml
   [excel] session). If a session opened with session start/attach is alive, it is used
3. Inject: vbaloop push (src -> in-place injection into the workbook -> saved only if the VBE compile passes; supports --prune/--force.
   Backup to _dbg\backups\ right before injecting)
4. Run: vbaloop test (runs every Test_* + reads the LogShim jsonl; waits until done, no default time limit -
   run long ones in the background and stop the command to cancel at the next LogShim call)
5. Fix: on fail, fix the cause and go back to 1; when green you are done (commit now if you use version control)
6. Release: when the task is finished (green and reported, or you stop working on this project), run `vbaloop lease release`.
   Commands take the project lease automatically and nothing else frees it while your session stays open,
   so skipping this blocks the next agent for up to 20 minutes. If you opened a session with session start, session stop releases it instead

- Every command supports --json: {command, ok, exitCode, hint, errors[], payload}
- exit code: success 0 / environment error 3 (session/Excel/COM/busy/quarantined - follow the hint) / otherwise 2
- Quarantined (hint mentions recovery clear / --discard): the tool could not safely clean up Excel, so it stopped instead of killing it.
  Check vbaloop recovery status --json against the current task's launch record (PID + start time). If it is the tool-owned
  Excel you started for isolated tests/measurements, all unsaved changes are disposable task output, and it was never handed
  to the user or used for unrelated workbooks, run vbaloop session stop --discard --json without extra confirmation and continue.
  Busy/unresponsive alone does not revoke that evidence; owner: tool or a hidden window alone is not enough.
  Otherwise ask the user (honor prior discard authorization); if they close Excel, vbaloop recovery clear;
  if they authorize discarding it, vbaloop session stop --discard. Attached user Excel and Office recovery Excel are excluded.
  Never bypass a refusal with direct process kill, marker deletion, or doctor --kill-orphans; orphan cleanup still needs user approval
- busy (exit 3): Excel is in cell edit mode or showing a dialog - retry the same command shortly
- If the user already has the canonical workbook open in Excel (read-only refusal hint): after the user confirms,
  vbaloop session attach - push/test then work in that Excel and session stop only detaches
- Tests are Public Sub Test_*() (no arguments); assert with LogShim.AssertEq (no Debug.Assert)
- src\LogShim.bas is a vendored copy - do not edit directly; update with vbaloop init
- When push fails, the hint states the workbook state (untouched / fully restored / discarded / session poisoned - not saved).
  After a failed push the canonical workbook on disk is always in its pre-push state.
  On a compile error (payload.compile.result = compile_error), fix it and push again.
  Exception: if the error is pre-existing and unrelated to this change, ask the user; only after they confirm, push --skip-compile.
  If you fixed something but the same error repeats and push keeps returning cached:true, push --force bypasses the fingerprint cache.
  To revert the workbook file itself, vbaloop rollback --latest (list: rollback --list;
  exit 3 while a persistent session is alive - session stop first)
- In ephemeral mode test/run open the canonical workbook read-only - sheets changed by tests are not saved
- Distribution artifact: vbaloop dist --out <path> (clean copy build - filter/stamp/sanitize/VBE compile check/optional bootstrap).
  push must come first (refused when last push != src). Procedure/failure playbook: SKILL.md
- pull is blocked when there are unsynced src edits - by default apply them with push first;
  only when the workbook is the truth, overwrite with pull --force
- No direct COM access to the session Excel (GetActiveObject, New-Object -ComObject etc. - including reads/diagnostics).
  To see what is really in the workbook, vbaloop diff --json (read-only; --module <name> dumps a single module)
- src conventions: UTF-8 without BOM + CRLF + exactly one trailing newline. Characters outside CP949 (em dash U+2014, emoji) are refused by push
- If you find a defect in vbaloop itself: create a standard report with vbaloop bugreport "<title>" and fill it in
  (procedure: .claude\skills\vbaloop\SKILL.md)
- Detailed procedures/failure playbooks: .claude\skills\vbaloop\SKILL.md (managed by vbaloop init - do not edit directly)
<!-- vbaloop:managed:end -->
